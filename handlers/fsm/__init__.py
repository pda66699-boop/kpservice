from __future__ import annotations

import asyncio
import re
import unicodedata
from pathlib import Path
from typing import Any
from uuid import uuid4

from aiogram import F, Router
from aiogram.filters import Command
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.types import FSInputFile, KeyboardButton, Message, ReplyKeyboardMarkup
from pypdf import PdfReader, PdfWriter

from services.kp_builder import build_kp_pdf as build_kp

router = Router(name="fsm_root")

MANAGER_PROFILES: dict[int, dict[str, str]] = {}


class KpBuildStates(StatesGroup):
    template_docx = State()
    client_name = State()
    kp_number = State()
    excel_file = State()
    drawings_rtf = State()


class ManagerProfileStates(StatesGroup):
    choose_field = State()
    full_name = State()
    phone = State()


class ProposalReviewStates(StatesGroup):
    review = State()
    choose_fix = State()
    replace_template = State()
    replace_client_name = State()
    replace_kp_number = State()
    replace_excel = State()
    replace_drawings = State()
    delete_page = State()


def has_manager_profile(user_id: int | None) -> bool:
    return bool(user_id and user_id in MANAGER_PROFILES)


def get_generation_keyboard() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="📤 Загрузить шаблон")],
            [KeyboardButton(text="🏛️ Порталы"), KeyboardButton(text="🧱 Фасады")],
            [KeyboardButton(text="🌿 Зимние сады и зенитные фонари")],
            [KeyboardButton(text="🏡 Коттеджи (комплексное решение)")],
        ],
        resize_keyboard=True,
    )


def get_flow_keyboard() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="◀️ Назад"), KeyboardButton(text="📖 Инструкция")],
            [KeyboardButton(text="👤 Данные менеджера")],
            [KeyboardButton(text="🚀 Сгенерировать КП заново")],
        ],
        resize_keyboard=True,
    )


def _main_menu_keyboard() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="🚀 Начать генерацию КП")],
            [KeyboardButton(text="👤 Данные менеджера")],
            [KeyboardButton(text="📖 Инструкция")],
        ],
        resize_keyboard=True,
    )


def _manager_edit_keyboard() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="👤 ФИО"), KeyboardButton(text="📞 Телефон")],
            [KeyboardButton(text="❌ Отмена")],
        ],
        resize_keyboard=True,
    )


def _review_keyboard() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="✅ Да"), KeyboardButton(text="🛠️ Исправить")],
        ],
        resize_keyboard=True,
    )


def _fix_keyboard() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="🧩 Шаблон КП")],
            [KeyboardButton(text="👤 ФИО клиента"), KeyboardButton(text="🔢 Номер КП")],
            [KeyboardButton(text="📊 Заменить расчет"), KeyboardButton(text="📐 Заменить чертежи")],
            [KeyboardButton(text="🗑️ Удалить страницу")],
            [KeyboardButton(text="❌ Отмена")],
        ],
        resize_keyboard=True,
    )


def _safe_file_name(name: str, fallback: str) -> str:
    raw = name.strip() if name else fallback
    return re.sub(r"[^A-Za-z0-9._-]", "_", raw)


def _safe_output_filename(name: str, fallback: str) -> str:
    raw = name.strip() if name else fallback
    raw = unicodedata.normalize("NFC", raw)
    # Collapse all common Unicode spaces and remove zero-width separators.
    raw = re.sub(r"[\u200B\u200C\u200D\u2060\uFEFF]", "", raw)
    raw = re.sub(r"[\s\u00A0\u1680\u2000-\u200A\u2028\u2029\u202F\u205F\u3000]+", " ", raw)
    # Keep Cyrillic and spaces, replace only filesystem-forbidden characters.
    sanitized = re.sub(r'[<>:"/\\|?*\x00-\x1F]', "_", raw).strip(" .")
    sanitized = re.sub(r"\s{2,}", " ", sanitized)
    return sanitized or fallback


def _build_kp_filename(kp_number: str, client_name: str) -> str:
    kp_number_value = kp_number or "без номера"
    client_name_value = client_name or "клиента"
    base = f"КП № {kp_number_value} для {client_name_value}"
    return f"{_safe_output_filename(base, 'КП для клиента')}.pdf"


def _normalize_pdf_filename(filename: str) -> str:
    raw = (filename or "").strip()
    if not raw:
        return _build_kp_filename("", "")
    stem = Path(raw).stem
    suffix = Path(raw).suffix.lower()
    if suffix != ".pdf":
        suffix = ".pdf"
    return f"{_safe_output_filename(stem, 'КП для клиента')}{suffix}"


def _format_manager_phone(value: str) -> str | None:
    digits = re.sub(r"\D", "", value)
    if len(digits) == 10:
        digits = "8" + digits
    elif len(digits) == 11 and digits[0] in {"7", "8"}:
        digits = "8" + digits[1:]
    else:
        return None
    return f"{digits[0]} {digits[1:4]} {digits[4:7]}-{digits[7:9]}-{digits[9:11]}"


async def _download_document(message: Message, target_path: Path) -> None:
    if not message.document:
        raise ValueError("No document in message")
    file = await message.bot.get_file(message.document.file_id)
    target_path.parent.mkdir(parents=True, exist_ok=True)
    await message.bot.download_file(file.file_path, destination=str(target_path))


async def _rebuild_kp_from_state(message: Message, state: FSMContext) -> Path:
    fsm_data = await state.get_data()
    user_id = message.from_user.id if message.from_user else None
    template_path = fsm_data.get("template_path")
    excel_path = fsm_data.get("excel_path")
    output_dir = fsm_data.get("output_dir", "artifacts")
    drawings_rtf_path = fsm_data.get("drawings_rtf_path")
    kp_number = str(fsm_data.get("kp_number", "")).strip()
    client_name = str(fsm_data.get("client_name", "")).strip()
    kp_filename = _normalize_pdf_filename(_build_kp_filename(kp_number, client_name))
    manager_phone = (
        MANAGER_PROFILES.get(user_id, {}).get("manager_phone")
        if user_id is not None
        else None
    )
    manager_name = (
        MANAGER_PROFILES.get(user_id, {}).get("manager_name")
        if user_id is not None
        else None
    )

    if not template_path or not excel_path:
        raise ValueError("Недостаточно данных для сборки КП")

    render_data: dict[str, Any] = {
        "client_name": client_name,
        "kp_number": kp_number,
        "manager_name": manager_name or "",
        "manager_phone": manager_phone or "",
    }

    final_pdf_path = await asyncio.to_thread(
        build_kp,
        template_path=template_path,
        data=render_data,
        excel_path=excel_path,
        output_dir=output_dir,
        kp_filename=kp_filename,
        drawings_rtf_path=drawings_rtf_path,
    )
    await state.update_data(kp_filename=kp_filename, last_final_pdf=str(final_pdf_path))
    return Path(final_pdf_path)


async def _send_kp_and_review(message: Message, state: FSMContext, final_pdf_path: Path) -> None:
    await message.answer_document(FSInputFile(str(final_pdf_path)))
    await state.set_state(ProposalReviewStates.review)
    await message.answer(
        "✅ КП сформировано корректно?",
        reply_markup=_review_keyboard(),
    )


@router.message(Command("create_kp"))
async def start_kp_build(message: Message, state: FSMContext) -> None:
    user_id = message.from_user.id if message.from_user else None
    if not has_manager_profile(user_id):
        await state.clear()
        await state.set_state(ManagerProfileStates.full_name)
        await message.answer(
            "Сначала необходимо добавить данные менеджера — они сохранятся в памяти и будут использоваться при формировании КП.\n"
            "В любой момент ты можешь изменить свои данные через кнопку «👤 Данные менеджера»."
        )
        await message.answer("Укажи свои ФИО (достаточно фамилии и имени).")
        return

    await state.clear()
    await state.set_state(KpBuildStates.template_docx)
    await message.answer(
        "🚀 Запускаем генерацию КП.\n"
        "Нажми «📤 Загрузить шаблон», чтобы перейти к шагу 1.",
        reply_markup=get_generation_keyboard(),
    )


@router.message(F.text.casefold() == "📤 загрузить шаблон".casefold())
async def start_from_upload_button(message: Message, state: FSMContext) -> None:
    user_id = message.from_user.id if message.from_user else None
    if not has_manager_profile(user_id):
        await state.clear()
        await state.set_state(ManagerProfileStates.full_name)
        await message.answer(
            "Сначала заполни данные менеджера.\n"
            "Укажи свои ФИО (достаточно фамилии и имени)."
        )
        return

    await state.set_state(KpBuildStates.template_docx)
    await message.answer("🧩 Шаг 1. Пришли шаблон КП, соответствующий запросу клиента (формат .docx)")


@router.message(F.text.casefold() == "создать кп")
async def start_kp_build_from_text(message: Message, state: FSMContext) -> None:
    await start_kp_build(message, state)


@router.message(F.text.casefold() == "🚀 начать генерацию кп")
async def start_kp_build_from_button(message: Message, state: FSMContext) -> None:
    await start_kp_build(message, state)


@router.message(F.text.casefold() == "🚀 сгенерировать кп заново".casefold())
async def restart_kp_build(message: Message, state: FSMContext) -> None:
    await state.clear()
    await state.set_state(KpBuildStates.template_docx)
    await message.answer(
        "🔄 Начинаем заново.\n"
        "Нажми «📤 Загрузить шаблон»",
        reply_markup=get_generation_keyboard(),
    )


@router.message(F.text.casefold() == "👤 данные менеджера".casefold())
@router.message(F.text.casefold() == "данные менеджера")
async def edit_manager_profile(message: Message, state: FSMContext) -> None:
    user_id = message.from_user.id if message.from_user else None
    profile = MANAGER_PROFILES.get(user_id) if user_id is not None else None
    if not profile:
        await state.clear()
        await state.set_state(ManagerProfileStates.full_name)
        await message.answer("🧑‍💼 Сначала укажи свои ФИО (достаточно фамилии и имени).")
        return

    await state.clear()
    await state.set_state(ManagerProfileStates.choose_field)
    await message.answer(
        "🧑‍💼 Сейчас сохранены такие данные:\n"
        f"ФИО: {profile.get('manager_name', '—')}\n"
        f"Номер телефона: {profile.get('manager_phone', '—')}\n\n"
        "Какую информацию будем менять?",
        reply_markup=_manager_edit_keyboard(),
    )


@router.message(ManagerProfileStates.choose_field, F.text.casefold() == "👤 фио")
@router.message(ManagerProfileStates.choose_field, F.text.casefold() == "фио")
async def manager_choose_name(message: Message, state: FSMContext) -> None:
    await state.update_data(edit_mode="name_only")
    await state.set_state(ManagerProfileStates.full_name)
    await message.answer("✍️ Введи новое ФИО.")


@router.message(ManagerProfileStates.choose_field, F.text.casefold() == "📞 телефон")
@router.message(ManagerProfileStates.choose_field, F.text.casefold() == "телефон")
async def manager_choose_phone(message: Message, state: FSMContext) -> None:
    await state.update_data(edit_mode="phone_only")
    await state.set_state(ManagerProfileStates.phone)
    await message.answer("📞 Введи новый номер телефона.")


@router.message(ManagerProfileStates.choose_field, F.text.casefold() == "❌ отмена")
@router.message(ManagerProfileStates.choose_field, F.text.casefold() == "отмена")
async def manager_edit_cancel(message: Message, state: FSMContext) -> None:
    await state.clear()
    await message.answer("✅ Отменено. Открываю основное меню.", reply_markup=_main_menu_keyboard())


@router.message(Command("cancel"))
async def cancel_kp_build(message: Message, state: FSMContext) -> None:
    await state.clear()
    await message.answer("🛑 Мастер создания КП отменён.")


@router.message(ManagerProfileStates.full_name, F.text)
async def receive_manager_name(message: Message, state: FSMContext) -> None:
    manager_name = (message.text or "").strip()
    if len(manager_name) < 3:
        await message.answer("Укажи ФИО корректно (минимум 3 символа).")
        return

    user_id = message.from_user.id if message.from_user else None
    if user_id is None:
        await message.answer("Не удалось определить пользователя. Повтори попытку.")
        return

    data = await state.get_data()
    edit_mode = data.get("edit_mode")
    if edit_mode == "name_only":
        profile = MANAGER_PROFILES.get(user_id, {})
        profile["manager_name"] = manager_name
        MANAGER_PROFILES[user_id] = profile
        await state.clear()
        await message.answer(
            f"✅ ФИО обновлено: {manager_name}",
            reply_markup=_main_menu_keyboard(),
        )
        return

    await state.update_data(manager_name=manager_name, edit_mode=None)
    await state.set_state(ManagerProfileStates.phone)
    await message.answer("📞 Укажи свой контактный номер телефона для указания в КП")


@router.message(ManagerProfileStates.phone, F.text)
async def receive_manager_phone_profile(message: Message, state: FSMContext) -> None:
    raw_phone = (message.text or "").strip()
    formatted = _format_manager_phone(raw_phone)
    if formatted is None:
        await message.answer("Некорректный телефон. Пример формата: 8 900 123-45-67")
        return

    user_id = message.from_user.id if message.from_user else None
    if user_id is None:
        await message.answer("Не удалось определить пользователя. Повтори попытку.")
        return

    data = await state.get_data()
    edit_mode = data.get("edit_mode")
    manager_name = str(data.get("manager_name", "")).strip()
    if not manager_name:
        manager_name = MANAGER_PROFILES.get(user_id, {}).get("manager_name", "")
    MANAGER_PROFILES[user_id] = {
        "manager_name": manager_name,
        "manager_phone": formatted,
    }
    await state.clear()

    if edit_mode == "phone_only":
        await message.answer(
            f"✅ Телефон обновлён: {formatted}",
            reply_markup=_main_menu_keyboard(),
        )
        return

    await message.answer(
        f"✅ Данные сохранены.\n"
        f"ФИО: {manager_name}\n"
        f"Телефон: {formatted}\n\n"
        "Теперь нажми «📤 Загрузить шаблон».",
        reply_markup=get_generation_keyboard(),
    )


@router.message(KpBuildStates.template_docx, F.text.casefold() == "📤 загрузить шаблон".casefold())
async def ask_template_upload(message: Message, state: FSMContext) -> None:
    await state.set_state(KpBuildStates.template_docx)
    await message.answer("🧩 Шаг 1. Пришли шаблон КП, соответствующий запросу клиента (формат .docx)")


@router.message(KpBuildStates.template_docx, F.document)
async def receive_template_docx(message: Message, state: FSMContext) -> None:
    document = message.document
    if document is None:
        await message.answer("Пришли файл шаблона .docx")
        return
    if not document.file_name or not document.file_name.lower().endswith(".docx"):
        await message.answer("Нужен файл формата .docx")
        return

    user_id = message.from_user.id if message.from_user else 0
    job_dir = Path("artifacts") / f"kp_{user_id}_{uuid4().hex}"
    template_path = job_dir / _safe_file_name(document.file_name, "template.docx")
    await _download_document(message, template_path)

    await state.update_data(output_dir=str(job_dir), template_path=str(template_path))
    await state.set_state(KpBuildStates.client_name)
    await message.answer(
        "👤 Шаг 2. Для кого оформляем КП?\nУкажи только ФИО в соответствии с договором или заявкой (но в родительном падеже)",
        reply_markup=get_flow_keyboard(),
    )


@router.message(KpBuildStates.client_name, F.text)
async def receive_client_name(message: Message, state: FSMContext) -> None:
    client_name = (message.text or "").strip()
    if client_name.casefold() == "◀️ назад".casefold():
        await state.set_state(KpBuildStates.template_docx)
        await message.answer(
            "🧩 Шаг 1. Пришли шаблон КП, соответствующий запросу клиента (формат .docx)",
            reply_markup=get_generation_keyboard(),
        )
        return
    if not client_name:
        await message.answer("Поле клиента не может быть пустым.")
        return
    await state.update_data(client_name=client_name)
    await state.set_state(KpBuildStates.kp_number)
    await message.answer("🔢 Шаг 3. Введи номер КП в соответствии с внутренней нумерацией", reply_markup=get_flow_keyboard())


@router.message(KpBuildStates.kp_number, F.text)
async def receive_kp_number(message: Message, state: FSMContext) -> None:
    kp_number = (message.text or "").strip()
    if kp_number.casefold() == "◀️ назад".casefold():
        await state.set_state(KpBuildStates.client_name)
        await message.answer(
            "👤 Шаг 2. Для кого оформляем КП?\nУкажи только ФИО в соответствии с договором или заявкой (но в родительном падеже)",
            reply_markup=get_flow_keyboard(),
        )
        return
    if not kp_number:
        await message.answer("Номер КП не может быть пустым.")
        return
    await state.update_data(kp_number=kp_number)
    await state.set_state(KpBuildStates.excel_file)
    await message.answer(
        "📊 Шаг 4. Пришли файл расчета стоимости (форматы .xlsx/.xls/.xlsm)",
        reply_markup=get_flow_keyboard(),
    )


@router.message(KpBuildStates.excel_file, F.document)
async def receive_excel_file(message: Message, state: FSMContext) -> None:
    document = message.document
    if document is None:
        await message.answer("Пришли Excel-файл.")
        return

    ext = Path(document.file_name or "").suffix.lower()
    if ext not in {".xlsx", ".xls", ".xlsm"}:
        await message.answer("Нужен Excel-файл (.xlsx/.xls/.xlsm)")
        return

    fsm_data = await state.get_data()
    output_dir = Path(fsm_data["output_dir"])
    excel_path = output_dir / _safe_file_name(document.file_name or "price.xlsx", "price.xlsx")
    await _download_document(message, excel_path)

    await state.update_data(excel_path=str(excel_path))
    await state.set_state(KpBuildStates.drawings_rtf)
    await message.answer("📐 Шаг 5. Пришли чертежи (в формате .rtf/.pdf)", reply_markup=get_flow_keyboard())


@router.message(KpBuildStates.client_name, F.text.casefold() == "◀️ назад".casefold())
async def back_to_template_step(message: Message, state: FSMContext) -> None:
    await state.set_state(KpBuildStates.template_docx)
    await message.answer(
        "🧩 Шаг 1. Пришли шаблон КП, соответствующий запросу клиента (формат .docx)",
        reply_markup=get_generation_keyboard(),
    )


@router.message(KpBuildStates.kp_number, F.text.casefold() == "◀️ назад".casefold())
async def back_to_client_step(message: Message, state: FSMContext) -> None:
    await state.set_state(KpBuildStates.client_name)
    await message.answer(
        "👤 Шаг 2. Для кого оформляем КП?\nУкажи только ФИО в соответствии с договором или заявкой (но в родительном падеже)",
        reply_markup=get_flow_keyboard(),
    )


@router.message(KpBuildStates.excel_file, F.text.casefold() == "◀️ назад".casefold())
async def back_to_kp_number_step(message: Message, state: FSMContext) -> None:
    await state.set_state(KpBuildStates.kp_number)
    await message.answer("🔢 Шаг 3. Введи номер КП в соответствии с внутренней нумерацией", reply_markup=get_flow_keyboard())


@router.message(KpBuildStates.drawings_rtf, F.text.casefold() == "◀️ назад".casefold())
async def back_to_excel_step(message: Message, state: FSMContext) -> None:
    await state.set_state(KpBuildStates.excel_file)
    await message.answer(
        "📊 Шаг 4. Пришли файл расчета стоимости (форматы .xlsx/.xls/.xlsm)",
        reply_markup=get_flow_keyboard(),
    )


@router.message(KpBuildStates.drawings_rtf, F.document)
async def receive_drawings_rtf(message: Message, state: FSMContext) -> None:
    document = message.document
    if document is None:
        await message.answer("Пришли RTF-файл или отправь 'Пропустить'.")
        return
    if not document.file_name:
        await message.answer("Нужен файл чертежей в формате .rtf или .pdf")
        return

    drawings_ext = Path(document.file_name).suffix.lower()
    if drawings_ext not in {".rtf", ".pdf"}:
        await message.answer("Нужен файл чертежей в формате .rtf или .pdf")
        return

    fsm_data = await state.get_data()
    output_dir = Path(fsm_data["output_dir"])
    drawings_path = output_dir / _safe_file_name(document.file_name, f"drawings{drawings_ext}")
    await _download_document(message, drawings_path)
    await state.update_data(drawings_rtf_path=str(drawings_path))
    await _finalize_kp_build(message, state)


@router.message(KpBuildStates.drawings_rtf, F.text.casefold() == "пропустить")
async def skip_drawings(message: Message, state: FSMContext) -> None:
    await state.update_data(drawings_rtf_path=None)
    await _finalize_kp_build(message, state)


async def _finalize_kp_build(message: Message, state: FSMContext) -> None:
    try:
        await message.answer("⏳ Формирую КП, это может занять до 1-2 минут...")
        final_pdf_path = await _rebuild_kp_from_state(message, state)
    except Exception as exc:
        await message.answer(
            "❌ Не удалось сформировать КП.\n"
            f"Причина: {exc.__class__.__name__}: {exc}"
        )
        await state.clear()
        return

    await _send_kp_and_review(message, state, final_pdf_path)


@router.message(ProposalReviewStates.review, F.text.casefold() == "✅ да".casefold())
@router.message(ProposalReviewStates.review, F.text.casefold() == "да")
async def review_accept(message: Message, state: FSMContext) -> None:
    await state.clear()
    await message.answer(
        "Отлично, удачи в поиске клиента для следующего КП! 🚀",
        reply_markup=_main_menu_keyboard(),
    )


@router.message(ProposalReviewStates.review, F.text.casefold() == "🛠️ исправить".casefold())
@router.message(ProposalReviewStates.review, F.text.casefold() == "исправить")
async def review_fix(message: Message, state: FSMContext) -> None:
    await state.set_state(ProposalReviewStates.choose_fix)
    await message.answer(
        "Что именно ты хочешь исправить?",
        reply_markup=_fix_keyboard(),
    )


@router.message(ProposalReviewStates.choose_fix, F.text.casefold() == "❌ отмена")
@router.message(ProposalReviewStates.choose_fix, F.text.casefold() == "отмена")
async def fix_cancel(message: Message, state: FSMContext) -> None:
    await state.set_state(ProposalReviewStates.review)
    await message.answer("Ок, без изменений. КП считаем готовым?", reply_markup=_review_keyboard())


@router.message(ProposalReviewStates.choose_fix, F.text.casefold() == "🧩 шаблон кп")
@router.message(ProposalReviewStates.choose_fix, F.text.casefold() == "шаблон кп")
async def fix_template_request(message: Message, state: FSMContext) -> None:
    await state.set_state(ProposalReviewStates.replace_template)
    await message.answer("Пришли новый шаблон КП в формате .docx")


@router.message(ProposalReviewStates.choose_fix, F.text.casefold() == "👤 фио клиента")
@router.message(ProposalReviewStates.choose_fix, F.text.casefold() == "фио клиента")
async def fix_client_name_request(message: Message, state: FSMContext) -> None:
    await state.set_state(ProposalReviewStates.replace_client_name)
    await message.answer("Введи новое ФИО клиента (в родительном падеже).")


@router.message(ProposalReviewStates.choose_fix, F.text.casefold() == "🔢 номер кп")
@router.message(ProposalReviewStates.choose_fix, F.text.casefold() == "номер кп")
async def fix_kp_number_request(message: Message, state: FSMContext) -> None:
    await state.set_state(ProposalReviewStates.replace_kp_number)
    await message.answer("Введи новый номер КП.")


@router.message(ProposalReviewStates.choose_fix, F.text.casefold() == "📊 заменить расчет")
@router.message(ProposalReviewStates.choose_fix, F.text.casefold() == "заменить расчет")
async def fix_excel_request(message: Message, state: FSMContext) -> None:
    await state.set_state(ProposalReviewStates.replace_excel)
    await message.answer("Пришли новый файл расчёта (.xlsx/.xls/.xlsm).")


@router.message(ProposalReviewStates.choose_fix, F.text.casefold() == "📐 заменить чертежи")
@router.message(ProposalReviewStates.choose_fix, F.text.casefold() == "заменить чертежи")
async def fix_drawings_request(message: Message, state: FSMContext) -> None:
    await state.set_state(ProposalReviewStates.replace_drawings)
    await message.answer("Пришли новые чертежи (.rtf/.pdf) или отправь 'Пропустить', чтобы убрать блок чертежей.")


@router.message(ProposalReviewStates.choose_fix, F.text.casefold() == "🗑️ удалить страницу")
@router.message(ProposalReviewStates.choose_fix, F.text.casefold() == "удалить страницу")
async def fix_delete_page_request(message: Message, state: FSMContext) -> None:
    await state.set_state(ProposalReviewStates.delete_page)
    await message.answer("Укажи номер страницы для удаления (начиная с 1).")


@router.message(ProposalReviewStates.replace_template, F.document)
async def fix_template_apply(message: Message, state: FSMContext) -> None:
    document = message.document
    if not document or not document.file_name or not document.file_name.lower().endswith(".docx"):
        await message.answer("Нужен файл шаблона в формате .docx")
        return

    fsm_data = await state.get_data()
    output_dir = Path(fsm_data["output_dir"])
    template_path = output_dir / _safe_file_name(document.file_name, "template_updated.docx")
    await _download_document(message, template_path)
    await state.update_data(template_path=str(template_path))
    await _finalize_kp_build(message, state)


@router.message(ProposalReviewStates.replace_client_name, F.text)
async def fix_client_name_apply(message: Message, state: FSMContext) -> None:
    client_name = (message.text or "").strip()
    if not client_name:
        await message.answer("ФИО клиента не может быть пустым.")
        return
    fsm_data = await state.get_data()
    kp_number = str(fsm_data.get("kp_number", "")).strip()
    await state.update_data(
        client_name=client_name,
        kp_filename=_build_kp_filename(kp_number=kp_number, client_name=client_name),
    )
    await _finalize_kp_build(message, state)


@router.message(ProposalReviewStates.replace_kp_number, F.text)
async def fix_kp_number_apply(message: Message, state: FSMContext) -> None:
    kp_number = (message.text or "").strip()
    if not kp_number:
        await message.answer("Номер КП не может быть пустым.")
        return
    fsm_data = await state.get_data()
    client_name = str(fsm_data.get("client_name", "")).strip()
    await state.update_data(
        kp_number=kp_number,
        kp_filename=_build_kp_filename(kp_number=kp_number, client_name=client_name),
    )
    await _finalize_kp_build(message, state)


@router.message(ProposalReviewStates.replace_excel, F.document)
async def fix_excel_apply(message: Message, state: FSMContext) -> None:
    document = message.document
    if document is None:
        await message.answer("Пришли Excel-файл.")
        return
    ext = Path(document.file_name or "").suffix.lower()
    if ext not in {".xlsx", ".xls", ".xlsm"}:
        await message.answer("Нужен Excel-файл (.xlsx/.xls/.xlsm)")
        return

    fsm_data = await state.get_data()
    output_dir = Path(fsm_data["output_dir"])
    excel_path = output_dir / _safe_file_name(document.file_name or "price_updated.xlsx", "price_updated.xlsx")
    await _download_document(message, excel_path)
    await state.update_data(excel_path=str(excel_path))
    await _finalize_kp_build(message, state)


@router.message(ProposalReviewStates.replace_drawings, F.text.casefold() == "пропустить")
async def fix_drawings_skip(message: Message, state: FSMContext) -> None:
    await state.update_data(drawings_rtf_path=None)
    await _finalize_kp_build(message, state)


@router.message(ProposalReviewStates.replace_drawings, F.document)
async def fix_drawings_apply(message: Message, state: FSMContext) -> None:
    document = message.document
    if document is None:
        await message.answer("Пришли файл чертежей (.rtf/.pdf) или отправь 'Пропустить'.")
        return
    ext = Path(document.file_name or "").suffix.lower()
    if ext not in {".rtf", ".pdf"}:
        await message.answer("Нужен файл чертежей в формате .rtf или .pdf")
        return

    fsm_data = await state.get_data()
    output_dir = Path(fsm_data["output_dir"])
    drawings_path = output_dir / _safe_file_name(document.file_name or f"drawings_updated{ext}", f"drawings_updated{ext}")
    await _download_document(message, drawings_path)
    await state.update_data(drawings_rtf_path=str(drawings_path))
    await _finalize_kp_build(message, state)


@router.message(ProposalReviewStates.delete_page, F.text)
async def fix_delete_page_apply(message: Message, state: FSMContext) -> None:
    raw = (message.text or "").strip()
    if not raw.isdigit():
        await message.answer("Введи номер страницы цифрами, например: 3")
        return

    page_number = int(raw)
    if page_number < 1:
        await message.answer("Номер страницы должен быть больше 0.")
        return

    fsm_data = await state.get_data()
    final_pdf_path = fsm_data.get("last_final_pdf")
    if not final_pdf_path:
        await message.answer("Не найдено последнее КП для исправления. Сформируй КП заново.")
        await state.clear()
        return

    pdf_path = Path(final_pdf_path)
    if not pdf_path.exists():
        await message.answer("Файл КП не найден. Сформируй КП заново.")
        await state.clear()
        return

    reader = PdfReader(str(pdf_path))
    total_pages = len(reader.pages)
    if page_number > total_pages:
        await message.answer(f"В файле только {total_pages} стр. Укажи номер в этом диапазоне.")
        return

    writer = PdfWriter()
    delete_index = page_number - 1
    for idx, page in enumerate(reader.pages):
        if idx != delete_index:
            writer.add_page(page)

    with pdf_path.open("wb") as out_file:
        writer.write(out_file)

    await message.answer("✅ Страница удалена. Отправляю обновлённое КП.")
    await _send_kp_and_review(message, state, pdf_path)
