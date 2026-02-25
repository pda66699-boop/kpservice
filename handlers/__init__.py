from aiogram import Router
from aiogram.filters import Command, CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.types import KeyboardButton, Message, ReplyKeyboardMarkup

from handlers.fsm import ManagerProfileStates, get_flow_keyboard, get_generation_keyboard, has_manager_profile

router = Router(name="handlers_root")


@router.message(CommandStart())
async def start_command(message: Message, state: FSMContext) -> None:
    user_id = message.from_user.id if message.from_user else None
    if has_manager_profile(user_id):
        current_state = await state.get_state()
        if current_state and current_state.startswith("KpBuildStates:"):
            if current_state != "KpBuildStates:template_docx":
                await message.answer(
                    "🚀 Продолжаем формирование КП.",
                    reply_markup=get_flow_keyboard(),
                )
                return
            await message.answer(
                "🚀 Продолжаем формирование КП.\n"
                "Нажми «📤 Загрузить шаблон», чтобы перейти к шагу 1.",
                reply_markup=get_generation_keyboard(),
            )
            return

        keyboard = ReplyKeyboardMarkup(
            keyboard=[
                [KeyboardButton(text="🚀 Начать генерацию КП")],
                [KeyboardButton(text="👤 Данные менеджера")],
                [KeyboardButton(text="📖 Инструкция")],
            ],
            resize_keyboard=True,
        )
        await message.answer(
            "👋 Привет! Желаешь сгенерировать новое КП?",
            reply_markup=keyboard,
        )
        return

    keyboard = ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="🚀 Начать генерацию КП")],
            [KeyboardButton(text="👤 Данные менеджера")],
            [KeyboardButton(text="📖 Инструкция")],
        ],
        resize_keyboard=True,
    )
    await message.answer(
        "👋 Привет! Я помогу тебе собрать КП для клиента всего в 6 шагов.\n"
        "🧭 Твоя задача — просто пошагово предоставить мне нужную информацию.",
        reply_markup=keyboard,
    )

    if not has_manager_profile(user_id):
        await state.clear()
        await state.set_state(ManagerProfileStates.full_name)
        await message.answer(
            "Сначала необходимо добавить данные менеджера — они сохранятся в памяти и будут использоваться при формировании КП.\n"
            "В любой момент ты можешь изменить свои данные через кнопку «👤 Данные менеджера»."
        )
        await message.answer("Укажи свои ФИО (достаточно фамилии и имени).")


@router.message(Command("help"))
async def help_command(message: Message) -> None:
    await message.answer(
        "📌 MVP-сценарий:\n"
        "1) Нажми '🚀 Начать генерацию КП'\n"
        "2) Нажми '📤 Загрузить шаблон' и загрузи шаблон .docx\n"
        "3) Введи данные клиента/КП/телефон\n"
        "4) Загрузи Excel (.xlsx/.xls/.xlsm)\n"
        "5) Загрузи чертежи (.rtf или .pdf) или отправь 'Пропустить'\n"
        "6) Получи итоговый PDF ✅"
    )


@router.message(lambda message: (message.text or "").casefold() == "📖 инструкция".casefold())
@router.message(lambda message: (message.text or "").casefold() == "инструкция")
async def send_instruction(message: Message) -> None:
    await message.answer(
        "📘 Инструкция по выбору шаблона КП\n\n"
        "🧩 Выбор нужного шаблона КП:\n"
        "Определи, что именно интересует клиента, и скачай соответствующий шаблон.\n\n"
        "🏛️ Порталы\n"
        "Для чего: складные двери (типа «гармошка») и подъёмно-сдвижные конструкции для террас, беседок, балконов.\n\n"
        "🧱 Фасады\n"
        "Для чего: стоечно-ригельное остекление, панорамные решения.\n\n"
        "🌿 Зимние сады и зенитные фонари\n"
        "Для чего: зимние сады, оранжереи, зенитные фонари для естественного освещения.\n\n"
        "🏡 Коттеджи (комплексное решение)\n"
        "Для чего: полное остекление и архитектурные решения для частных домов «под ключ».\n\n"
        "💡 Совет:\n"
        "Если клиент «хочет всё и сразу» — смело берите комплексную форму. "
        "Это сэкономит время и покажет экспертизу компании.\n\n"
        "✍️ При заполнении персональных данных проверь, чтобы все поля были заполнены аккуратно и без опечаток.\n"
        "🔢 Номер КП заполняется по внутренней нумерации.\n"
        "👤 ФИО заказчика — как в договоре или заявке.\n\n"
        "🧑‍💼 Для заполнения информации о менеджере (твои ФИО, телефон) "
        "нажми «👤 Данные менеджера» в главном меню."
    )
