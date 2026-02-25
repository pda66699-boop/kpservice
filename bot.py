from __future__ import annotations

import asyncio
import logging

from aiogram import Bot, Dispatcher
from aiogram.fsm.storage.memory import MemoryStorage

from config import TOKEN
from handlers import router as handlers_router
from handlers.fsm import router as fsm_router


async def main() -> None:
    if not TOKEN or not TOKEN.strip():
        raise ValueError("TOKEN is not set in config.py")

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
    )

    bot = Bot(token=TOKEN)
    dp = Dispatcher(storage=MemoryStorage())

    dp.include_router(handlers_router)
    dp.include_router(fsm_router)

    await bot.delete_webhook(drop_pending_updates=True)
    await dp.start_polling(bot)
    return None


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except (KeyboardInterrupt, SystemExit):
        logging.info("Bot stopped")
