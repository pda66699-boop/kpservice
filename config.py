from __future__ import annotations

import os

try:
    from dotenv import load_dotenv
except ModuleNotFoundError:  # pragma: no cover
    def load_dotenv() -> bool:
        return False

load_dotenv()

TOKEN = os.getenv("BOT_TOKEN")

if not TOKEN or not TOKEN.strip():
    raise ValueError("BOT_TOKEN is not set. Add it to .env file")
