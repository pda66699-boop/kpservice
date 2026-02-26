from __future__ import annotations

import sqlite3
from pathlib import Path
from typing import Any

DB_PATH = Path(__file__).resolve().parents[1] / "artifacts" / "manager_profiles.sqlite3"


def _connect() -> sqlite3.Connection:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_manager_profiles_db() -> None:
    with _connect() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS manager_profiles (
                user_id INTEGER PRIMARY KEY,
                manager_name TEXT NOT NULL DEFAULT '',
                manager_phone TEXT NOT NULL DEFAULT '',
                updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            )
            """
        )
        conn.commit()


def get_manager_profile(user_id: int) -> dict[str, str] | None:
    init_manager_profiles_db()
    with _connect() as conn:
        row = conn.execute(
            "SELECT manager_name, manager_phone FROM manager_profiles WHERE user_id = ?",
            (user_id,),
        ).fetchone()
    if row is None:
        return None
    profile: dict[str, str] = {
        "manager_name": str(row["manager_name"] or "").strip(),
        "manager_phone": str(row["manager_phone"] or "").strip(),
    }
    return profile


def has_manager_profile(user_id: int) -> bool:
    profile = get_manager_profile(user_id)
    return bool(profile and profile.get("manager_name") and profile.get("manager_phone"))


def save_manager_profile(
    user_id: int,
    manager_name: str | None = None,
    manager_phone: str | None = None,
) -> dict[str, str]:
    init_manager_profiles_db()
    current = get_manager_profile(user_id) or {"manager_name": "", "manager_phone": ""}
    name = (manager_name if manager_name is not None else current["manager_name"]).strip()
    phone = (manager_phone if manager_phone is not None else current["manager_phone"]).strip()

    with _connect() as conn:
        conn.execute(
            """
            INSERT INTO manager_profiles(user_id, manager_name, manager_phone, updated_at)
            VALUES(?, ?, ?, CURRENT_TIMESTAMP)
            ON CONFLICT(user_id) DO UPDATE SET
                manager_name = excluded.manager_name,
                manager_phone = excluded.manager_phone,
                updated_at = CURRENT_TIMESTAMP
            """,
            (user_id, name, phone),
        )
        conn.commit()

    return {"manager_name": name, "manager_phone": phone}

