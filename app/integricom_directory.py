from __future__ import annotations

import sqlite3
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path

INTEGRICOM_DIRECTORY_DB = Path(__file__).resolve().parent / "data" / "integricom_users.sqlite3"


@dataclass
class IntegricomDirectoryUser:
    email: str
    first_name: str
    last_name: str
    branch: str
    is_active: bool
    created_at: str
    updated_at: str
    last_seen_at: str | None


def _utc_now() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def _connect() -> sqlite3.Connection:
    INTEGRICOM_DIRECTORY_DB.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(INTEGRICOM_DIRECTORY_DB, timeout=30)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    return conn


def init_integricom_directory() -> None:
    with _connect() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS integricom_users (
                email TEXT PRIMARY KEY,
                first_name TEXT NOT NULL DEFAULT '',
                last_name TEXT NOT NULL DEFAULT '',
                branch TEXT NOT NULL,
                is_active INTEGER NOT NULL DEFAULT 1,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                last_seen_at TEXT
            )
            """
        )
        conn.execute(
            "UPDATE integricom_users SET branch = 'Home Office' WHERE TRIM(COALESCE(branch, '')) = ''"
        )
        conn.commit()


def list_integricom_users(*, active_only: bool = False) -> dict[str, IntegricomDirectoryUser]:
    init_integricom_directory()
    with _connect() as conn:
        query = """
            SELECT email, first_name, last_name, branch, is_active, created_at, updated_at, last_seen_at
            FROM integricom_users
        """
        if active_only:
            query += " WHERE is_active = 1"
        rows = conn.execute(query).fetchall()

    users: dict[str, IntegricomDirectoryUser] = {}
    for row in rows:
        email = (row["email"] or "").strip().lower()
        if not email:
            continue
        users[email] = IntegricomDirectoryUser(
            email=email,
            first_name=row["first_name"] or "",
            last_name=row["last_name"] or "",
            branch=row["branch"] or "",
            is_active=bool(row["is_active"]),
            created_at=row["created_at"] or "",
            updated_at=row["updated_at"] or "",
            last_seen_at=row["last_seen_at"],
        )
    return users


def upsert_integricom_users(
    users: list[dict[str, str]],
    *,
    last_seen_at: str | None = None,
) -> None:
    if not users:
        return

    init_integricom_directory()
    now = _utc_now()
    seen = last_seen_at or now
    with _connect() as conn:
        for user in users:
            email = (user.get("email") or "").strip().lower()
            if not email:
                continue
            first_name = (user.get("first_name") or "").strip()
            last_name = (user.get("last_name") or "").strip()
            branch = (user.get("branch") or "").strip()
            if not branch:
                continue

            conn.execute(
                """
                INSERT INTO integricom_users (
                    email, first_name, last_name, branch, is_active, created_at, updated_at, last_seen_at
                )
                VALUES (?, ?, ?, ?, 1, ?, ?, ?)
                ON CONFLICT(email) DO UPDATE SET
                    first_name=excluded.first_name,
                    last_name=excluded.last_name,
                    branch=excluded.branch,
                    is_active=1,
                    updated_at=excluded.updated_at,
                    last_seen_at=excluded.last_seen_at
                """,
                (email, first_name, last_name, branch, now, now, seen),
            )
        conn.commit()


def touch_seen_integricom_users(users: list[dict[str, str]]) -> None:
    if not users:
        return

    init_integricom_directory()
    now = _utc_now()
    with _connect() as conn:
        for user in users:
            email = (user.get("email") or "").strip().lower()
            if not email:
                continue
            first_name = (user.get("first_name") or "").strip()
            last_name = (user.get("last_name") or "").strip()
            conn.execute(
                """
                UPDATE integricom_users
                SET
                    first_name=CASE WHEN ? <> '' THEN ? ELSE first_name END,
                    last_name=CASE WHEN ? <> '' THEN ? ELSE last_name END,
                    is_active=1,
                    updated_at=?,
                    last_seen_at=?
                WHERE email=?
                """,
                (first_name, first_name, last_name, last_name, now, now, email),
            )
        conn.commit()


def find_missing_integricom_users(current_emails: set[str]) -> list[IntegricomDirectoryUser]:
    init_integricom_directory()
    normalized = {email.strip().lower() for email in current_emails if email.strip()}

    with _connect() as conn:
        if not normalized:
            rows = conn.execute(
                """
                SELECT email, first_name, last_name, branch, is_active, created_at, updated_at, last_seen_at
                FROM integricom_users
                WHERE is_active = 1
                ORDER BY email
                """
            ).fetchall()
        else:
            placeholders = ",".join("?" for _ in normalized)
            rows = conn.execute(
                f"""
                SELECT email, first_name, last_name, branch, is_active, created_at, updated_at, last_seen_at
                FROM integricom_users
                WHERE is_active = 1
                  AND email NOT IN ({placeholders})
                ORDER BY email
                """,
                tuple(sorted(normalized)),
            ).fetchall()

    missing: list[IntegricomDirectoryUser] = []
    for row in rows:
        missing.append(
            IntegricomDirectoryUser(
                email=row["email"] or "",
                first_name=row["first_name"] or "",
                last_name=row["last_name"] or "",
                branch=row["branch"] or "",
                is_active=bool(row["is_active"]),
                created_at=row["created_at"] or "",
                updated_at=row["updated_at"] or "",
                last_seen_at=row["last_seen_at"],
            )
        )
    return missing


def deactivate_integricom_users(emails: list[str]) -> int:
    if not emails:
        return 0

    init_integricom_directory()
    now = _utc_now()
    normalized = sorted({(email or "").strip().lower() for email in emails if (email or "").strip()})
    if not normalized:
        return 0

    placeholders = ",".join("?" for _ in normalized)
    with _connect() as conn:
        result = conn.execute(
            f"""
            UPDATE integricom_users
            SET is_active = 0, updated_at = ?
            WHERE LOWER(TRIM(email)) IN ({placeholders})
            """,
            (now, *normalized),
        )
        conn.commit()
        return int(result.rowcount or 0)
