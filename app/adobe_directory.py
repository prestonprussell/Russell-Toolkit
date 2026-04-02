from __future__ import annotations

import sqlite3
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path

ADOBE_DIRECTORY_DB = Path(__file__).resolve().parent / "data" / "adobe_users.sqlite3"


@dataclass
class AdobeDirectoryUser:
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
    ADOBE_DIRECTORY_DB.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(ADOBE_DIRECTORY_DB, timeout=30)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    return conn


def _rebuild_adobe_users_table(conn: sqlite3.Connection, columns: set[str]) -> None:
    now = _utc_now()
    conn.execute(
        """
        CREATE TABLE adobe_users_new (
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

    first_name_expr = "COALESCE(first_name, '')" if "first_name" in columns else "''"
    last_name_expr = "COALESCE(last_name, '')" if "last_name" in columns else "''"
    branch_expr = (
        "COALESCE(NULLIF(TRIM(branch), ''), 'Home Office')" if "branch" in columns else "'Home Office'"
    )
    is_active_expr = "COALESCE(is_active, 1)" if "is_active" in columns else "1"
    created_at_expr = f"COALESCE(created_at, '{now}')" if "created_at" in columns else f"'{now}'"
    updated_at_expr = f"COALESCE(updated_at, '{now}')" if "updated_at" in columns else f"'{now}'"
    last_seen_expr = "last_seen_at" if "last_seen_at" in columns else "NULL"

    conn.execute(
        f"""
        INSERT OR REPLACE INTO adobe_users_new (
            email, first_name, last_name, branch, is_active, created_at, updated_at, last_seen_at
        )
        SELECT
            LOWER(TRIM(email)),
            {first_name_expr},
            {last_name_expr},
            {branch_expr},
            {is_active_expr},
            {created_at_expr},
            {updated_at_expr},
            {last_seen_expr}
        FROM adobe_users
        WHERE TRIM(COALESCE(email, '')) <> ''
        """
    )
    conn.execute("DROP TABLE adobe_users")
    conn.execute("ALTER TABLE adobe_users_new RENAME TO adobe_users")


def init_adobe_directory() -> None:
    with _connect() as conn:
        exists = conn.execute(
            "SELECT 1 FROM sqlite_master WHERE type='table' AND name='adobe_users'"
        ).fetchone()
        if not exists:
            conn.execute(
                """
                CREATE TABLE adobe_users (
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
            conn.commit()
            return

        info = conn.execute("PRAGMA table_info(adobe_users)").fetchall()
        columns = {str(row["name"]).lower() for row in info}
        required = {
            "email",
            "first_name",
            "last_name",
            "branch",
            "is_active",
            "created_at",
            "updated_at",
            "last_seen_at",
        }
        needs_rebuild = "department" in columns or not required.issubset(columns)
        if needs_rebuild:
            _rebuild_adobe_users_table(conn, columns)

        conn.execute(
            "UPDATE adobe_users SET branch = 'Home Office' WHERE TRIM(COALESCE(branch, '')) = ''"
        )
        conn.commit()


def list_adobe_users(*, active_only: bool = False) -> dict[str, AdobeDirectoryUser]:
    init_adobe_directory()
    with _connect() as conn:
        query = """
            SELECT email, first_name, last_name, branch, is_active, created_at, updated_at, last_seen_at
            FROM adobe_users
        """
        if active_only:
            query += " WHERE is_active = 1"
        rows = conn.execute(query).fetchall()

    users: dict[str, AdobeDirectoryUser] = {}
    for row in rows:
        email = (row["email"] or "").strip().lower()
        if not email:
            continue
        users[email] = AdobeDirectoryUser(
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


def upsert_adobe_users(
    users: list[dict[str, str]],
    *,
    last_seen_at: str | None = None,
) -> None:
    if not users:
        return

    init_adobe_directory()
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
                INSERT INTO adobe_users (
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


def touch_seen_users(users: list[dict[str, str]]) -> None:
    if not users:
        return

    init_adobe_directory()
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
                UPDATE adobe_users
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


def find_missing_users(current_emails: set[str]) -> list[AdobeDirectoryUser]:
    init_adobe_directory()
    normalized = {email.strip().lower() for email in current_emails if email.strip()}

    with _connect() as conn:
        if not normalized:
            rows = conn.execute(
                """
                SELECT email, first_name, last_name, branch, is_active, created_at, updated_at, last_seen_at
                FROM adobe_users
                WHERE is_active = 1
                ORDER BY email
                """
            ).fetchall()
        else:
            placeholders = ",".join("?" for _ in normalized)
            rows = conn.execute(
                f"""
                SELECT email, first_name, last_name, branch, is_active, created_at, updated_at, last_seen_at
                FROM adobe_users
                WHERE is_active = 1
                  AND email NOT IN ({placeholders})
                ORDER BY email
                """,
                tuple(sorted(normalized)),
            ).fetchall()

    missing: list[AdobeDirectoryUser] = []
    for row in rows:
        missing.append(
            AdobeDirectoryUser(
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


def deactivate_adobe_users(emails: list[str]) -> int:
    if not emails:
        return 0

    init_adobe_directory()
    now = _utc_now()
    normalized = sorted({(email or "").strip().lower() for email in emails if (email or "").strip()})
    if not normalized:
        return 0

    placeholders = ",".join("?" for _ in normalized)
    with _connect() as conn:
        result = conn.execute(
            f"""
            UPDATE adobe_users
            SET is_active = 0, updated_at = ?
            WHERE LOWER(TRIM(email)) IN ({placeholders})
            """,
            (now, *normalized),
        )
        conn.commit()
        return int(result.rowcount or 0)
