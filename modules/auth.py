import hashlib
import os
import secrets
from datetime import datetime

import streamlit as st
from dotenv import load_dotenv

from modules.database import get_conn, init_db

load_dotenv()

DEFAULT_ADMIN_USERNAME = os.getenv("ADMIN_USERNAME", "admin")
DEFAULT_ADMIN_PASSWORD = os.getenv("ADMIN_PASSWORD", "admin123")


USER_TABLE_SQL = """
create table if not exists users (
    id integer primary key autoincrement,
    username text not null unique,
    password_hash text not null,
    role text not null default 'user',
    full_name text,
    is_active integer not null default 1,
    created_at text default current_timestamp,
    last_login_at text
);
"""


def _hash_password(password: str, salt: str | None = None) -> str:
    if not password:
        raise ValueError("Password cannot be empty.")
    salt = salt or secrets.token_hex(16)
    digest = hashlib.pbkdf2_hmac(
        "sha256",
        password.encode("utf-8"),
        salt.encode("utf-8"),
        120_000,
    ).hex()
    return f"pbkdf2_sha256${salt}${digest}"


def _verify_password(password: str, stored_hash: str) -> bool:
    try:
        method, salt, digest = stored_hash.split("$", 2)
        if method != "pbkdf2_sha256":
            return False
        candidate = _hash_password(password, salt)
        return secrets.compare_digest(candidate, stored_hash)
    except Exception:
        return False


def init_auth_db() -> None:
    init_db()
    with get_conn() as conn:
        conn.executescript(USER_TABLE_SQL)


def ensure_default_admin() -> None:
    init_auth_db()
    with get_conn() as conn:
        existing = conn.execute(
            "select id from users where username = ?",
            (DEFAULT_ADMIN_USERNAME,),
        ).fetchone()
        if existing:
            return
        conn.execute(
            """
            insert into users (username, password_hash, role, full_name, is_active)
            values (?, ?, 'admin', 'Default Admin', 1)
            """,
            (DEFAULT_ADMIN_USERNAME, _hash_password(DEFAULT_ADMIN_PASSWORD)),
        )


def create_user(username: str, password: str, role: str = "user", full_name: str = "") -> tuple[bool, str]:
    username = (username or "").strip().lower()
    role = role if role in {"admin", "user"} else "user"

    if not username:
        return False, "Username is required."
    if not password or len(password) < 6:
        return False, "Password must be at least 6 characters."

    init_auth_db()
    try:
        with get_conn() as conn:
            conn.execute(
                """
                insert into users (username, password_hash, role, full_name, is_active)
                values (?, ?, ?, ?, 1)
                """,
                (username, _hash_password(password), role, full_name.strip()),
            )
        return True, f"User '{username}' created."
    except Exception as exc:
        return False, f"Could not create user: {exc}"


def list_users() -> list[dict]:
    init_auth_db()
    with get_conn() as conn:
        rows = conn.execute(
            """
            select id, username, role, full_name, is_active, created_at, last_login_at
            from users
            order by role, username
            """
        ).fetchall()
        return [dict(row) for row in rows]


def set_user_active(user_id: int, is_active: bool) -> None:
    init_auth_db()
    with get_conn() as conn:
        conn.execute(
            "update users set is_active = ? where id = ?",
            (1 if is_active else 0, user_id),
        )


def authenticate(username: str, password: str) -> tuple[bool, str]:
    init_auth_db()
    username = (username or "").strip().lower()

    with get_conn() as conn:
        user = conn.execute(
            "select * from users where username = ? and is_active = 1",
            (username,),
        ).fetchone()

        if not user or not _verify_password(password, user["password_hash"]):
            return False, "Invalid username or password."

        conn.execute(
            "update users set last_login_at = ? where id = ?",
            (datetime.utcnow().isoformat(), user["id"]),
        )

    st.session_state["authenticated"] = True
    st.session_state["user"] = {
        "id": user["id"],
        "username": user["username"],
        "role": user["role"],
        "full_name": user["full_name"],
    }
    return True, "Logged in."


def current_user() -> dict | None:
    return st.session_state.get("user") if st.session_state.get("authenticated") else None


def logout() -> None:
    st.session_state.pop("authenticated", None)
    st.session_state.pop("user", None)


def require_login(required_role: str | None = None) -> dict:
    ensure_default_admin()
    user = current_user()

    if not user:
        st.warning("Please log in to continue.")
        st.page_link("app.py", label="Go to login page")
        st.stop()

    if required_role and user.get("role") != required_role:
        st.error("You do not have permission to view this page.")
        st.stop()

    with st.sidebar:
        st.write(f"Logged in as **{user['username']}**")
        st.caption(f"Role: {user['role']}")
        if st.button("Log out"):
            logout()
            st.rerun()

    return user
