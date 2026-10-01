"""Simple token-based auth with persistent password (changeable).

Credentials stored in data/credentials.json (created on first run).

Default account (first run / reset):
  username: admin
  password: admin123

Env overrides (only used when credentials file does not exist yet):
  ORDER_ADMIN_USER, ORDER_ADMIN_PASS, ORDER_API_SECRET
"""
from __future__ import annotations

import base64
import hashlib
import hmac
import json
import os
import secrets
import time
from typing import Optional, Tuple

from fastapi import Header, HTTPException, status

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CRED_PATH = os.path.join(BASE_DIR, "data", "credentials.json")

SECRET = os.environ.get("ORDER_API_SECRET", "order-api-dev-secret-change-me")
TOKEN_TTL_SEC = int(os.environ.get("ORDER_TOKEN_TTL", str(7 * 24 * 3600)))  # 7 days

# Defaults only when credentials file is missing
DEFAULT_USER = os.environ.get("ORDER_ADMIN_USER", "admin")
DEFAULT_PASS = os.environ.get("ORDER_ADMIN_PASS", "admin123")


def _b64(data: bytes) -> str:
    return base64.urlsafe_b64encode(data).decode("ascii").rstrip("=")


def _b64decode(s: str) -> bytes:
    pad = "=" * (-len(s) % 4)
    return base64.urlsafe_b64decode(s + pad)


def _hash_password(password: str, salt: str) -> str:
    dk = hashlib.pbkdf2_hmac(
        "sha256",
        password.encode("utf-8"),
        salt.encode("utf-8"),
        120_000,
    )
    return _b64(dk)


def _ensure_credentials() -> dict:
    """Load or create credentials file."""
    os.makedirs(os.path.dirname(CRED_PATH), exist_ok=True)
    if os.path.isfile(CRED_PATH):
        with open(CRED_PATH, "r", encoding="utf-8") as f:
            return json.load(f)

    salt = secrets.token_hex(16)
    data = {
        "username": DEFAULT_USER,
        "salt": salt,
        "password_hash": _hash_password(DEFAULT_PASS, salt),
    }
    with open(CRED_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)
    return data


def _save_credentials(data: dict) -> None:
    os.makedirs(os.path.dirname(CRED_PATH), exist_ok=True)
    with open(CRED_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)


def get_admin_username() -> str:
    return _ensure_credentials().get("username") or DEFAULT_USER


def verify_password(username: str, password: str) -> bool:
    cred = _ensure_credentials()
    if username != cred.get("username"):
        return False
    salt = cred.get("salt") or ""
    expected = cred.get("password_hash") or ""
    actual = _hash_password(password, salt)
    return hmac.compare_digest(actual, expected)


def change_password(username: str, old_password: str, new_password: str) -> Tuple[bool, str]:
    """
    Change password for the admin user.
    Returns (ok, message).
    """
    if not new_password or len(new_password) < 4:
        return False, "New password must be at least 4 characters"
    if not verify_password(username, old_password):
        return False, "Current password is incorrect"

    cred = _ensure_credentials()
    if username != cred.get("username"):
        return False, "User not found"

    salt = secrets.token_hex(16)
    cred["salt"] = salt
    cred["password_hash"] = _hash_password(new_password, salt)
    _save_credentials(cred)
    return True, "Password changed successfully"


def create_token(username: str) -> str:
    exp = int(time.time()) + TOKEN_TTL_SEC
    payload = f"{username}:{exp}".encode("utf-8")
    sig = hmac.new(SECRET.encode("utf-8"), payload, hashlib.sha256).digest()
    return f"{_b64(payload)}.{_b64(sig)}"


def parse_token(token: str) -> Optional[str]:
    """Return username if token valid, else None."""
    try:
        parts = token.strip().split(".")
        if len(parts) != 2:
            return None
        payload = _b64decode(parts[0])
        sig = _b64decode(parts[1])
        expected = hmac.new(SECRET.encode("utf-8"), payload, hashlib.sha256).digest()
        if not hmac.compare_digest(sig, expected):
            return None
        text = payload.decode("utf-8")
        username, exp_s = text.rsplit(":", 1)
        if int(exp_s) < int(time.time()):
            return None
        if username != get_admin_username():
            return None
        return username
    except Exception:
        return None


def _extract_bearer(authorization: Optional[str]) -> Optional[str]:
    if not authorization:
        return None
    parts = authorization.split(None, 1)
    if len(parts) == 2 and parts[0].lower() == "bearer":
        return parts[1].strip()
    return authorization.strip() or None


def get_current_user(authorization: Optional[str] = Header(None)) -> str:
    """Require login — raises 401 if missing/invalid."""
    token = _extract_bearer(authorization)
    if not token:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Login required",
            headers={"WWW-Authenticate": "Bearer"},
        )
    user = parse_token(token)
    if not user:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Invalid or expired token",
            headers={"WWW-Authenticate": "Bearer"},
        )
    return user


def get_optional_user(authorization: Optional[str] = Header(None)) -> Optional[str]:
    """Return username if logged in, else None (no error)."""
    token = _extract_bearer(authorization)
    if not token:
        return None
    return parse_token(token)
