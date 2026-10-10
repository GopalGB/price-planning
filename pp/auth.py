"""Password, session, Origin, CSRF, and login-rate-limit services."""

from __future__ import annotations

import base64
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
import hashlib
import hmac
import re
import secrets
import threading
import time
import uuid

from .db import token_hash
from .errors import AppError


EMAIL = re.compile(r"[^\s@]+@[^\s@]+\.[^\s@]+")
SCRYPT_N = 2 ** 14


@dataclass(frozen=True)
class LoginAttempt:
    email: str
    password: str
    email_key: str
    ip_key: str
    user: dict | None
    encoded: str


@dataclass(frozen=True)
class LoginVerification:
    valid: bool
    upgraded_hash: str | None


def hash_password(password: str) -> str:
    salt = secrets.token_bytes(16)
    digest = hashlib.scrypt(password.encode(), salt=salt, n=SCRYPT_N, r=8, p=1, dklen=32)
    return "scrypt$16384$8$1$" + base64.urlsafe_b64encode(salt).decode() + "$" + base64.urlsafe_b64encode(digest).decode()


def verify_password(password: str, encoded: str) -> bool:
    if isinstance(encoded, str) and ":" in encoded and "$" not in encoded:
        return _verify_legacy_password(password, encoded)
    try:
        name, n, r, p, salt_text, digest_text = encoded.split("$")
        if name != "scrypt":
            return False
        salt, expected = base64.urlsafe_b64decode(salt_text), base64.urlsafe_b64decode(digest_text)
        actual = hashlib.scrypt(password.encode(), salt=salt, n=int(n), r=int(r), p=int(p), dklen=len(expected))
        return hmac.compare_digest(actual, expected)
    except (ValueError, TypeError):
        return False


def _verify_legacy_password(password, encoded):
    try:
        salt_text, digest_text = encoded.split(":")
        salt, expected = bytes.fromhex(salt_text), bytes.fromhex(digest_text)
        if len(salt) != 16 or len(expected) != 64:
            return False
        actual = hashlib.scrypt(password.encode(), salt=salt, n=SCRYPT_N, r=8, p=1, dklen=len(expected))
        return hmac.compare_digest(actual, expected)
    except (ValueError, TypeError):
        return False


class AuthService:
    def __init__(self, database, config):
        self.database = database
        self.config = config
        self._dummy = hash_password("constant-dummy-password")
        self._login_lock = threading.Lock()
        self._pending = {}

    def signup(self, email: object, name: object, password: object) -> tuple[dict, str, str]:
        if not self.config.demo_mode:
            raise AppError(403, "forbidden", "Signup is disabled")
        email, name, password = _validate_signup(email, name, password)
        user_id, now = uuid.uuid4().hex, _now()
        with self.database.transaction() as connection:
            if connection.execute("SELECT 1 FROM pp_users WHERE email=?", (email,)).fetchone():
                raise AppError(409, "exists", "This email already has an account")
            connection.execute(
                "INSERT INTO pp_users(id,email,name,password_hash,created_at) VALUES(?,?,?,?,?)",
                (user_id, email, name, hash_password(password), now),
            )
            user = {"id": user_id, "email": email, "name": name}
            token, csrf = self._session(connection, user_id)
        return user, token, csrf

    def login(self, email: object, password: object, remote: str) -> tuple[dict, str, str]:
        attempt = self.begin_login(email, password, remote)
        try:
            verification = self.verify_login(attempt)
        except Exception:
            self.cancel_login(attempt)
            raise
        return self.finish_login(attempt, verification)

    def begin_login(self, email: object, password: object, remote: str) -> LoginAttempt:
        email = str(email or "").strip().lower()
        password = str(password or "")
        email_key = _login_key("email", remote, email)
        ip_key = _login_key("ip", remote)
        with self._login_lock:
            user = self._reserve_login(email, (email_key, ip_key))
        encoded = user["password_hash"] if user else self._dummy
        return LoginAttempt(email, password, email_key, ip_key, user, encoded)

    def verify_login(self, attempt: LoginAttempt) -> LoginVerification:
        valid = verify_password(attempt.password, attempt.encoded)
        valid = attempt.user is not None and valid
        upgraded = hash_password(attempt.password) if valid and not attempt.encoded.startswith("scrypt$") else None
        return LoginVerification(valid, upgraded)

    def finish_login(self, attempt: LoginAttempt, verification: LoginVerification):
        with self._login_lock:
            try:
                result = self._record_login(attempt, verification)
            finally:
                self._release_pending(attempt)
        if not verification.valid:
            raise AppError(401, "invalid_credentials", "Email or password is incorrect")
        return result

    def cancel_login(self, attempt: LoginAttempt) -> None:
        with self._login_lock:
            self._release_pending(attempt)

    def _reserve_login(self, email, keys):
        cutoff = time.time() - 60
        with self.database.transaction() as connection:
            connection.execute("DELETE FROM pp_login_attempts WHERE attempted_at<?", (cutoff,))
            counts = _failure_counts(connection, keys)
            if any(counts.get(key, 0) + self._pending.get(key, 0) >= 5 for key in keys):
                raise AppError(429, "rate_limited", "Too many failed logins. Try again in one minute")
            row = connection.execute("SELECT * FROM pp_users WHERE email=?", (email,)).fetchone()
        for key in keys:
            self._pending[key] = self._pending.get(key, 0) + 1
        return dict(row) if row else None

    def _record_login(self, attempt, verification):
        with self.database.transaction() as connection:
            if not verification.valid:
                connection.executemany(
                    "INSERT INTO pp_login_attempts(key,attempted_at) VALUES(?,?)",
                    [(attempt.email_key, time.time()), (attempt.ip_key, time.time())],
                )
                return None
            connection.execute("DELETE FROM pp_login_attempts WHERE key=?", (attempt.email_key,))
            if verification.upgraded_hash:
                connection.execute("UPDATE pp_users SET password_hash=? WHERE id=?",
                                   (verification.upgraded_hash, attempt.user["id"]))
            token, csrf = self._session(connection, attempt.user["id"])
            user = {key: attempt.user[key] for key in ("id", "email", "name")}
            return user, token, csrf

    def _release_pending(self, attempt):
        for key in (attempt.email_key, attempt.ip_key):
            remaining = self._pending.get(key, 0) - 1
            if remaining:
                self._pending[key] = remaining
            else:
                self._pending.pop(key, None)

    def authenticate(self, token: str | None) -> tuple[dict, str]:
        if not token:
            raise AppError(401, "unauthorized", "Sign in is required")
        with self.database.connect() as connection:
            row = connection.execute(
                "SELECT u.id,u.email,u.name,s.csrf,s.expires_at FROM pp_sessions s JOIN pp_users u ON u.id=s.user_id WHERE s.token_hash=?",
                (token_hash(token),),
            ).fetchone()
        if row is None or row["expires_at"] <= _now():
            raise AppError(401, "unauthorized", "Session expired")
        return {"id": row["id"], "email": row["email"], "name": row["name"]}, row["csrf"]

    def logout(self, token: str) -> None:
        with self.database.transaction() as connection:
            connection.execute("DELETE FROM pp_sessions WHERE token_hash=?", (token_hash(token),))

    def _session(self, connection, user_id: str) -> tuple[str, str]:
        token, csrf, now = secrets.token_urlsafe(32), secrets.token_urlsafe(24), datetime.now(timezone.utc)
        expires = now + timedelta(hours=8)
        connection.execute(
            "INSERT INTO pp_sessions(token_hash,user_id,csrf,expires_at,created_at) VALUES(?,?,?,?,?)",
            (token_hash(token), user_id, csrf, _iso(expires), _iso(now)),
        )
        return token, csrf


def require_origin(origin: str | None, allowed_origins: tuple[str, ...]) -> None:
    if not origin or origin.rstrip("/") not in allowed_origins:
        raise AppError(403, "origin", "Request Origin is not allowed")


def _login_key(kind, remote, email=""):
    return hashlib.sha256(f"{kind}|{remote}|{email}".encode()).hexdigest()


def _failure_counts(connection, keys):
    rows = connection.execute(
        "SELECT key,COUNT(*) AS failures FROM pp_login_attempts WHERE key IN (?,?) GROUP BY key",
        keys,
    ).fetchall()
    return {row["key"]: row["failures"] for row in rows}


def require_csrf(actual: str | None, expected: str) -> None:
    if not actual or not hmac.compare_digest(actual, expected):
        raise AppError(403, "csrf", "Refresh the page before making changes")


def _validate_signup(email, name, password):
    if not all(isinstance(value, str) for value in (email, name, password)):
        raise AppError(422, "invalid", "Email, name, and password are required")
    email, name = email.strip().lower(), name.strip()
    if not EMAIL.fullmatch(email) or not name or len(name) > 100:
        raise AppError(422, "invalid", "Enter a valid name and email")
    if not 12 <= len(password) <= 1024:
        raise AppError(422, "invalid", "Password must be 12 to 1,024 characters")
    return email, name, password


def _iso(value):
    return value.isoformat().replace("+00:00", "Z")


def _now():
    return _iso(datetime.now(timezone.utc))
