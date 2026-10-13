"""Sanitized application errors."""

from dataclasses import dataclass


@dataclass(frozen=True)
class ErrorPayload:
    message: str
    extra: dict


class AppError(Exception):
    def __init__(self, status: int, code: str, payload: str | ErrorPayload):
        message = payload.message if isinstance(payload, ErrorPayload) else payload
        super().__init__(message)
        self.status = status
        self.code = code
        self.message = message
        self.extra = payload.extra if isinstance(payload, ErrorPayload) else {}
