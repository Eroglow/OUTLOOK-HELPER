import os
import re
from typing import Optional


EMAIL_REGEX = re.compile(
	# Simplified RFC 5322 compliant pattern for practical validation
	r"^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$"
)


def is_windows_platform() -> bool:
	return os.name == "nt"


def validate_email(email_address: str) -> bool:
	"""Return True if the given string looks like an email address."""
	if not email_address:
		return False
	return EMAIL_REGEX.match(email_address.strip()) is not None


def coalesce_line_endings(text: str) -> str:
	"""Normalize line endings to Windows CRLF to improve Outlook rendering."""
	if text is None:
		return ""
	return text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")


def safe_getenv(name: str, default: Optional[str] = None) -> Optional[str]:
	value = os.getenv(name)
	return value if value is not None else default
