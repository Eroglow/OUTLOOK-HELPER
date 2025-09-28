from typing import Optional, List, Dict
from dataclasses import dataclass
from .utils import is_windows_platform


@dataclass
class MailRequest:
	recipient: str
	subject: str
	body: str
	is_html: bool = False
	attachment_path: Optional[str] = None
	preview_only: bool = False  # True => Display(), False => Send()


class OutlookMailer:
	"""COM-based Outlook mail sender. No-ops on non-Windows for dev/test."""

	def __init__(self) -> None:
		self._outlook = None
		self._namespace = None
		if is_windows_platform():
			self._ensure_outlook()

	def _ensure_outlook(self) -> None:
		if not is_windows_platform():
			return
		try:
			import win32com.client as win32  # type: ignore
			self._outlook = win32.gencache.EnsureDispatch("Outlook.Application")
			self._namespace = self._outlook.GetNamespace("MAPI")
		except Exception as exc:  # pragma: no cover
			raise RuntimeError(f"Failed to initialize Outlook COM: {exc}")

	def send(self, request: MailRequest) -> None:
		if not is_windows_platform():
			# For non-Windows dev, simply simulate success
			return
		if self._outlook is None:
			self._ensure_outlook()
		mail = self._outlook.CreateItem(0)  # olMailItem
		mail.To = request.recipient
		mail.Subject = request.subject
		if request.is_html:
			mail.HTMLBody = request.body
		else:
			mail.Body = request.body
		if request.attachment_path:
			mail.Attachments.Add(Source=request.attachment_path)
		if request.preview_only:
			mail.Display()
		else:
			mail.Send()
