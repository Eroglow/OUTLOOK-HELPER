from typing import List, Dict, Callable, Optional
from dataclasses import dataclass
from PyQt5.QtCore import QObject, QThread, pyqtSignal
from .mailer import OutlookMailer, MailRequest
from .template import render_subject_and_body
from .utils import validate_email


@dataclass
class SendConfig:
	records: List[Dict[str, object]]
	recipient_field: str
	name_field: Optional[str]
	subject_template: str
	body_template: str
	is_html: bool
	preview_only: bool
	attachment_path: Optional[str]


class SendWorker(QObject):
	progress = pyqtSignal(int, int, int)  # processed, success, failed
	log = pyqtSignal(str)
	finished = pyqtSignal()
	cancelled = pyqtSignal()

	def __init__(self, config: SendConfig) -> None:
		super().__init__()
		self._config = config
		self._stop = False
		self._mailer = OutlookMailer()

	def stop(self) -> None:
		self._stop = True

	def run(self) -> None:
		success = 0
		failed = 0
		processed = 0
		self.log.emit("開始處理...")
		for record in self._config.records:
			if self._stop:
				self.log.emit("已停止。")
				self.cancelled.emit()
				break
			recipient_value = str(record.get(self._config.recipient_field, "")).strip()
			if not validate_email(recipient_value):
				failed += 1
				processed += 1
				self.log.emit(f"略過無效電子郵件：{recipient_value}")
				self.progress.emit(processed, success, failed)
				continue
			context = {k: ("" if v is None else v) for k, v in record.items()}
			rendered = render_subject_and_body(
				self._config.subject_template,
				self._config.body_template,
				context,
				self._config.is_html,
			)
			try:
				request = MailRequest(
					recipient=recipient_value,
					subject=str(rendered["subject"]),
					body=str(rendered["body"]),
					is_html=self._config.is_html,
					attachment_path=self._config.attachment_path,
					preview_only=self._config.preview_only,
				)
				self.log.emit(f"正在寄送給 {recipient_value}...")
				self._mailer.send(request)
				success += 1
			except Exception as exc:  # pragma: no cover
				failed += 1
				self.log.emit(f"寄送給 {recipient_value} 失敗：{exc}")
			processed += 1
			self.progress.emit(processed, success, failed)
		else:
			self.log.emit("全部處理完成。")
			self.finished.emit()
