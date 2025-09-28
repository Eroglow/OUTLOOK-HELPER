from typing import List, Dict, Optional
import os
from PyQt5.QtWidgets import (
	QApplication,
	QMainWindow,
	QWidget,
	QVBoxLayout,
	QHBoxLayout,
	QPushButton,
	QLabel,
	QComboBox,
	QPlainTextEdit,
	QTableWidget,
	QTableWidgetItem,
	QFileDialog,
	QCheckBox,
	QLineEdit,
	QProgressBar,
	QSplitter,
)
from PyQt5.QtCore import Qt, QThread
import pandas as pd

from .data_loader import read_tabular_file, read_from_clipboard, dataframe_to_records
from .worker import SendWorker, SendConfig


class MainWindow(QMainWindow):
	def __init__(self) -> None:
		super().__init__()
		self.setWindowTitle("Excel/剪貼簿群發郵件工具")
		self._df: Optional[pd.DataFrame] = None
		self._columns: List[str] = []
		self._records: List[Dict[str, object]] = []
		self._thread: Optional[QThread] = None
		self._worker: Optional[SendWorker] = None

		self._build_ui()

	def _build_ui(self) -> None:
		central = QWidget()
		root = QVBoxLayout()
		central.setLayout(root)
		self.setCentralWidget(central)

		# Top controls
		row1 = QHBoxLayout()
		self.btn_file = QPushButton("選擇檔案")
		self.btn_clip = QPushButton("從剪貼簿貼上")
		row1.addWidget(self.btn_file)
		row1.addWidget(self.btn_clip)
		row1.addStretch(1)
		root.addLayout(row1)

		# Preview table
		self.table = QTableWidget(0, 0)
		self.table.setEditTriggers(QTableWidget.NoEditTriggers)
		root.addWidget(self.table)

		# Field mapping
		row2 = QHBoxLayout()
		row2.addWidget(QLabel("電子信箱："))
		self.cb_email = QComboBox()
		row2.addWidget(self.cb_email)
		row2.addSpacing(12)
		row2.addWidget(QLabel("姓名："))
		self.cb_name = QComboBox()
		row2.addWidget(self.cb_name)
		row2.addStretch(1)
		root.addLayout(row2)

		# Subject and body
		row3 = QHBoxLayout()
		row3.addWidget(QLabel("主旨："))
		self.ed_subject = QLineEdit()
		row3.addWidget(self.ed_subject)
		root.addLayout(row3)

		split = QSplitter(Qt.Vertical)
		self.body_edit = QPlainTextEdit()
		self.body_edit.setPlaceholderText("親愛的 {姓名}，您好... 支援 {欄位名} 變數")
		split.addWidget(self.body_edit)

		self.log_edit = QPlainTextEdit()
		self.log_edit.setReadOnly(True)
		split.addWidget(self.log_edit)
		root.addWidget(split)

		# Options and actions
		row4 = QHBoxLayout()
		self.chk_html = QCheckBox("HTML 格式")
		self.chk_preview = QCheckBox("逐封預覽 (Display)")
		self.btn_preview = QPushButton("預覽郵件")
		self.btn_send = QPushButton("開始發送")
		self.btn_stop = QPushButton("停止")
		self.btn_stop.setEnabled(False)
		self.btn_attach = QPushButton("選擇附件")
		self.lbl_attach = QLabel("")
		row4.addWidget(self.chk_html)
		row4.addWidget(self.chk_preview)
		row4.addWidget(self.btn_preview)
		row4.addWidget(self.btn_send)
		row4.addWidget(self.btn_stop)
		row4.addWidget(self.btn_attach)
		row4.addWidget(self.lbl_attach)
		row4.addStretch(1)
		root.addLayout(row4)

		# Progress
		row5 = QHBoxLayout()
		self.progress = QProgressBar()
		self.lbl_progress = QLabel("進度：0/0，成功 0，失敗 0")
		row5.addWidget(self.progress)
		row5.addWidget(self.lbl_progress)
		root.addLayout(row5)

		# Connections
		self.btn_file.clicked.connect(self.on_select_file)
		self.btn_clip.clicked.connect(self.on_paste_clipboard)
		self.btn_preview.clicked.connect(self.on_preview)
		self.btn_send.clicked.connect(self.on_send)
		self.btn_stop.clicked.connect(self.on_stop)
		self.btn_attach.clicked.connect(self.on_select_attachment)

	def log(self, message: str) -> None:
		self.log_edit.appendPlainText(message)

	def _populate_table(self, df: pd.DataFrame) -> None:
		self.table.clear()
		self.table.setRowCount(min(len(df), 50))
		self.table.setColumnCount(len(df.columns))
		self.table.setHorizontalHeaderLabels([str(c) for c in df.columns])
		for r in range(min(len(df), 50)):
			for c, col_name in enumerate(df.columns):
				val = df.iloc[r][col_name]
				item = QTableWidgetItem("" if pd.isna(val) else str(val))
				self.table.setItem(r, c, item)

	def _update_mappings(self, columns: List[str]) -> None:
		self.cb_email.clear()
		self.cb_name.clear()
		self.cb_email.addItems(columns)
		self.cb_name.addItems([""] + columns)

	def on_select_file(self) -> None:
		file_path, _ = QFileDialog.getOpenFileName(self, "選擇檔案", "", "Excel/CSV (*.xlsx *.xlsm *.csv)")
		if not file_path:
			return
		df, cols = read_tabular_file(file_path)
		self._df = df
		self._columns = cols
		self._records = dataframe_to_records(df)
		self._populate_table(df)
		self._update_mappings(cols)
		self.log(f"已載入 {len(df)} 筆資料。")
		self.progress.setMaximum(len(self._records))
		self.progress.setValue(0)
		self.lbl_progress.setText(f"進度：0/{len(self._records)}，成功 0，失敗 0")

	def on_paste_clipboard(self) -> None:
		df, cols = read_from_clipboard()
		self._df = df
		self._columns = cols
		self._records = dataframe_to_records(df)
		self._populate_table(df)
		self._update_mappings(cols)
		self.log(f"已從剪貼簿貼上 {len(df)} 筆資料。")
		self.progress.setMaximum(len(self._records))
		self.progress.setValue(0)
		self.lbl_progress.setText(f"進度：0/{len(self._records)}，成功 0，失敗 0")

	def _current_config(self, one_sample_only: bool = False) -> Optional[SendConfig]:
		if not self._records:
			self.log("尚未載入資料。")
			return None
		email_field = self.cb_email.currentText().strip()
		name_field = self.cb_name.currentText().strip() or None
		if not email_field:
			self.log("請選擇電子信箱欄位。")
			return None
		records = self._records[:1] if one_sample_only else self._records
		return SendConfig(
			records=records,
			recipient_field=email_field,
			name_field=name_field,
			subject_template=self.ed_subject.text(),
			body_template=self.body_edit.toPlainText(),
			is_html=self.chk_html.isChecked(),
			preview_only=self.chk_preview.isChecked(),
			attachment_path=self.lbl_attach.text().strip() or None,
		)

	def on_preview(self) -> None:
		config = self._current_config(one_sample_only=True)
		if not config:
			return
		self._start_worker(config)

	def on_send(self) -> None:
		config = self._current_config(one_sample_only=False)
		if not config:
			return
		self._start_worker(config)

	def on_stop(self) -> None:
		if self._worker is not None:
			self._worker.stop()

	def on_select_attachment(self) -> None:
		file_path, _ = QFileDialog.getOpenFileName(self, "選擇附件", "", "所有檔案 (*.*)")
		if not file_path:
			return
		self.lbl_attach.setText(file_path)

	def _start_worker(self, config: SendConfig) -> None:
		if self._thread is not None:
			return
		self.btn_send.setEnabled(False)
		self.btn_stop.setEnabled(True)
		self.progress.setMaximum(len(config.records))
		self.progress.setValue(0)
		self.lbl_progress.setText(f"進度：0/{len(config.records)}，成功 0，失敗 0")
		self._thread = QThread()
		self._worker = SendWorker(config)
		self._worker.moveToThread(self._thread)
		self._thread.started.connect(self._worker.run)
		self._worker.progress.connect(self._on_progress)
		self._worker.log.connect(self.log)
		self._worker.finished.connect(self._on_finished)
		self._worker.cancelled.connect(self._on_finished)
		self._thread.start()

	def _on_progress(self, processed: int, success: int, failed: int) -> None:
		self.progress.setValue(processed)
		self.lbl_progress.setText(f"進度：{processed}/{self.progress.maximum()}，成功 {success}，失敗 {failed}")

	def _on_finished(self) -> None:
		if self._thread is not None:
			self._thread.quit()
			self._thread.wait()
			self._thread = None
			self._worker = None
		self.btn_send.setEnabled(True)
		self.btn_stop.setEnabled(False)
