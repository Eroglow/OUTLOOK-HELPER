import os
import sys
import re
import threading
import queue
from dataclasses import dataclass
from typing import List, Optional, Tuple

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd


EMAIL_REGEX = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")


RECIPIENT_NAME_HEADERS = {
    "name", "姓名", "收件人", "receiver", "recipient", "full name", "名稱", "名字",
}
RECIPIENT_EMAIL_HEADERS = {
    "email", "e-mail", "mail", "電子郵件", "郵件", "郵箱", "信箱",
}


@dataclass
class Recipient:
    display_name: str
    email: str


def normalize_header(header: str) -> str:
    if header is None:
        return ""
    return str(header).strip().lower().replace(" ", "")


def extract_recipients_from_dataframe(df: pd.DataFrame) -> List[Recipient]:
    if df is None or df.empty:
        return []

    normalized_columns = [normalize_header(c) for c in df.columns]
    column_map = {normalize_header(c): c for c in df.columns}

    name_col: Optional[str] = None
    email_col: Optional[str] = None

    # Try to find explicit name/email columns by known headers
    for col_key in normalized_columns:
        if col_key in {h.replace(" ", "") for h in RECIPIENT_NAME_HEADERS}:
            name_col = column_map[col_key]
        if col_key in {h.replace(" ", "") for h in RECIPIENT_EMAIL_HEADERS}:
            email_col = column_map[col_key]

    # If email column not found, search for any column containing @-like content
    if email_col is None:
        for original_col in df.columns:
            series = df[original_col].astype(str)
            if series.str.contains(EMAIL_REGEX).any():
                email_col = original_col
                break

    recipients: List[Recipient] = []
    if email_col is None:
        return recipients

    for _, row in df.iterrows():
        raw_email = str(row.get(email_col, "")).strip()
        match = EMAIL_REGEX.search(raw_email)
        if not match:
            continue
        email = match.group(0)

        candidate_name = ""
        if name_col is not None:
            candidate_name = str(row.get(name_col, "")).strip()

        if not candidate_name:
            # Use prefix before @ as fallback
            candidate_name = email.split("@")[0]

        recipients.append(Recipient(display_name=candidate_name, email=email))

    return recipients


def parse_clipboard_text_to_dataframe(text: str) -> pd.DataFrame:
    if not text:
        return pd.DataFrame()

    # Normalize line endings
    text = text.replace("\r\n", "\n").replace("\r", "\n")

    # Try tab-separated first (Excel clipboard)
    delimiter = "\t" if "\t" in text else ","

    lines = [line for line in text.split("\n") if line.strip() != ""]
    if not lines:
        return pd.DataFrame()

    parts = [line.split(delimiter) for line in lines]

    # Determine if first row is header by checking presence of known headers or non-email content
    first_row = [normalize_header(p) for p in parts[0]]
    header_like = any(
        (p in {h.replace(" ", "") for h in RECIPIENT_NAME_HEADERS | RECIPIENT_EMAIL_HEADERS})
        for p in first_row
    ) or not any(EMAIL_REGEX.search(cell or "") for cell in parts[0])

    if header_like:
        df = pd.DataFrame(parts[1:], columns=parts[0])
    else:
        # Synthesize headers
        max_len = max(len(p) for p in parts)
        columns = [f"col{i+1}" for i in range(max_len)]
        df = pd.DataFrame(parts, columns=columns)

    return df


class BulkMailerApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Outlook 群發郵件工具")
        self.geometry("900x650")

        self.recipients: List[Recipient] = []
        self.attachments: List[str] = []
        self.sending = False
        self.cancel_requested = False
        self.status_queue: "queue.Queue[str]" = queue.Queue()

        self._build_ui()

        # Poll status queue to update UI
        self.after(200, self._drain_status_queue)

    def _build_ui(self) -> None:
        main = ttk.Frame(self, padding=12)
        main.pack(fill=tk.BOTH, expand=True)

        # Top controls
        toolbar = ttk.Frame(main)
        toolbar.pack(fill=tk.X)

        btn_excel = ttk.Button(toolbar, text="選擇 Excel 文件", command=self.on_select_excel)
        btn_excel.pack(side=tk.LEFT, padx=4, pady=4)

        btn_clip = ttk.Button(toolbar, text="從剪貼簿貼上", command=self.on_paste_clipboard)
        btn_clip.pack(side=tk.LEFT, padx=4, pady=4)

        btn_attach = ttk.Button(toolbar, text="選擇附件", command=self.on_select_attachments)
        btn_attach.pack(side=tk.LEFT, padx=4, pady=4)

        # Recipients list
        recipients_frame = ttk.LabelFrame(main, text="已識別的收件人（姓名 + 郵件）")
        recipients_frame.pack(fill=tk.BOTH, expand=True, padx=4, pady=8)

        self.txt_recipients = tk.Text(recipients_frame, height=10)
        self.txt_recipients.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)

        # Subject and body
        form = ttk.Frame(main)
        form.pack(fill=tk.BOTH, expand=False)

        ttk.Label(form, text="主題：").grid(row=0, column=0, sticky=tk.W, padx=4, pady=4)
        self.entry_subject = ttk.Entry(form)
        self.entry_subject.grid(row=0, column=1, sticky=tk.EW, padx=4, pady=4)
        form.columnconfigure(1, weight=1)

        ttk.Label(form, text="郵件內容（可用 {姓名}）：").grid(row=1, column=0, sticky=tk.NW, padx=4, pady=4)
        self.txt_body = tk.Text(form, height=12)
        self.txt_body.grid(row=1, column=1, sticky=tk.NSEW, padx=4, pady=4)
        form.rowconfigure(1, weight=1)

        # Actions
        actions = ttk.Frame(main)
        actions.pack(fill=tk.X)

        self.btn_preview = ttk.Button(actions, text="預覽郵件", command=self.on_preview)
        self.btn_preview.pack(side=tk.LEFT, padx=4, pady=8)

        self.btn_send = ttk.Button(actions, text="發送郵件", command=self.on_send)
        self.btn_send.pack(side=tk.LEFT, padx=4, pady=8)

        self.btn_cancel = ttk.Button(actions, text="中斷", command=self.on_cancel, state=tk.DISABLED)
        self.btn_cancel.pack(side=tk.LEFT, padx=4, pady=8)

        # Status
        self.var_status = tk.StringVar(value="就緒")
        status_bar = ttk.Label(main, textvariable=self.var_status, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(fill=tk.X, padx=2, pady=2)

    def on_select_excel(self) -> None:
        file_path = filedialog.askopenfilename(
            title="選擇 Excel 文件",
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")],
        )
        if not file_path:
            return
        try:
            df = pd.read_excel(file_path)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("讀取失敗", f"無法讀取 Excel：\n{exc}")
            return

        self.recipients = extract_recipients_from_dataframe(df)
        self._render_recipients()

    def on_paste_clipboard(self) -> None:
        try:
            text = self.clipboard_get()
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("讀取失敗", f"無法讀取剪貼簿：\n{exc}")
            return

        df = parse_clipboard_text_to_dataframe(text)
        self.recipients = extract_recipients_from_dataframe(df)
        self._render_recipients()

    def on_select_attachments(self) -> None:
        files = filedialog.askopenfilenames(title="選擇附件")
        if not files:
            return
        self.attachments = list(files)
        messagebox.showinfo("附件", f"已選擇 {len(self.attachments)} 個附件。")

    def _render_recipients(self) -> None:
        self.txt_recipients.delete("1.0", tk.END)
        for r in self.recipients:
            self.txt_recipients.insert(tk.END, f"{r.display_name}\t{r.email}\n")
        self.var_status.set(f"已識別 {len(self.recipients)} 位收件人")

    def _render_body_for(self, recipient: Recipient) -> str:
        raw = self.txt_body.get("1.0", tk.END)
        name = recipient.display_name.strip() or "您好"
        return raw.replace("{姓名}", name)

    def on_preview(self) -> None:
        if not self.recipients:
            messagebox.showwarning("預覽", "請先載入收件人名單。")
            return
        subject = self.entry_subject.get().strip()
        body = self._render_body_for(self.recipients[0])

        preview = tk.Toplevel(self)
        preview.title("預覽第一封郵件")
        preview.geometry("700x500")
        ttk.Label(preview, text=f"To: {self.recipients[0].email}").pack(anchor=tk.W, padx=8, pady=4)
        ttk.Label(preview, text=f"Subject: {subject}").pack(anchor=tk.W, padx=8, pady=4)
        txt = tk.Text(preview)
        txt.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        txt.insert(tk.END, body)
        txt.configure(state=tk.DISABLED)

    def on_send(self) -> None:
        if self.sending:
            return
        if not self.recipients:
            messagebox.showwarning("發送", "請先載入收件人名單。")
            return
        subject = self.entry_subject.get().strip()
        if not subject:
            messagebox.showwarning("發送", "請輸入郵件主題。")
            return

        if not messagebox.askyesno("確認", f"即將發送 {len(self.recipients)} 封郵件，確定嗎？"):
            return

        self.sending = True
        self.cancel_requested = False
        self.btn_send.configure(state=tk.DISABLED)
        self.btn_cancel.configure(state=tk.NORMAL)
        self.var_status.set("開始發送...")

        thread = threading.Thread(target=self._send_worker, daemon=True)
        thread.start()

    def on_cancel(self) -> None:
        if not self.sending:
            return
        self.cancel_requested = True
        self.var_status.set("正在中斷...")

    def _send_worker(self) -> None:
        # Deferred import to avoid non-Windows crash
        try:
            import win32com.client  # type: ignore
        except Exception as exc:  # noqa: BLE001
            self.status_queue.put(f"錯誤：無法載入 Outlook 介面：{exc}")
            self.status_queue.put("_done_error_")
            return

        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
        except Exception as exc:  # noqa: BLE001
            self.status_queue.put(f"錯誤：無法啟動 Outlook：{exc}")
            self.status_queue.put("_done_error_")
            return

        total = len(self.recipients)
        failures: List[Tuple[Recipient, str]] = []

        for idx, r in enumerate(self.recipients, start=1):
            if self.cancel_requested:
                self.status_queue.put("_cancelled_")
                break
            try:
                mail = outlook.CreateItem(0)
                mail.To = r.email
                mail.Subject = self.entry_subject.get().strip()
                mail.Body = self._render_body_for(r)

                for path in self.attachments:
                    if os.path.isfile(path):
                        mail.Attachments.Add(Source=path)

                # For development, you may switch to Display() first
                mail.Send()
            except Exception as exc:  # noqa: BLE001
                failures.append((r, str(exc)))

            self.status_queue.put(f"正在發送第 {idx}/{total} 封郵件...")

        if not self.cancel_requested:
            if failures:
                # Write log next to script
                log_path = os.path.join(os.path.dirname(sys.argv[0]), "send_errors.log")
                try:
                    with open(log_path, "w", encoding="utf-8") as f:
                        for rec, err in failures:
                            f.write(f"{rec.display_name}\t{rec.email}\t{err}\n")
                    self.status_queue.put(f"完成，但有 {len(failures)} 封失敗，已寫入日誌：{log_path}")
                except Exception as exc:  # noqa: BLE001
                    self.status_queue.put(f"完成，但寫入日誌失敗：{exc}")
            else:
                self.status_queue.put(f"全部 {total} 封郵件已成功建立並發送！")

        self.status_queue.put("_done_")

    def _drain_status_queue(self) -> None:
        try:
            while True:
                msg = self.status_queue.get_nowait()
                if msg == "_done_" or msg == "_done_error_" or msg == "_cancelled_":
                    self.sending = False
                    self.btn_send.configure(state=tk.NORMAL)
                    self.btn_cancel.configure(state=tk.DISABLED)
                    if msg == "_cancelled_":
                        self.var_status.set("已中斷發送。")
                    # else keep last message
                else:
                    self.var_status.set(msg)
        except queue.Empty:
            pass
        finally:
            self.after(300, self._drain_status_queue)


def main() -> None:
    app = BulkMailerApp()
    app.mainloop()


if __name__ == "__main__":
    main()

