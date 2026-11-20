"""โปรแกรมดึงข้อมูลใบเสร็จรับเงินจาก PDF -> Excel (ลากไฟล์ PDF มาวางได้เลย)"""
import os
import queue
import subprocess
import sys
import threading
from datetime import datetime
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import receipt_core as core

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    Root = TkinterDnD.Tk
    HAS_DND = True
except Exception:  # ไม่มี/โหลด tkinterdnd2 ไม่ได้ -> ใช้ปุ่มเลือกไฟล์แทน
    Root = tk.Tk
    HAS_DND = False

FONT = "Tahoma" if sys.platform.startswith("win") else "Helvetica"


class App(Root):
    def __init__(self):
        super().__init__()
        self.title("ดึงข้อมูลใบเสร็จรับเงิน PDF → Excel")
        self.geometry("980x640")
        self.minsize(820, 520)
        self.q = queue.Queue()
        self.busy = False
        self.last_xlsx = None
        self._build()
        self.after(100, self._poll)

    # ---------- UI ----------
    def _build(self):
        pad = {"padx": 14, "pady": 6}
        self.drop = tk.Label(
            self, text="ลากไฟล์ PDF มาวางที่นี่\nหรือคลิกเพื่อเลือกไฟล์" if HAS_DND else "คลิกเพื่อเลือกไฟล์ PDF",
            font=(FONT, 15), bg="#eaf2fb", fg="#1f4e78", relief="groove", bd=2, height=5, cursor="hand2")
        self.drop.pack(fill="x", **pad)
        self.drop.bind("<Button-1>", lambda e: self.choose())
        if HAS_DND:
            self.drop.drop_target_register(DND_FILES)
            self.drop.dnd_bind("<<Drop>>", self._on_drop)
            self.drop.dnd_bind("<<DropEnter>>", lambda e: self.drop.config(bg="#cfe3f7"))
            self.drop.dnd_bind("<<DropLeave>>", lambda e: self.drop.config(bg="#eaf2fb"))

        bar = ttk.Frame(self); bar.pack(fill="x", **pad)
        self.status = tk.StringVar(value="พร้อมใช้งาน")
        ttk.Label(bar, textvariable=self.status, font=(FONT, 10)).pack(side="left")
        self.btn_folder = ttk.Button(bar, text="เปิดโฟลเดอร์", command=self.open_folder, state="disabled")
        self.btn_folder.pack(side="right", padx=(6, 0))
        self.btn_open = ttk.Button(bar, text="เปิดไฟล์ Excel", command=self.open_xlsx, state="disabled")
        self.btn_open.pack(side="right")

        self.progress = ttk.Progressbar(self, mode="determinate")
        self.progress.pack(fill="x", padx=14)

        cols = ["file", "page"] + [k for k, _ in core.FIELDS]
        self.tree = ttk.Treeview(self, columns=cols, show="headings", height=12)
        widths = {"page": 50, "delivery_date": 110, "receipt_date": 130,
                  "doc_no": 110, "payer": 180, "amount": 90}
        self.tree.heading("file", text="ชื่อไฟล์")
        self.tree.column("file", width=170, anchor="w")
        self.tree.heading("page", text="หน้า")
        self.tree.column("page", width=50, anchor="center")
        for k, h in core.FIELDS:
            self.tree.heading(k, text=h)
            self.tree.column(k, width=widths[k], anchor="e" if k == "amount" else "w")
        self.tree.tag_configure("warn", background="#fff2cc")
        sb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        frame = ttk.Frame(self); frame.pack(fill="both", expand=True, padx=14, pady=6)
        self.tree.pack(in_=frame, side="left", fill="both", expand=True)
        sb.pack(in_=frame, side="right", fill="y")

        self.log = tk.Text(self, height=5, state="disabled", font=(FONT, 9), bg="#f7f7f7")
        self.log.pack(fill="x", padx=14, pady=(0, 10))

    def say(self, msg):
        self.log.config(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.config(state="disabled")

    # ---------- เลือก/วางไฟล์ ----------
    def choose(self):
        if self.busy:
            return
        files = filedialog.askopenfilenames(title="เลือกไฟล์ PDF", filetypes=[("PDF", "*.pdf")])
        if files:
            self.start(list(files))

    def _on_drop(self, event):
        self.drop.config(bg="#eaf2fb")
        files = [f for f in self.tk.splitlist(event.data) if f.lower().endswith(".pdf")]
        if not files:
            messagebox.showwarning("ไฟล์ไม่ถูกต้อง", "กรุณาวางไฟล์ .pdf เท่านั้น")
        elif self.busy:
            messagebox.showinfo("กำลังทำงาน", "รอให้งานปัจจุบันเสร็จก่อน")
        else:
            self.start(files)

    # ---------- ประมวลผล (เธรดแยก ไม่ให้หน้าต่างค้าง) ----------
    def start(self, files):
        self.busy = True
        self.tree.delete(*self.tree.get_children())
        self.btn_open.config(state="disabled"); self.btn_folder.config(state="disabled")
        self.progress.config(value=0)
        threading.Thread(target=self._work, args=(files,), daemon=True).start()

    def _work(self, files):
        """อ่านทุกไฟล์ แล้วรวมเป็น Excel ไฟล์เดียว (ถ้าเลือกไฟล์เดียว ใช้ชื่อเดียวกับ PDF)"""
        if len(files) == 1:
            out = Path(files[0]).with_suffix(".xlsx")
        else:
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            out = Path(files[0]).parent / f"รวมใบเสร็จ_{len(files)}ไฟล์_{stamp}.xlsx"
        all_rows = []
        try:
            self.q.put(("status", "กำลังโหลดโมเดล OCR …"))
            core.get_engine()
        except Exception as e:  # noqa
            self.q.put(("error", f"โหลดโมเดลไม่ได้: {e}"))
            self.q.put(("end",))
            return
        for fi, f in enumerate(files, 1):
            name = Path(f).name
            try:
                total = core.count_pages(f)

                def prog(page, n, r, name=name, fi=fi):
                    self.q.put(("row", page, r))
                    self.q.put(("prog", page, n, f"[{fi}/{len(files)}] {name}: หน้า {page}/{n}"))

                self.q.put(("status", f"กำลังอ่าน {name} ({total} หน้า)"))
                all_rows += core.read_pdf(f, prog)
            except Exception as e:  # noqa
                self.q.put(("error", f"{name}: {e}"))
        if all_rows:
            try:
                core.write_excel(all_rows, out)
                self.q.put(("done", str(out), len(all_rows)))
            except PermissionError:
                self.q.put(("error", f"บันทึกไม่ได้: ไฟล์ Excel '{out.name}' เปิดค้างอยู่ กรุณาปิดก่อนแล้วลองใหม่"))
            except Exception as e:  # noqa
                self.q.put(("error", f"บันทึกไม่ได้: {e}"))
        self.q.put(("end",))

    def _poll(self):
        try:
            while True:
                m = self.q.get_nowait()
                k = m[0]
                if k == "status":
                    self.status.set(m[1]); self.say(m[1])
                elif k == "row":
                    r = m[2]
                    vals = [r.get("file", ""), m[1]] + [("" if r[f] is None else (f"{r[f]:,.2f}" if f == "amount" else r[f])) for f, _ in core.FIELDS]
                    miss = any(r[f] in ("", None) for f, _ in core.FIELDS)
                    self.tree.insert("", "end", values=vals, tags=("warn",) if miss else ())
                elif k == "prog":
                    self.progress.config(maximum=m[2], value=m[1]); self.status.set(m[3])
                elif k == "done":
                    self.last_xlsx = m[1]
                    self.say(f"✔ สร้างไฟล์แล้ว: {m[1]} ({m[2]} หน้า)")
                    self.status.set("เสร็จแล้ว: " + Path(m[1]).name)
                    self.btn_open.config(state="normal"); self.btn_folder.config(state="normal")
                elif k == "error":
                    self.say("✘ " + m[1]); self.status.set("เกิดข้อผิดพลาด")
                    messagebox.showerror("เกิดข้อผิดพลาด", m[1])
                elif k == "end":
                    self.busy = False
        except queue.Empty:
            pass
        self.after(100, self._poll)

    # ---------- เปิดผลลัพธ์ ----------
    def open_xlsx(self):
        if self.last_xlsx:
            _open(self.last_xlsx)

    def open_folder(self):
        if self.last_xlsx:
            _open(str(Path(self.last_xlsx).parent))


def _open(path):
    if sys.platform.startswith("win"):
        os.startfile(path)  # noqa
    elif sys.platform == "darwin":
        subprocess.run(["open", path])
    else:
        subprocess.run(["xdg-open", path])


if __name__ == "__main__":
    App().mainloop()
