"""แกนหลัก: PDF สแกน -> OCR (RapidOCR, ภาษาไทย) -> ดึงข้อมูลใบเสร็จ -> Excel
ไม่ขึ้นกับ GUI ใช้ผ่านบรรทัดคำสั่งได้:  python receipt_core.py ไฟล์.pdf [-o ผลลัพธ์.xlsx]
"""
import re
from pathlib import Path

DATE_RE = re.compile(r"(\d{1,2})\s*-\s*([ก-๙]+(?:\.[ก-๙]+)*)\.?\s*-\s*(\d{2,4})")
RCPT_RE = re.compile(r"([0-9A-Za-z]{2})\s*/\s*(\d{1,3})\s*/\s*(\d{3,7})")
AMOUNT_RE = re.compile(r"^\d{1,3}(?:,\d{3})*(?:\.\d{1,2})?$|^\d+(?:\.\d{1,2})?$")

FIELDS = [("delivery_date", "วันที่ใบส่งของ"),
          ("receipt_date", "วันที่ใบเสร็จรับเงิน"), ("doc_no", "เลขที่"),
          ("payer", "ได้รับเงินจาก"), ("amount", "จำนวนเงิน")]

_engine = None


def _models_dir():
    """โฟลเดอร์ models (โมเดล OCR ที่แถมมากับโปรแกรม ใช้ออฟไลน์ ไม่ต้องดาวน์โหลด)"""
    import sys
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / "models"


def get_engine():
    """โหลดโมเดล OCR ครั้งเดียว"""
    global _engine
    if _engine is None:
        from rapidocr import LangRec, ModelType, OCRVersion, RapidOCR
        params = {"Rec.lang_type": LangRec.TH, "Rec.ocr_version": OCRVersion.PPOCRV5,
                  "Rec.model_type": ModelType.MOBILE, "Global.log_level": "error"}
        md = _models_dir()
        files = {"Det.model_path": "PP-OCRv6_det_small.onnx",
                 "Cls.model_path": "ch_ppocr_mobile_v2.0_cls_mobile.onnx",
                 "Rec.model_path": "th_PP-OCRv5_rec_mobile.onnx"}
        if all((md / f).exists() for f in files.values()):
            params.update({k: str(md / f) for k, f in files.items()})
        _engine = RapidOCR(params=params)
    return _engine


def ocr_image(img):
    """img: PIL image -> รายการ token {text,x,y,w,h} พิกัด 0-1 (y=0 บนสุด)"""
    import numpy as np
    W, H = img.size
    r = get_engine()(np.array(img.convert("RGB")))
    out = []
    if r.boxes is None:
        return out
    for box, txt in zip(r.boxes, r.txts):
        xs = [p[0] for p in box]; ys = [p[1] for p in box]
        out.append({"text": txt, "x": min(xs) / W, "y": min(ys) / H,
                    "w": (max(xs) - min(xs)) / W, "h": (max(ys) - min(ys)) / H})
    return out


def render_pages(pdf, dpi=200):
    """เรนเดอร์ทีละหน้าเป็นภาพ (pypdfium2 ไม่ต้องติดตั้งโปรแกรมเพิ่ม)"""
    import pypdfium2 as pdfium
    doc = pdfium.PdfDocument(str(pdf))
    try:
        n = len(doc)
        for i in range(n):
            yield i + 1, n, doc[i].render(scale=dpi / 72).to_pil()
    finally:
        doc.close()


def count_pages(pdf):
    import pypdfium2 as pdfium
    doc = pdfium.PdfDocument(str(pdf))
    try:
        return len(doc)
    finally:
        doc.close()


MONTHS = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."]


def _dist(a, b):
    d = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        prev, d[0] = d[0], i
        for j, cb in enumerate(b, 1):
            prev, d[j] = d[j], min(d[j] + 1, d[j - 1] + 1, prev + (ca != cb))
    return d[-1]


def fix_month(m):
    """แก้ชื่อเดือนที่ OCR อ่านเพี้ยน (เช่น ม.ก. -> ม.ค.) ให้เป็นชื่อเดือนย่อที่ใกล้ที่สุด"""
    m = m.rstrip(".") + "."
    best = min(MONTHS, key=lambda x: _dist(m, x))
    return best if _dist(m, best) <= 1 else m


def clean_name(s):
    s = re.sub(r"^[\s:.\-]+|[\s:.\-ๆ|]+$", "", s)
    return re.sub(r"\s+", " ", s)


def parse_page(tokens):
    toks = sorted(tokens, key=lambda t: (t["y"], t["x"]))
    res = {"delivery_date": "", "receipt_date": "", "doc_no": "", "payer": "", "amount": None}

    # เลขที่ (เลขฝั่งซ้ายของใบเสร็จ เช่น 69/1/00325) แก้ตัวอักษรที่ OCR อ่านเพี้ยน G->6, O->0
    for t in toks:
        if t["x"] < 0.4 and t["y"] < 0.25:
            m = RCPT_RE.search(t["text"])
            if m:
                yy = m.group(1).translate(str.maketrans("GgOoIl", "690011"))
                res["doc_no"] = f"{yy}/{m.group(2)}/{m.group(3)}"
                break

    # วันที่: ใบส่งของอยู่บน, ใบเสร็จรับเงินอยู่ล่าง (ฝั่งขวา)
    dates = []
    for t in toks:
        if t["x"] > 0.5 and t["y"] < 0.4:
            m = DATE_RE.search(t["text"])
            if m:
                dates.append(f"{int(m.group(1))}-{fix_month(m.group(2))}-{m.group(3)}")
    if dates:
        res["delivery_date"] = dates[0]
    if len(dates) > 1:
        res["receipt_date"] = dates[1]

    # ได้รับเงินจาก
    for t in toks:
        if "ได้รับเงินจาก" in t["text"]:
            name = t["text"].split("ได้รับเงินจาก", 1)[1]
            if not clean_name(name):
                same = [u for u in toks if u is not t and abs(u["y"] - t["y"]) < 0.015 and u["x"] > t["x"] + t["w"] - 0.02]
                name = " ".join(u["text"] for u in sorted(same, key=lambda u: u["x"]))
            res["payer"] = clean_name(name)
            break

    # จำนวนเงิน: ตัวเลขในแถวเดียวกับ "จำนวนเงิน" ก่อนคำว่า บาท
    label = next((t for t in toks if "จำนวนเงิน" in t["text"]), None)
    ly = label["y"] if label else 0.26
    row = [t for t in toks if abs(t["y"] - ly) < 0.025 and t["x"] < 0.5]
    for t in sorted(row, key=lambda t: t["x"]):
        txt = t["text"].replace(" ", "")
        if AMOUNT_RE.match(txt):
            res["amount"] = float(txt.replace(",", ""))
            break
    return res


def write_excel(rows, path):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    wb = Workbook()
    ws = wb.active
    ws.title = "ใบเสร็จรับเงิน"
    ws.append(["หน้า"] + [h for _, h in FIELDS] + ["หมายเหตุ"])
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E78")
        c.alignment = Alignment(horizontal="center", vertical="center")
    warn = PatternFill("solid", fgColor="FFF2CC")
    amount_col = 1 + [k for k, _ in FIELDS].index("amount") + 1
    for page, r in rows:
        missing = [h for k, h in FIELDS if r[k] in ("", None)]
        ws.append([page] + [r[k] for k, _ in FIELDS] + [("ไม่พบ: " + ", ".join(missing)) if missing else ""])
        rr = ws.max_row
        for j, (k, _) in enumerate(FIELDS, start=2):
            if r[k] in ("", None):
                ws.cell(rr, j).fill = warn
        ws.cell(rr, amount_col).number_format = "#,##0.00"
        ws.cell(rr, 1).alignment = Alignment(horizontal="center")
    for i, w in enumerate([8, 18, 22, 16, 28, 14, 26], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    wb.save(path)


def process_pdf(pdf, out=None, progress=None):
    """ประมวลผล PDF ทั้งไฟล์ คืน (path ของ Excel, rows) ; progress(หน้า, จำนวนหน้า, ผลของหน้านั้น)"""
    pdf = Path(pdf)
    out = Path(out) if out else pdf.with_suffix(".xlsx")
    rows = []
    for page, total, img in render_pages(pdf):
        r = parse_page(ocr_image(img))
        rows.append((page, r))
        if progress:
            progress(page, total, r)
    write_excel(rows, out)
    return out, rows


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf"); ap.add_argument("-o", "--output")
    a = ap.parse_args()
    path, _ = process_pdf(a.pdf, a.output, lambda p, n, r: print(f"หน้า {p}/{n}: {r}"))
    print("สร้างไฟล์แล้ว:", path)
