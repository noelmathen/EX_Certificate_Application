#oman.py

#!/usr/bin/env python3
import os, re, glob, logging
import fitz                # PyMuPDF
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font

# —————————————————————————————————————
# CONFIGURATION
# —————————————————————————————————————
SOURCE_DIR = "Oman Certificates"
OUT_XLSX   = "Oman_Certificates.xlsx"
SHEET_NAME = "Oman Certificates"

# 27 data fields  +  'Filename'
FIELDS = [
    "PROJECT DESCRIPTION", "CLIENT TAG", "CIRCUIT ID", "DESCRIPTION", "SYSTEM",
    "MANUFACTURER", "TYPE/MODEL", "SERIAL NUMBER", "Ex PROTECTION", "EPL",
    "CERTIFIED BODY", "Ex CERT No", "DATE INSPECTED", "PROJECT WBS NO",
    "EX INSPECTION TAG No", "AREA CLASSIFICATION", "AREA CLASS", "LAYOUT DWG",
    "LOCATION", "AREA", "GRID REFERENCE", "ACCESS ARRANGEMENT", "IP RATING",
    "REPAIR CATEGORY", "NEXT INSPECTION DUE", "PASS / FAIL",
    "Filename"
]

Y_TOL = 5               # vertical tolerance when matching spans

def normalize_key(s: str) -> str:
    return re.sub(r'[^A-Z0-9]', "", s.upper())

# fast lookup “normalised key” ➜ field name
NORMALIZED_MAP = {normalize_key(f): f for f in FIELDS if f != "Filename"}

# —————————————————————————————————————
# LOGGER
# —————————————————————————————————————
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s", datefmt="%H:%M:%S"
)
logger = logging.getLogger(__name__)

# —————————————————————————————————————
# 1)  EXTRACT SPANS FROM PAGE 1
# —————————————————————————————————————
def extract_spans(pdf_path):
    doc, spans = fitz.open(pdf_path), []
    data = doc[0].get_text("dict")
    for block in data["blocks"]:
        if block.get("type") != 0:  # text only
            continue
        for line in block["lines"]:
            for sp in line["spans"]:
                txt = sp["text"].strip()
                if not txt:
                    continue
                x0, y0, x1, y1 = sp["bbox"]
                spans.append({
                    "text": txt,
                    "kn":   normalize_key(txt),
                    "bbox": (x0, y0, x1, y1),
                    "yp":   (y0 + y1) / 2
                })
    return spans

# —————————————————————————————————————
# 2)  RE-JOIN BROKEN FIELD KEYS
# —————————————————————————————————————
def extend_spans_with_merged(spans):
    rows, merged = {}, []
    for s in spans:                         # bucket into “rows”
        key = int(s["yp"] // Y_TOL)
        rows.setdefault(key, []).append(s)

    for row in rows.values():               # left➜right concatenation
        row.sort(key=lambda s: s["bbox"][0])
        n = len(row)
        for i in range(n):
            text, x0, y0 = row[i]["text"], *row[i]["bbox"][:2]
            for j in range(i + 1, min(i + 5, n)):
                text += row[j]["text"]
                kn = normalize_key(text)
                if kn in NORMALIZED_MAP:
                    x1 = row[j]["bbox"][2]
                    merged.append({
                        "text": text, "kn": kn,
                        "bbox": (x0, y0, x1, row[j]['bbox'][3]),
                        "yp":   (row[i]['yp'] + row[j]['yp']) / 2
                    })
                    break
    return spans + merged

# —————————————————————————————————————
# 3)  PARSE ONE PDF
# —————————————————————————————————————
def parse_certificate(pdf_path, seq_no):
    spans  = extend_spans_with_merged(extract_spans(pdf_path))
    keys, vals = [], []
    for s in spans:
        (keys if s["kn"] in NORMALIZED_MAP else vals).append(s)

    meta = {"Sl. No": seq_no, "Filename": os.path.basename(pdf_path)}

    for k in keys:                          # nearest value on same line
        fld, x_end, y_mid = NORMALIZED_MAP[k["kn"]], k["bbox"][2], k["yp"]
        cands = [
            (v["bbox"][0] - x_end, v)
            for v in vals
            if abs(v["yp"] - y_mid) <= Y_TOL and v["bbox"][0] > x_end
        ]
        meta[fld] = cands and min(cands)[1]["text"].strip() or ""

    for f in FIELDS:                        # guarantee every column
        meta.setdefault(f, "")
    missing = [f for f in FIELDS if not meta[f].strip()]
    return meta, missing

# —————————————————————————————————————
# 4)  DRIVER
# —————————————————————————————————————
def main():
    pdfs = sorted(glob.glob(os.path.join(SOURCE_DIR, "*.pdf")))
    if not pdfs:
        logger.error(f"No PDFs found in '{SOURCE_DIR}'")
        return

    records, errors = [], []
    for idx, pdf in enumerate(pdfs, 1):
        name = os.path.basename(pdf)
        logger.info(f"🔍 ({idx}/{len(pdfs)}) {name}")
        try:
            meta, missing = parse_certificate(pdf, idx)
        except Exception as e:
            errors.append(f"{name}: EXCEPTION – {e}")
            meta = {"Sl. No": idx, "Filename": name, **{f: "" for f in FIELDS}}
            records.append(meta)
            continue

        if missing:                         # leave blank row for manual fix
            errors.append(f"{name}: missing {len(missing)}  ➜ {missing}")
            blank = {"Sl. No": idx, "Filename": name, **{f: "" for f in FIELDS}}
            records.append(blank)
        else:
            records.append(meta)

    # — WRITE EXCEL —
    df = pd.DataFrame(records, columns=["Sl. No"] + FIELDS)
    df.to_excel(OUT_XLSX, sheet_name=SHEET_NAME, index=False)

    wb, ws = load_workbook(OUT_XLSX), None
    ws = wb[SHEET_NAME]
    for c in ws[1]:             # bold header
        c.font = Font(bold=True)
    for col in ws.columns:      # auto-width
        ws.column_dimensions[col[0].column_letter].width = \
            max((len(str(c.value)) for c in col), default=0) + 2
    wb.save(OUT_XLSX)

    logger.info(f"✅ Wrote {len(records)} rows to '{OUT_XLSX}'")
    if errors:
        logger.warning("⚠️ Some PDFs had issues:")
        for e in errors:
            logger.warning(" - %s", e)

if __name__ == "__main__":
    main()
