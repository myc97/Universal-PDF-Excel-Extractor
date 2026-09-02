import gc
import io
import re
import tempfile
import time
from pathlib import Path

import fitz
import pandas as pd
import pdfplumber
import pytesseract
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font
from pdf2image import convert_from_path

st.set_page_config(page_title="NEET Universal PDF Extractor", page_icon="📄", layout="wide")

COLUMNS = ["Sr. No.", "AIR", "NEET Roll No.", "CET Form No.", "Name", "G", "Cat.", "Quota", "Code", "College", "PDF Page"]
POS = {"name": 39, "gender": 73, "cat": 76, "quota": 88, "code": 115, "college": 120}
MAX_BATCH = 150


def clean(v):
    if v is None:
        return ""
    return re.sub(r"\s+", " ", str(v).replace("\xa0", " ").replace("\r", " ").replace("\n", " ")).strip()


def header_line(s):
    s = re.sub(r"[^a-z]+", " ", s.lower())
    return sum(x in s for x in ("sr no", "air", "neet roll", "cet form", "name", "quota", "code", "college")) >= 3


def data_line(s):
    return re.match(r"^\s*\d+\s+\d+\s+\d{8,12}\s+\d{7,12}\s+", s or "") is not None


def college_heading(s):
    s = clean(s)
    if not s or data_line(s) or header_line(s) or len(s) < 6:
        return False
    low = s.lower()
    if low in {"college list", "college wise", "allotment list", "seat allotment", "merit list", "candidate list"}:
        return False
    cues = ("medical college", "college", "institute", "university", "hospital", "society", "academy", "medical sciences")
    return any(x in low for x in cues) or bool(re.match(r"^\d{4}\s*[-:]\s*", s))


def find_heading(lines, header_idx):
    if header_idx is None:
        header_idx = len(lines)
    candidates = [clean(x) for x in lines[max(0, header_idx - 6):header_idx] if college_heading(x)]
    return candidates[-1] if candidates else ""


def first4(line):
    return re.match(r"^\s*(?P<sr>\d+)\s+(?P<air>\d+)\s+(?P<roll>\d{8,12})\s+(?P<form>\d{7,12})\s+", line)


def fixed_row(line, page_no, college=""):
    if not data_line(line):
        return None
    m = first4(line)
    if not m:
        return None
    raw = line.rstrip("\r\n")
    if len(raw) < POS["college"]:
        raw = raw.ljust(POS["college"])
    name = raw[POS["name"]:POS["gender"]]
    gender = raw[POS["gender"]:POS["cat"]]
    cat = raw[POS["cat"]:POS["quota"]]
    quota = raw[POS["quota"]:POS["code"]]
    code = raw[POS["code"]:POS["college"]].rstrip(":").strip()
    row_college = raw[POS["college"]:].strip()
    code_m = re.search(r"\b\d{4}\b", code)
    if not code_m:
        return None
    code = code_m.group(0)
    return {
        "Sr. No.": m.group("sr"), "AIR": m.group("air"), "NEET Roll No.": m.group("roll"),
        "CET Form No.": m.group("form"), "Name": clean(name), "G": clean(gender),
        "Cat.": clean(cat), "Quota": clean(quota), "Code": code,
        "College": clean(row_college) or clean(college), "PDF Page": page_no,
    }


def fitz_lines(page):
    out = []
    for block in page.get_text("blocks", sort=True):
        for line in block[4].splitlines():
            if line.strip():
                out.append(line.rstrip("\r\n"))
    return out


def table_records(table, page_no, college):
    rows = [[clean(x) for x in r] for r in (table or []) if r and any(clean(x) for x in r)]
    if not rows:
        return []
    hidx, best = -1, 0
    aliases = {"sr", "air", "roll", "form", "name", "gender", "cat", "quota", "code", "college"}
    def norm(x):
        x = re.sub(r"[^a-z0-9]+", " ", clean(x).lower()).strip()
        mp = {"sr no":"sr", "serial no":"sr", "serial number":"sr", "s no":"sr", "neet roll no":"roll", "neet roll number":"roll", "roll no":"roll", "cet form no":"form", "cet form number":"form", "form no":"form", "candidate name":"name", "category":"cat", "institute":"college", "institute name":"college", "college code":"code"}
        return mp.get(x, x)
    for i, r in enumerate(rows[:8]):
        score = sum(norm(x) in aliases for x in r)
        if score > best:
            best, hidx = score, i
    if hidx < 0 or best < 3:
        return []
    mapping = {}
    for i, x in enumerate(rows[hidx]):
        n = norm(x)
        if n in aliases and n not in mapping:
            mapping[n] = i
    required = {"sr", "air", "roll", "form", "name", "quota", "code"}
    if not required.issubset(mapping):
        return []
    out = []
    for r in rows[hidx + 1:]:
        def val(k):
            i = mapping.get(k)
            return r[i] if i is not None and i < len(r) else ""
        if not re.fullmatch(r"\d+", val("sr")) or not re.fullmatch(r"\d+", val("air")):
            continue
        if not re.fullmatch(r"\d{8,12}", val("roll")) or not re.fullmatch(r"\d{7,12}", val("form")):
            continue
        cm = re.search(r"\b\d{4}\b", val("code"))
        if not cm:
            continue
        out.append({"Sr. No.":val("sr"),"AIR":val("air"),"NEET Roll No.":val("roll"),"CET Form No.":val("form"),"Name":val("name"),"G":val("gender"),"Cat.":val("cat"),"Quota":val("quota"),"Code":cm.group(),"College":clean(val("college")) or clean(college),"PDF Page":page_no})
    return out


def searchable_page(page, page_no, current_college):
    lines = fitz_lines(page)
    hi = next((i for i, x in enumerate(lines) if header_line(x)), None)
    heading = find_heading(lines, hi)
    if heading:
        current_college = heading
    fixed = [r for line in lines if (r := fixed_row(line, page_no, current_college))]
    if fixed:
        return fixed, fixed[-1].get("College") or current_college
    recs = []
    try:
        for table in page.extract_tables() or []:
            recs.extend(table_records(table, page_no, current_college))
    except Exception:
        pass
    if recs and not current_college:
        current_college = next((r["College"] for r in recs if r["College"]), current_college)
    return recs, current_college


def ocr_page(pdf_path, page_no, dpi, current_college):
    try:
        imgs = convert_from_path(pdf_path, dpi=dpi, first_page=page_no, last_page=page_no, use_cropbox=True, thread_count=1)
        if not imgs:
            return [], current_college
        img = imgs[0]
        text = pytesseract.image_to_string(img, lang="eng", config="--psm 6")
        lines = [x for x in text.splitlines() if x.strip()]
        hi = next((i for i, x in enumerate(lines) if header_line(x)), None)
        heading = find_heading(lines, hi)
        if heading:
            current_college = heading
        recs = [r for line in lines if (r := fixed_row(line, page_no, current_college))]
        if not recs:
            # Conservative OCR fallback: preserve the first four fields and locate a 4-digit code.
            for line in lines:
                m = first4(line)
                if not m:
                    continue
                tail = line[m.end():]
                cm = re.search(r"\b\d{4}\b", tail)
                if not cm:
                    continue
                before, after = tail[:cm.start()], tail[cm.end():]
                parts = [clean(x) for x in re.split(r"\s{2,}", before) if clean(x)]
                if len(parts) >= 4:
                    recs.append({"Sr. No.":m.group("sr"),"AIR":m.group("air"),"NEET Roll No.":m.group("roll"),"CET Form No.":m.group("form"),"Name":parts[0],"G":parts[-3],"Cat.":parts[-2],"Quota":parts[-1],"Code":cm.group(),"College":clean(after) or current_college,"PDF Page":page_no})
        if recs:
            current_college = next((r["College"] for r in recs if r["College"]), current_college)
        return recs, current_college
    except Exception:
        return [], current_college
    finally:
        try: img.close()
        except Exception: pass
        gc.collect()


def detect_type(path):
    try:
        with fitz.open(path) as doc:
            text = "".join((doc[i].get_text("text") or "") for i in range(min(3, len(doc))))
        return "searchable" if len(text.strip()) > 30 else "scanned"
    except Exception:
        return "scanned"


def normalize(records):
    if not records:
        return None, 0
    df = pd.DataFrame(records)
    for c in COLUMNS:
        if c not in df.columns: df[c] = ""
    df = df[COLUMNS].copy().map(clean)
    # Fill missing college by code from rows where it was detected.
    code_map = {}
    for _, r in df.iterrows():
        if r["Code"] and r["College"]: code_map.setdefault(r["Code"], r["College"])
    df.loc[df["College"].eq(""), "College"] = df.loc[df["College"].eq(""), "Code"].map(code_map).fillna("")
    before = len(df)
    df = df.drop_duplicates(subset=[c for c in COLUMNS if c != "PDF Page"], keep="first")
    dup = before - len(df)
    df["_p"] = pd.to_numeric(df["PDF Page"], errors="coerce")
    df["_s"] = pd.to_numeric(df["Sr. No."], errors="coerce")
    return df.sort_values(["_p", "_s"], na_position="last").drop(columns=["_p", "_s"]).reset_index(drop=True), dup


def excel_bytes(df, stats):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        limit = 1_048_000
        for n, start in enumerate(range(0, len(df), limit), 1):
            df.iloc[start:start+limit].to_excel(writer, index=False, sheet_name="Extracted_Data" if n == 1 else f"Data_{n}")
        pd.DataFrame([stats]).to_excel(writer, index=False, sheet_name="Extraction_Summary")
    buf.seek(0)
    wb = load_workbook(buf)
    for ws in wb.worksheets:
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")
    ws = wb["Extracted_Data"]
    widths = [12,12,20,20,38,8,14,28,10,55,12]
    for i, w in enumerate(widths, 1): ws.column_dimensions[chr(64+i)].width = w
    out = io.BytesIO(); wb.save(out); return out.getvalue()


def run(path, batch, dpi, progress, status):
    started = time.time(); kind = detect_type(path)
    with fitz.open(path) as doc: total = len(doc)
    current_college = ""; records = []; errors = []
    batches = [list(range(i, min(i+batch, total))) for i in range(0, total, batch)]
    status.info(f"Detected **{kind.title()} PDF** | Pages: **{total:,}** | Batches: **{len(batches):,}**")
    for bi, indexes in enumerate(batches, 1):
        t = time.time()
        if kind == "searchable":
            try:
                with pdfplumber.open(path) as pdf:
                    for idx in indexes:
                        try:
                            recs, current_college = searchable_page(pdf.pages[idx], idx+1, current_college); records.extend(recs)
                        except Exception as e: errors.append(f"Page {idx+1}: {e}")
            except Exception as e: errors.append(f"Batch {bi}: {e}")
        else:
            for idx in indexes:
                recs, current_college = ocr_page(path, idx+1, dpi, current_college); records.extend(recs)
        progress.progress(bi/len(batches)); elapsed=time.time()-started; eta=(elapsed/bi)*(len(batches)-bi)
        status.info(f"Batch **{bi}/{len(batches)}** | Pages **{indexes[0]+1}-{indexes[-1]+1}** | Rows: **{len(records):,}** | ETA: **{eta:.0f}s** | Batch: **{time.time()-t:.1f}s**")
        gc.collect()
    df, dup = normalize(records)
    stats={"PDF Type":kind,"Pages Processed":total,"Batches":len(batches),"Rows Extracted":0 if df is None else len(df),"Duplicates Removed":dup,"Errors":len(errors),"Time Seconds":round(time.time()-started,2)}
    return df, stats, errors


st.title("📄 NEET Universal PDF → Excel Extractor")
st.caption("Table extraction + fixed-column alignment + college context + OCR | Streamlit Cloud safe")
with st.sidebar:
    st.header("⚙️ Settings")
    batch = st.number_input("Batch Size", 1, MAX_BATCH, MAX_BATCH, 10)
    dpi = st.slider("OCR DPI", 120, 250, 180, 10)
    st.subheader("Output columns")
    st.code("\n".join(COLUMNS))

uploaded = st.file_uploader("Upload PDF", type=["pdf"])
if uploaded:
    st.success(f"Loaded: **{uploaded.name}** ({uploaded.size/(1024*1024):.2f} MB)")
    if st.button("🚀 Start Extraction", type="primary", use_container_width=True):
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
        tmp.write(uploaded.getbuffer()); tmp.close()
        progress=st.progress(0); status=st.empty()
        try:
            with st.spinner("Processing PDF..."):
                df, stats, errors = run(tmp.name, int(batch), int(dpi), progress, status)
            if df is None or df.empty:
                st.error("No valid rows were detected. This PDF may use a different row structure or OCR may need a higher DPI.")
            else:
                st.success(f"🎉 Completed — **{len(df):,} rows** extracted. Duplicates removed: **{stats['Duplicates Removed']:,}**.")
                a,b,c,d=st.columns(4); a.metric("Pages",f"{stats['Pages Processed']:,}"); b.metric("Rows",f"{len(df):,}"); c.metric("Duplicates",f"{stats['Duplicates Removed']:,}"); d.metric("Errors",f"{stats['Errors']:,}")
                st.subheader("Preview"); st.dataframe(df.head(100), use_container_width=True, height=500)
                data=excel_bytes(df, stats)
                st.download_button("⬇️ Download Excel", data=data, file_name=f"{Path(uploaded.name).stem}_Extracted.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                if errors:
                    with st.expander(f"⚠️ Processing errors ({len(errors)})"): st.text("\n".join(errors[:200]))
        finally:
            try: Path(tmp.name).unlink(missing_ok=True)
            except Exception: pass
            gc.collect()
else:
    st.info("Upload a NEET counselling/allotment PDF to begin.")
