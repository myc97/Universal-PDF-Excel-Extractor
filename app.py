from pathlib import Path
import py_compile, zipfile, textwrap

app = r'''import gc
import io
import re
import time
from pathlib import Path

import fitz  # PyMuPDF
import numpy as np
import pandas as pd
import pdfplumber
import pytesseract
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font
from pdf2image import convert_from_path

st.set_page_config(page_title="NEET Universal PDF Extractor", page_icon="📄", layout="wide")

TARGET_COLUMNS = [
    "Sr. No.", "AIR", "NEET Roll No.", "CET Form No.", "Name",
    "G", "Cat.", "Quota", "Code", "College", "PDF Page"
]

# Based on the fixed-alignment extractor supplied by the user.
DEFAULT_POSITIONS = {
    "name": 39,
    "gender": 73,
    "category": 76,
    "quota": 88,
    "code": 115,
    "college": 120,
}

DEFAULT_BATCH_SIZE = 150
MAX_BATCH_SIZE = 150
DEFAULT_OCR_DPI = 180
MIN_OCR_DPI = 120
MAX_OCR_DPI = 250

GENERIC_HEADING_WORDS = {
    "college list", "college wise", "college-wise", "allotment list",
    "seat allotment", "provisional allotment", "merit list",
    "candidate list", "rank list", "allotment", "page", "sr no",
    "sr. no.", "air", "neet roll no", "cet form no", "name", "quota",
    "code", "college"
}


def clean_cell(value):
    if value is None:
        return ""
    value = str(value).replace("\xa0", " ").replace("\n", " ").replace("\r", " ")
    return re.sub(r"\s+", " ", value).strip()


def clean_field(value):
    if value is None:
        return ""
    value = str(value).replace("\xa0", " ")
    return re.sub(r"[ \t]+", " ", value).strip()


def normalize_header(value):
    s = clean_cell(value).lower()
    s = re.sub(r"[^a-z0-9]+", " ", s).strip()
    aliases = {
        "sr": "sr", "sr no": "sr", "sr number": "sr", "serial no": "sr",
        "serial number": "sr", "s no": "sr",
        "air": "air", "all india rank": "air",
        "neet roll no": "roll", "neet roll number": "roll", "roll no": "roll",
        "roll number": "roll",
        "cet form no": "form", "cet form number": "form", "form no": "form",
        "form number": "form",
        "name": "name", "candidate name": "name",
        "g": "gender", "gender": "gender",
        "cat": "cat", "category": "cat",
        "quota": "quota",
        "code": "code", "college code": "code",
        "college": "college", "institute": "college", "institute name": "college",
    }
    return aliases.get(s, s)


def looks_like_header(text):
    n = normalize_header(text)
    return n in {"sr", "air", "roll", "form", "name", "gender", "cat", "quota", "code", "college"}


def is_header_line(line):
    normalized = re.sub(r"[^a-z]+", " ", line.lower())
    hits = sum(x in normalized for x in [
        "sr no", "air", "neet roll", "cet form", "name", "quota", "code", "college"
    ])
    return hits >= 3


def is_data_row(line):
    return re.match(
        r"^\s*\d+\s+\d+\s+\d{8,12}\s+\d{7,12}\s+", line
    ) is not None


def clean_code(value):
    value = clean_field(value).rstrip(":").strip()
    m = re.search(r"\b(\d{4})\b", value)
    return m.group(1) if m else ""


def looks_like_college(line):
    s = clean_cell(line)
    if not s or len(s) < 5 or is_data_row(s) or is_header_line(s):
        return False
    low = s.lower()
    if low in GENERIC_HEADING_WORDS:
        return False
    if re.fullmatch(r"[\d\s\-/:.()]+", s):
        return False
    # Strong institute/college cues; also allow common NEET institute naming patterns.
    cues = (
        "college", "medical college", "institute", "university",
        "hospital", "society", "academy", "school of", "government medical",
        "memorial", "medical sciences", "medical institute"
    )
    if any(cue in low for cue in cues):
        return True
    # Code-like heading: e.g. "1103 - ABC Medical..."
    if re.match(r"^\s*\d{4}\s*[-:]\s*", s):
        return True
    return False


def detect_heading_context(lines, header_index=None):
    """Find a likely college heading immediately before the table/header."""
    if header_index is None:
        header_index = next((i for i, x in enumerate(lines) if is_header_line(x)), len(lines))
    candidates = []
    for line in lines[max(0, header_index - 5):header_index]:
        s = clean_cell(line)
        if not s:
            continue
        if looks_like_college(s):
            candidates.append(s)
    return candidates[-1] if candidates else ""


def extract_first_four(line):
    m = re.match(
        r"^\s*(?P<sr>\d+)\s+(?P<air>\d+)\s+"
        r"(?P<roll>\d{8,12})\s+(?P<form>\d{7,12})\s+",
        line
    )
    return m


def extract_fixed_row(line, page_number, positions=None, fallback_college=""):
    positions = positions or DEFAULT_POSITIONS
    if not is_data_row(line):
        return None

    first_four = extract_first_four(line)
    if not first_four:
        return None

    # IMPORTANT: slice raw aligned text before collapsing spaces.
    end_needed = positions["college"]
    raw = line.rstrip("\r\n")
    if len(raw) < end_needed:
        raw = raw.ljust(end_needed)

    sr = first_four.group("sr")
    air = first_four.group("air")
    roll = first_four.group("roll")
    form = first_four.group("form")

    name = raw[positions["name"]:positions["gender"]]
    gender = raw[positions["gender"]:positions["category"]]
    category = raw[positions["category"]:positions["quota"]]
    quota = raw[positions["quota"]:positions["code"]]
    code = raw[positions["code"]:positions["college"]]
    college = raw[positions["college"]:]

    code = clean_code(code)
    if not code:
        return None

    college = clean_field(college) or clean_field(fallback_college)

    return {
        "Sr. No.": sr,
        "AIR": air,
        "NEET Roll No.": roll,
        "CET Form No.": form,
        "Name": clean_field(name),
        "G": clean_field(gender),
        "Cat.": clean_field(category),
        "Quota": clean_field(quota),
        "Code": code,
        "College": college,
        "PDF Page": page_number,
    }


def get_page_lines_fitz(page):
    lines = []
    try:
        blocks = page.get_text("blocks", sort=True)
        for block in blocks:
            text = block[4] or ""
            for line in text.splitlines():
                line = line.rstrip("\r\n")
                if line.strip():
                    lines.append(line)
    except Exception:
        pass
    return lines


def detect_pdf_type(pdf_path):
    try:
        with fitz.open(pdf_path) as doc:
            sample = min(3, len(doc))
            chars = 0
            text_pages = 0
            for i in range(sample):
                text = doc[i].get_text("text") or ""
                if text.strip():
                    text_pages += 1
                    chars += len(text.strip())
            return "searchable" if text_pages and chars > 30 else "scanned"
    except Exception:
        return "scanned"


def table_to_records(table, page_number, fallback_college=""):
    if not table:
        return []

    rows = [[clean_cell(c) for c in row] for row in table if row]
    rows = [r for r in rows if any(r)]
    if not rows:
        return []

    # Find the strongest header row.
    header_idx = -1
    best_score = 0
    for i, row in enumerate(rows[:8]):
        score = sum(1 for cell in row if normalize_header(cell) in {
            "sr", "air", "roll", "form", "name", "gender", "cat", "quota", "code", "college"
        })
        if score > best_score:
            best_score, header_idx = score, i

    if header_idx >= 0 and best_score >= 3:
        headers = [normalize_header(x) for x in rows[header_idx]]
        mapping = {}
        for idx, h in enumerate(headers):
            if h in {"sr", "air", "roll", "form", "name", "gender", "cat", "quota", "code", "college"} and h not in mapping:
                mapping[h] = idx

        required = {"sr", "air", "roll", "form", "name", "quota", "code"}
        if required.issubset(mapping):
            records = []
            for row in rows[header_idx + 1:]:
                def val(key):
                    i = mapping.get(key)
                    return row[i] if i is not None and i < len(row) else ""

                sr, air, roll, form = val("sr"), val("air"), val("roll"), val("form")
                if not (re.fullmatch(r"\d+", sr or "") and re.fullmatch(r"\d+", air or "") and
                        re.fullmatch(r"\d{8,12}", roll or "") and re.fullmatch(r"\d{7,12}", form or "")):
                    continue
                code = clean_code(val("code"))
                if not code:
                    continue
                college = clean_field(val("college")) or clean_field(fallback_college)
                records.append({
                    "Sr. No.": sr,
                    "AIR": air,
                    "NEET Roll No.": roll,
                    "CET Form No.": form,
                    "Name": val("name"),
                    "G": val("gender"),
                    "Cat.": val("cat"),
                    "Quota": val("quota"),
                    "Code": code,
                    "College": college,
                    "PDF Page": page_number,
                })
            if records:
                return records

    # Table structure exists but header mapping failed: reconstruct each row and use fixed alignment.
    records = []
    for row in rows:
        joined = " ".join(row)
        record = extract_fixed_row(joined, page_number, fallback_college=fallback_college)
        if record:
            records.append(record)
    return records


def extract_searchable_page(page, page_number, current_college):
    lines = get_page_lines_fitz(page)
    if not lines:
        return [], current_college

    # Update context from headings before processing rows.
    header_index = next((i for i, x in enumerate(lines) if is_header_line(x)), None)
    heading = detect_heading_context(lines, header_index)
    if heading:
        current_college = heading

    records = []

    # First attempt: pdfplumber table extraction.
    try:
        tables = page.extract_tables()
        for table in tables or []:
            recs = table_to_records(table, page_number, current_college)
            records.extend(recs)
    except Exception:
        pass

    # Fixed-alignment parser is authoritative for the user's non-table layout.
    fixed_records = []
    for line in lines:
        record = extract_fixed_row(line, page_number, fallback_college=current_college)
        if record:
            fixed_records.append(record)

    # Prefer fixed records when present because they preserve Quota character positions.
    if fixed_records:
        records = fixed_records
    else:
        # Generic multi-space fallback only when no fixed records were found.
        for line in lines:
            if is_data_row(line):
                continue
            parts = [clean_cell(x) for x in re.split(r"\s{2,}", line) if clean_cell(x)]
            if len(parts) >= 8:
                pass

    # If rows themselves contain a college, update state for subsequent pages.
    for r in records:
        if r.get("College"):
            current_college = r["College"]

    return records, current_college


def ocr_page(pdf_path, page_number, dpi, current_college):
    image = None
    try:
        images = convert_from_path(
            pdf_path, dpi=dpi, first_page=page_number, last_page=page_number,
            use_cropbox=True, thread_count=1
        )
        if not images:
            return [], current_college

        image = images[0]
        text = pytesseract.image_to_string(image, lang="eng", config="--psm 6")
        lines = [x.rstrip("\r\n") for x in text.splitlines() if x.strip()]
        if not lines:
            return [], current_college

        header_index = next((i for i, x in enumerate(lines) if is_header_line(x)), None)
        heading = detect_heading_context(lines, header_index)
        if heading:
            current_college = heading

        records = []
        for line in lines:
            record = extract_fixed_row(line, page_number, fallback_college=current_college)
            if record:
                records.append(record)

        # OCR spacing can be imperfect. A conservative token fallback handles rows
        # that have recognizable columns but do not preserve exact character positions.
        if not records:
            for line in lines:
                m = extract_first_four(line)
                if not m:
                    continue
                tail = line[m.end():]
                parts = [clean_cell(x) for x in re.split(r"\s{2,}", tail) if clean_cell(x)]
                if len(parts) >= 5:
                    # Only use this fallback if the quota/code can be identified safely.
                    code_match = re.search(r"\b\d{4}\b", tail)
                    if not code_match:
                        continue
                    before_code = tail[:code_match.start()].strip()
                    after_code = tail[code_match.end():].strip()
                    fields = re.split(r"\s{2,}", before_code)
                    if len(fields) >= 4:
                        name, gender, cat = fields[0], fields[-3], fields[-2]
                        quota = fields[-1]
                        college = after_code or current_college
                        records.append({
                            "Sr. No.": m.group("sr"), "AIR": m.group("air"),
                            "NEET Roll No.": m.group("roll"), "CET Form No.": m.group("form"),
                            "Name": clean_field(name), "G": clean_field(gender),
                            "Cat.": clean_field(cat), "Quota": clean_field(quota),
                            "Code": code_match.group(), "College": clean_field(college),
                            "PDF Page": page_number,
                        })

        for r in records:
            if r.get("College"):
                current_college = r["College"]

        return records, current_college
    except Exception:
        return [], current_college
    finally:
        try:
            if image is not None:
                image.close()
        except Exception:
            pass
        gc.collect()


def create_batches(total_pages, batch_size):
    batch_size = max(1, min(int(batch_size), MAX_BATCH_SIZE))
    return [list(range(i, min(i + batch_size, total_pages))) for i in range(0, total_pages, batch_size)]


def normalize_output(records):
    if not records:
        return None, 0

    df = pd.DataFrame(records)
    for col in TARGET_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    df = df[TARGET_COLUMNS].copy()

    for col in TARGET_COLUMNS:
        df[col] = df[col].map(clean_cell)

    # Fill missing college using same code where a reliable college was found.
    code_to_college = {}
    for _, row in df.iterrows():
        code = row["Code"]
        college = row["College"]
        if code and college:
            code_to_college.setdefault(code, college)

    if code_to_college:
        mask = df["College"].eq("")
        df.loc[mask, "College"] = df.loc[mask, "Code"].map(code_to_college).fillna("")

    before = len(df)
    df = df.drop_duplicates(
        subset=[c for c in TARGET_COLUMNS if c != "PDF Page"],
        keep="first"
    ).copy()
    duplicates_removed = before - len(df)

    df["_page"] = pd.to_numeric(df["PDF Page"], errors="coerce")
    df["_sr"] = pd.to_numeric(df["Sr. No."], errors="coerce")
    df = df.sort_values(["_page", "_sr"], na_position="last").drop(columns=["_page", "_sr"])
    return df.reset_index(drop=True), duplicates_removed


def dataframe_to_excel(df, stats):
    output = io.BytesIO()
    max_rows = 1_048_000
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        if len(df) <= max_rows:
            df.to_excel(writer, index=False, sheet_name="Extracted Data")
        else:
            for n, start in enumerate(range(0, len(df), max_rows), 1):
                df.iloc[start:start + max_rows].to_excel(
                    writer, index=False, sheet_name=f"Data_{n}"
                )
        summary = pd.DataFrame([stats])
        summary.to_excel(writer, index=False, sheet_name="Extraction Summary")

    output.seek(0)
    wb = load_workbook(output)
    for ws in wb.worksheets:
        ws.freeze_panes = "A2"
        if ws.max_row:
            for cell in ws[1]:
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal="center", vertical="center")
            ws.auto_filter.ref = ws.dimensions

    if "Extracted Data" in wb.sheetnames:
        ws = wb["Extracted Data"]
        widths = {"A": 12, "B": 12, "C": 20, "D": 20, "E": 38,
                  "F": 8, "G": 14, "H": 28, "I": 10, "J": 50, "K": 12}
        for col, width in widths.items():
            ws.column_dimensions[col].width = width
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(vertical="center")

    final = io.BytesIO()
    wb.save(final)
    final.seek(0)
    return final.getvalue()


def run_extraction(pdf_path, batch_size, dpi, progress, status):
    started = time.time()
    pdf_type = detect_pdf_type(pdf_path)

    with fitz.open(pdf_path) as doc:
        total_pages = len(doc)

    batches = create_batches(total_pages, batch_size)
    records = []
    errors = []
    current_college = ""

    status.info(
        f"Detected **{'Searchable PDF' if pdf_type == 'searchable' else 'Scanned PDF / OCR'}** | "
        f"Pages: **{total_pages:,}** | Batches: **{len(batches):,}**"
    )

    for batch_no, page_indexes in enumerate(batches, 1):
        batch_started = time.time()
        try:
            if pdf_type == "searchable":
                with pdfplumber.open(pdf_path) as pdf:
                    for idx in page_indexes:
                        page_no = idx + 1
                        try:
                            page_records, current_college = extract_searchable_page(
                                page=page, page_number=page_no, current_college=current_college
                            )
                            records.extend(page_records)
                        except Exception as exc:
                            errors.append(f"Page {page_no}: {exc}")
            else:
                for idx in page_indexes:
                    page_no = idx + 1
                    page_records, current_college = ocr_page(
                        pdf_path, page_no, dpi, current_college
                    )
                    records.extend(page_records)
        except Exception as exc:
            errors.append(f"Batch {batch_no}: {exc}")

        progress.progress(batch_no / max(1, len(batches)))
        elapsed = time.time() - started
        avg = elapsed / batch_no
        eta = avg * (len(batches) - batch_no)
        eta_text = f"{eta:.0f}s" if eta < 60 else f"{eta / 60:.1f} min"
        status.info(
            f"Batch **{batch_no}/{len(batches)}** complete | "
            f"Pages **{page_indexes[0] + 1}–{page_indexes[-1] + 1}** | "
            f"Rows found: **{len(records):,}** | ETA: **{eta_text}** | "
            f"Batch time: **{time.time() - batch_started:.1f}s**"
        )
        gc.collect()

    df, duplicates_removed = normalize_output(records)
    total_time = time.time() - started
    stats = {
        "PDF Type": pdf_type,
        "Pages Processed": total_pages,
        "Batches": len(batches),
        "Rows Extracted": 0 if df is None else len(df),
        "Duplicates Removed": duplicates_removed,
        "Parsing Errors": len(errors),
        "Time (seconds)": round(total_time, 2),
    }
    return df, stats, errors


st.title("📄 NEET Universal PDF → Excel Extractor")
st.caption(
    "Table extraction + fixed-column alignment + college context + OCR | "
    "Cloud-safe page-by-page processing"
)

with st.sidebar:
    st.header("⚙️ Settings")
    batch_size = st.number_input(
        "Batch Size", min_value=1, max_value=150, value=150, step=10,
        help="Maximum 150 pages. 150 is recommended for Streamlit Cloud."
    )
    dpi = st.slider(
        "OCR DPI", min_value=MIN_OCR_DPI, max_value=MAX_OCR_DPI,
        value=DEFAULT_OCR_DPI, step=10
    )
    st.markdown("**Output columns**")
    st.code("\n".join(TARGET_COLUMNS), language="text")

uploaded = st.file_uploader("Upload PDF", type=["pdf"])

if uploaded:
    st.success(f"Loaded: **{uploaded.name}**")
    if st.button("🚀 Start Extraction", type="primary", use_container_width=True):
        suffix = Path(uploaded.name).suffix or ".pdf"
        with open(Path.cwd() / f"_input_{int(time.time())}{suffix}", "wb") as f:
            f.write(uploaded.getbuffer())
            pdf_path = f.name

        progress = st.progress(0)
        status = st.empty()

        try:
            df, stats, errors = run_extraction(
                pdf_path, int(batch_size), int(dpi), progress, status
            )

            if df is None or df.empty:
                st.error(
                    "No valid rows were detected. The PDF may use a layout not covered by "
                    "the current row pattern, or OCR quality may be insufficient."
                )
            else:
                st.success(
                    f"🎉 Completed — **{len(df):,} rows** extracted. "
                    f"Duplicates removed: **{stats['Duplicates Removed']:,}**."
                )

                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Pages", f"{stats['Pages Processed']:,}")
                c2.metric("Rows", f"{len(df):,}")
                c3.metric("Duplicates", f"{stats['Duplicates Removed']:,}")
                c4.metric("Errors", f"{stats['Parsing Errors']:,}")

                st.subheader("Preview")
                st.dataframe(df.head(100), use_container_width=True, height=500)

                excel_bytes = dataframe_to_excel(df, stats)
                out_name = f"{Path(uploaded.name).stem}_Extracted.xlsx"
                st.download_button(
                    "⬇️ Download Excel",
                    data=excel_bytes,
                    file_name=out_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

                if errors:
                    with st.expander(f"⚠️ {len(errors)} processing errors"):
                        st.text("\n".join(errors[:200]))
        finally:
            try:
                Path(pdf_path).unlink(missing_ok=True)
            except Exception:
                pass
            gc.collect()
else:
    st.info(
        "Upload a NEET counselling/allotment PDF. Searchable PDFs use table extraction "
        "and fixed-position parsing; scanned PDFs use Tesseract OCR."
    )
