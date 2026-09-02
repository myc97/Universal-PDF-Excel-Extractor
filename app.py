import streamlit as st
import fitz
import pdfplumber
import pandas as pd
import numpy as np
import re
import time
import tempfile
import shutil
import gc
import os
from io import BytesIO

from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment

from pdf2image import convert_from_path
import pytesseract


# ============================================================
# STREAMLIT CONFIG
# ============================================================

st.set_page_config(
    page_title="Universal PDF Fast Extractor",
    page_icon="📄",
    layout="wide"
)


# ============================================================
# CONFIGURATION
# ============================================================

DEFAULT_BATCH_SIZE = 150
MAX_BATCH_SIZE = 150

DEFAULT_OCR_DPI = 180
MIN_OCR_DPI = 120
MAX_OCR_DPI = 250

# ============================================================
# SPECIAL NEET TABLE COLUMNS
# ============================================================

NEET_COLUMNS = [
    "Sr. No.",
    "AIR",
    "NEET Roll No.",
    "CET Form No.",
    "Name",
    "G",
    "Cat.",
    "Quota",
    "Code",
    "College",
    "PDF Page"
]

# ============================================================
# NEET TABLE FIXED POSITIONS
#
# These are based on the alignment in the supplied
# Sell_R1-mbbs.pdf structure.
#
# Quota starts exactly where the PDF header "Quota" starts.
# ============================================================

NAME_START = 39
G_START = 73
CAT_START = 76
QUOTA_START = 88
CODE_START = 115
COLLEGE_START = 120


# ============================================================
# GENERIC COLUMNS
# ============================================================

GENERIC_PAGE_COLUMN = "PDF Page"


# ============================================================
# SYSTEM CHECK
# ============================================================

def check_system():

    status = {
        "tesseract": False,
        "pdfplumber": False,
        "pandas": False,
        "fitz": False,
        "pdf2image": False
    }

    try:
        pytesseract.get_tesseract_version()
        status["tesseract"] = True
    except Exception:
        pass

    try:
        import pdfplumber
        status["pdfplumber"] = True
    except Exception:
        pass

    try:
        import pandas
        status["pandas"] = True
    except Exception:
        pass

    try:
        import fitz
        status["fitz"] = True
    except Exception:
        pass

    try:
        import pdf2image
        status["pdf2image"] = True
    except Exception:
        pass

    return status


# ============================================================
# TEXT CLEANING
# ============================================================

def clean_field(value):

    if value is None:
        return ""

    value = str(value)

    value = value.replace("\xa0", " ")
    value = value.replace("\r", " ")
    value = value.replace("\n", " ")

    value = re.sub(r"[ \t]+", " ", value)

    return value.strip()


# ============================================================
# CLEAN DATAFRAME
# ============================================================

def clean_dataframe(df):

    if df is None:
        return None

    if df.empty:
        return df

    df = df.copy()

    for column in df.columns:
        df[column] = df[column].apply(clean_field)

    df = df.replace("", np.nan)

    df = df.dropna(
        axis=0,
        how="all"
    )

    df = df.fillna("")

    return df.reset_index(drop=True)


# ============================================================
# GET PAGE COUNT
# ============================================================

def get_total_pages(pdf_path):

    try:

        with fitz.open(pdf_path) as doc:
            return len(doc)

    except Exception:

        with pdfplumber.open(pdf_path) as pdf:
            return len(pdf.pages)


# ============================================================
# CREATE BATCHES
# ============================================================

def create_batches(
    total_pages,
    batch_size=150
):

    batch_size = min(
        int(batch_size),
        MAX_BATCH_SIZE
    )

    batches = []

    for start in range(
        0,
        total_pages,
        batch_size
    ):

        end = min(
            start + batch_size,
            total_pages
        )

        batches.append(
            list(range(start, end))
        )

    return batches


# ============================================================
# PDF TYPE DETECTION
# ============================================================

def detect_pdf_type(pdf_path):

    try:

        with fitz.open(pdf_path) as doc:

            pages_to_check = min(
                3,
                len(doc)
            )

            total_chars = 0
            text_pages = 0

            for i in range(
                pages_to_check
            ):

                try:

                    text = doc[i].get_text(
                        "text"
                    )

                    if text and text.strip():

                        text_pages += 1
                        total_chars += len(
                            text.strip()
                        )

                except Exception:
                    continue

            if (
                text_pages > 0
                and total_chars > 30
            ):

                return "searchable"

            return "scanned"

    except Exception:

        return "scanned"


# ============================================================
# NEET DATA ROW DETECTION
# ============================================================

def is_neet_data_row(line):

    if not line:
        return False

    pattern = (
        r"^\s*"
        r"\d+\s+"
        r"\d+\s+"
        r"\d{8,12}\s+"
        r"\d{7,12}\s+"
    )

    return (
        re.match(
            pattern,
            line
        )
        is not None
    )


# ============================================================
# NEET HEADER DETECTION
# ============================================================

def looks_like_neet_header(text):

    if not text:
        return False

    normalized = re.sub(
        r"\s+",
        " ",
        text.lower()
    )

    required = [
        "sr",
        "air",
        "neet",
        "roll",
        "cet",
        "name",
        "quota",
        "code",
        "college"
    ]

    score = sum(
        1
        for word in required
        if word in normalized
    )

    return score >= 6


# ============================================================
# NEET STRUCTURE DETECTION
# ============================================================

def detect_neet_structure(pdf_path):

    """
    Detect whether the PDF resembles the supplied
    NEET allotment table.

    We check several first pages instead of relying
    only on page 1.
    """

    try:

        with fitz.open(pdf_path) as doc:

            pages_to_check = min(
                5,
                len(doc)
            )

            header_score = 0
            data_rows = 0

            for i in range(
                pages_to_check
            ):

                try:

                    page = doc[i]

                    text = page.get_text(
                        "text"
                    )

                    if looks_like_neet_header(
                        text
                    ):

                        header_score += 1

                    for line in text.splitlines():

                        if is_neet_data_row(
                            line
                        ):

                            data_rows += 1

                except Exception:
                    continue

            if (
                header_score >= 1
                and data_rows >= 2
            ):

                return True

            return False

    except Exception:

        return False


# ============================================================
# NEET ROW EXTRACTION
# ============================================================

def extract_neet_row(
    line,
    page_number,
    fallback_college=""
):

    if not is_neet_data_row(line):
        return None

    # --------------------------------------------------------
    # FIRST FOUR NUMERIC COLUMNS
    # --------------------------------------------------------

    first_four = re.match(
        r"^\s*"
        r"(?P<sr>\d+)\s+"
        r"(?P<air>\d+)\s+"
        r"(?P<roll>\d{8,12})\s+"
        r"(?P<form>\d{7,12})\s+",
        line
    )

    if not first_four:
        return None

    sr = first_four.group("sr")
    air = first_four.group("air")
    roll = first_four.group("roll")
    form = first_four.group("form")

    # --------------------------------------------------------
    # MAKE LINE LONG ENOUGH
    # --------------------------------------------------------

    if len(line) < COLLEGE_START:

        line = line.ljust(
            COLLEGE_START
        )

    # --------------------------------------------------------
    # FIXED POSITION EXTRACTION
    # --------------------------------------------------------

    name = line[
        NAME_START:G_START
    ]

    gender = line[
        G_START:CAT_START
    ]

    category = line[
        CAT_START:QUOTA_START
    ]

    quota = line[
        QUOTA_START:CODE_START
    ]

    code = line[
        CODE_START:COLLEGE_START
    ]

    college = line[
        COLLEGE_START:
    ]

    # --------------------------------------------------------
    # CLEAN
    # --------------------------------------------------------

    name = clean_field(name)
    gender = clean_field(gender)
    category = clean_field(category)
    quota = clean_field(quota)
    code = clean_field(code)
    college = clean_field(college)

    # --------------------------------------------------------
    # REMOVE COLON
    # --------------------------------------------------------

    code = code.rstrip(":").strip()

    # --------------------------------------------------------
    # COLLEGE FALLBACK
    # --------------------------------------------------------

    if not college and fallback_college:

        college = clean_field(
            fallback_college
        )

    # --------------------------------------------------------
    # VALIDATE CODE
    # --------------------------------------------------------

    if not re.fullmatch(
        r"\d{4}",
        code
    ):

        return None

    return {

        "Sr. No.": sr,

        "AIR": air,

        "NEET Roll No.": roll,

        "CET Form No.": form,

        "Name": name,

        "G": gender,

        "Cat.": category,

        "Quota": quota,

        "Code": code,

        "College": college,

        "PDF Page": page_number
    }


# ============================================================
# FIND COLLEGE / CONTEXT TEXT
# ============================================================

def detect_context_before_table(
    lines
):

    """
    Preserve useful text appearing above
    table headers, especially college names.

    This is intentionally conservative.
    """

    candidates = []

    for line in lines:

        line = clean_field(line)

        if not line:
            continue

        if looks_like_neet_header(line):
            continue

        if is_neet_data_row(line):
            continue

        # Ignore obvious page/header noise
        lower = line.lower()

        if any(
            x in lower
            for x in [
                "page ",
                "sr. no",
                "sr no",
                "rank",
                "neet ug",
                "allotment",
                "government of",
                "medical counselling"
            ]
        ):
            continue

        # College names are usually longer text.
        if len(line) >= 8:
            candidates.append(line)

    if candidates:
        return candidates[-1]

    return ""


# ============================================================
# NEET SEARCHABLE EXTRACTION
# ============================================================

def process_neet_searchable_batch(
    pdf_path,
    page_numbers
):

    records = []
    errors = []

    try:

        with fitz.open(pdf_path) as doc:

            for page_index in page_numbers:

                page_number = page_index + 1

                try:

                    page = doc[
                        page_index
                    ]

                    # ------------------------------------------------
                    # BLOCK EXTRACTION
                    # ------------------------------------------------

                    blocks = page.get_text(
                        "blocks"
                    )

                    page_lines = []

                    for block in blocks:

                        if len(block) < 5:
                            continue

                        block_text = block[4]

                        if not block_text:
                            continue

                        for line in block_text.splitlines():

                            line = line.rstrip(
                                "\n\r"
                            )

                            if line.strip():
                                page_lines.append(
                                    line
                                )

                    # ------------------------------------------------
                    # CONTEXT / COLLEGE
                    # ------------------------------------------------

                    fallback_college = (
                        detect_context_before_table(
                            page_lines
                        )
                    )

                    # ------------------------------------------------
                    # EXTRACT DATA ROWS
                    # ------------------------------------------------

                    for line in page_lines:

                        if not is_neet_data_row(
                            line
                        ):
                            continue

                        try:

                            row = extract_neet_row(
                                line,
                                page_number,
                                fallback_college
                            )

                            if row:

                                records.append(
                                    row
                                )

                            else:

                                errors.append(
                                    f"Page {page_number}: "
                                    f"Could not parse row: "
                                    f"{line}"
                                )

                        except Exception as row_error:

                            errors.append(
                                f"Page {page_number}: "
                                f"{row_error} | "
                                f"{line}"
                            )

                except Exception as page_error:

                    errors.append(
                        f"Page {page_number}: "
                        f"PAGE ERROR - "
                        f"{page_error}"
                    )

    except Exception as batch_error:

        errors.append(
            f"Batch error: {batch_error}"
        )

    return records, errors


# ============================================================
# GENERIC SEARCHABLE TABLE EXTRACTION
# ============================================================

def extract_tables_from_page(page):

    results = []

    try:

        tables = page.extract_tables()

        if tables:

            for table in tables:

                if not table:
                    continue

                try:

                    df = pd.DataFrame(
                        table
                    )

                    if df.empty:
                        continue

                    df = clean_dataframe(
                        df
                    )

                    if (
                        df is not None
                        and not df.empty
                    ):

                        results.append(
                            df
                        )

                except Exception:
                    continue

    except Exception:
        pass

    return results


# ============================================================
# GENERIC SEARCHABLE TEXT EXTRACTION
# ============================================================

def extract_text_generic(
    page,
    page_number
):

    results = []

    try:

        text = page.extract_text()

        if not text or not text.strip():
            return results

        lines = []

        for line in text.splitlines():

            line = clean_field(line)

            if line:
                lines.append(line)

        if not lines:
            return results

        rows = []

        for line in lines:

            # Multiple spaces indicate
            # possible column boundaries.
            parts = re.split(
                r"\s{2,}",
                line
            )

            parts = [
                clean_field(x)
                for x in parts
                if clean_field(x)
            ]

            if not parts:
                continue

            rows.append(parts)

        if not rows:
            return results

        max_columns = max(
            len(row)
            for row in rows
        )

        normalized = []

        for row in rows:

            row = list(row)

            if len(row) < max_columns:

                row.extend(
                    [""] *
                    (
                        max_columns
                        - len(row)
                    )
                )

            elif len(row) > max_columns:

                row = row[
                    :max_columns
                ]

            normalized.append(
                row
            )

        df = pd.DataFrame(
            normalized
        )

        df["PDF Page"] = (
            page_number
        )

        df = clean_dataframe(
            df
        )

        if (
            df is not None
            and not df.empty
        ):

            results.append(
                df
            )

    except Exception:
        pass

    return results


# ============================================================
# GENERIC SEARCHABLE BATCH
# ============================================================

def process_generic_searchable_batch(
    pdf_path,
    page_numbers
):

    batch_results = []

    try:

        with pdfplumber.open(
            pdf_path
        ) as pdf:

            for page_index in page_numbers:

                page_number = (
                    page_index + 1
                )

                try:

                    page = pdf.pages[
                        page_index
                    ]

                    # --------------------------------------------
                    # FIRST TRY REAL TABLE EXTRACTION
                    # --------------------------------------------

                    tables = (
                        extract_tables_from_page(
                            page
                        )
                    )

                    if tables:

                        for df in tables:

                            df = df.copy()

                            if (
                                "PDF Page"
                                not in df.columns
                            ):

                                df[
                                    "PDF Page"
                                ] = page_number

                            batch_results.append(
                                df
                            )

                    else:

                        # ----------------------------------------
                        # TEXT FALLBACK
                        # ----------------------------------------

                        text_results = (
                            extract_text_generic(
                                page,
                                page_number
                            )
                        )

                        batch_results.extend(
                            text_results
                        )

                except Exception:
                    continue

    except Exception:
        pass

    return batch_results


# ============================================================
# OCR LINE RECONSTRUCTION
# ============================================================

def ocr_page_to_dataframe(
    pdf_path,
    page_number,
    dpi
):

    image = None

    try:

        # ----------------------------------------------------
        # ONE PAGE ONLY
        # ----------------------------------------------------

        images = convert_from_path(
            pdf_path,
            dpi=dpi,
            first_page=page_number,
            last_page=page_number,
            use_cropbox=True,
            thread_count=1
        )

        if not images:
            return None

        image = images[0]

        # ----------------------------------------------------
        # OCR DATA
        # ----------------------------------------------------

        data = pytesseract.image_to_data(
            image,
            lang="eng",
            config="--psm 6",
            output_type=pytesseract.Output.DATAFRAME
        )

        if data is None:
            return None

        data = data.dropna(
            subset=["text"]
        )

        data["text"] = data[
            "text"
        ].astype(str).str.strip()

        data = data[
            data["text"] != ""
        ]

        if data.empty:
            return None

        # ----------------------------------------------------
        # GROUP WORDS BY OCR LINE
        # ----------------------------------------------------

        lines = []

        group_columns = [
            "block_num",
            "par_num",
            "line_num"
        ]

        for _, group in data.groupby(
            group_columns,
            sort=False
        ):

            group = group.sort_values(
                "left"
            )

            words = []

            for _, row in group.iterrows():

                text = clean_field(
                    row["text"]
                )

                if not text:
                    continue

                left = int(
                    row["left"]
                )

                words.append(
                    (
                        left,
                        text
                    )
                )

            if not words:
                continue

            # ------------------------------------------------
            # Reconstruct approximate spacing
            # ------------------------------------------------

            reconstructed = ""

            previous_right = None

            for left, text in words:

                if previous_right is None:

                    reconstructed = text

                else:

                    gap = (
                        left
                        - previous_right
                    )

                    # Approximate character
                    # width for OCR reconstruction.
                    spaces = max(
                        1,
                        min(
                            12,
                            int(
                                gap / 8
                            )
                        )
                    )

                    reconstructed += (
                        " "
                        * spaces
                        + text
                    )

                # Approximate width
                previous_right = (
                    left
                    + max(
                        8,
                        len(text) * 8
                    )
                )

            lines.append(
                reconstructed
            )

        if not lines:
            return None

        # ----------------------------------------------------
        # TRY NEET ROW PARSER ON OCR
        # ----------------------------------------------------

        neet_records = []

        for line in lines:

            if is_neet_data_row(
                line
            ):

                row = extract_neet_row(
                    line,
                    page_number
                )

                if row:
                    neet_records.append(
                        row
                    )

        if neet_records:

            return pd.DataFrame(
                neet_records,
                columns=NEET_COLUMNS
            )

        # ----------------------------------------------------
        # GENERIC OCR FALLBACK
        # ----------------------------------------------------

        rows = []

        for line in lines:

            parts = re.split(
                r"\s{2,}",
                line
            )

            parts = [
                clean_field(x)
                for x in parts
                if clean_field(x)
            ]

            if parts:
                rows.append(parts)

        if not rows:
            return None

        max_columns = max(
            len(row)
            for row in rows
        )

        normalized = []

        for row in rows:

            row = list(row)

            if len(row) < max_columns:

                row.extend(
                    [""] *
                    (
                        max_columns
                        - len(row)
                    )
                )

            elif len(row) > max_columns:

                row = row[
                    :max_columns
                ]

            normalized.append(
                row
            )

        df = pd.DataFrame(
            normalized
        )

        df["PDF Page"] = (
            page_number
        )

        return clean_dataframe(
            df
        )

    except Exception:

        return None

    finally:

        try:

            if image is not None:
                image.close()

        except Exception:
            pass

        try:
            del image
        except Exception:
            pass

        try:
            del images
        except Exception:
            pass

        try:
            del data
        except Exception:
            pass

        gc.collect()


# ============================================================
# OCR BATCH
# ============================================================

def process_ocr_batch(
    pdf_path,
    page_numbers,
    dpi
):

    batch_results = []

    for page_index in page_numbers:

        page_number = (
            page_index + 1
        )

        df = ocr_page_to_dataframe(
            pdf_path,
            page_number,
            dpi
        )

        if (
            df is not None
            and not df.empty
        ):

            batch_results.append(
                df
            )

        gc.collect()

    return batch_results


# ============================================================
# COMBINE DATAFRAMES
# ============================================================

def combine_dataframes(
    all_data
):

    if not all_data:
        return None

    valid = []

    for df in all_data:

        if not isinstance(
            df,
            pd.DataFrame
        ):
            continue

        if df.empty:
            continue

        valid.append(
            df.copy()
        )

    if not valid:
        return None

    # --------------------------------------------------------
    # SPECIAL NEET DATA
    # --------------------------------------------------------

    neet_like = all(
        set(
            [
                "Sr. No.",
                "AIR",
                "NEET Roll No.",
                "CET Form No.",
                "Name",
                "G",
                "Cat.",
                "Quota",
                "Code",
                "College",
                "PDF Page"
            ]
        ).issubset(
            set(df.columns)
        )
        for df in valid
    )

    if neet_like:

        combined = pd.concat(
            valid,
            ignore_index=True
        )

        combined = combined[
            NEET_COLUMNS
        ]

        # ----------------------------------------------------
        # Remove exact duplicates
        # ----------------------------------------------------

        before = len(
            combined
        )

        combined = (
            combined
            .drop_duplicates(
                subset=[
                    "Sr. No.",
                    "AIR",
                    "NEET Roll No.",
                    "CET Form No.",
                    "Name",
                    "G",
                    "Cat.",
                    "Quota",
                    "Code",
                    "College"
                ],
                keep="first"
            )
        )

        # ----------------------------------------------------
        # Sort
        # ----------------------------------------------------

        combined["_SrNumeric"] = (
            pd.to_numeric(
                combined["Sr. No."],
                errors="coerce"
            )
        )

        combined = (
            combined
            .sort_values(
                by=[
                    "PDF Page",
                    "_SrNumeric"
                ],
                kind="stable"
            )
            .drop(
                columns=[
                    "_SrNumeric"
                ]
            )
            .reset_index(
                drop=True
            )
        )

        return combined

    # --------------------------------------------------------
    # GENERIC DATA
    # --------------------------------------------------------

    max_columns = max(
        len(df.columns)
        for df in valid
    )

    normalized = []

    for df in valid:

        df = df.copy()

        df.columns = [
            str(column)
            for column in df.columns
        ]

        current_columns = len(
            df.columns
        )

        if current_columns < max_columns:

            for i in range(
                current_columns,
                max_columns
            ):

                df[
                    f"Column_{i + 1}"
                ] = ""

        normalized.append(
            df
        )

    combined = pd.concat(
        normalized,
        ignore_index=True,
        sort=False
    )

    combined = clean_dataframe(
        combined
    )

    return combined


# ============================================================
# EXCEL GENERATION
# ============================================================

def dataframe_to_excel(df):

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        max_excel_rows = 1_048_000

        total_rows = len(df)

        if total_rows <= max_excel_rows:

            sheet_name = (
                "Extracted Data"
            )

            df.to_excel(
                writer,
                index=False,
                sheet_name=sheet_name
            )

        else:

            sheet_number = 1

            for start in range(
                0,
                total_rows,
                max_excel_rows
            ):

                end = min(
                    start + max_excel_rows,
                    total_rows
                )

                chunk = df.iloc[
                    start:end
                ]

                chunk.to_excel(
                    writer,
                    index=False,
                    sheet_name=(
                        f"Data_{sheet_number}"
                    )
                )

                sheet_number += 1

    output.seek(0)

    excel_data = output.getvalue()

    # --------------------------------------------------------
    # FORMAT EXCEL
    # --------------------------------------------------------

    formatted_output = BytesIO()

    try:

        workbook = load_workbook(
            BytesIO(excel_data)
        )

        for worksheet in workbook.worksheets:

            # Header formatting
            for cell in worksheet[1]:

                cell.font = Font(
                    bold=True
                )

                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center"
                )

            worksheet.freeze_panes = "A2"

            worksheet.auto_filter.ref = (
                worksheet.dimensions
            )

            # ------------------------------------------------
            # NEET specific widths
            # ------------------------------------------------

            if (
                "Sr. No."
                in [
                    cell.value
                    for cell in worksheet[1]
                ]
            ):

                widths = {

                    "A": 12,
                    "B": 12,
                    "C": 20,
                    "D": 20,
                    "E": 38,
                    "F": 8,
                    "G": 14,
                    "H": 25,
                    "I": 10,
                    "J": 45,
                    "K": 12
                }

                for column, width in widths.items():

                    worksheet.column_dimensions[
                        column
                    ].width = width

            else:

                # Generic width
                for column_cells in worksheet.columns:

                    column_letter = (
                        column_cells[0]
                        .column_letter
                    )

                    max_length = 0

                    for cell in column_cells[:200]:

                        value = str(
                            cell.value
                            if cell.value
                            is not None
                            else ""
                        )

                        max_length = max(
                            max_length,
                            len(value)
                        )

                    worksheet.column_dimensions[
                        column_letter
                    ].width = min(
                        max(
                            max_length + 2,
                            10
                        ),
                        45
                    )

            # ------------------------------------------------
            # Vertical alignment
            # ------------------------------------------------

            for row in worksheet.iter_rows(
                min_row=2
            ):

                for cell in row:

                    cell.alignment = Alignment(
                        vertical="center"
                    )

        workbook.save(
            formatted_output
        )

        formatted_output.seek(0)

        return formatted_output.getvalue()

    except Exception:

        return excel_data


# ============================================================
# MAIN EXTRACTION
# ============================================================

def run_extraction(
    pdf_path,
    batch_size,
    dpi,
    progress_bar,
    status_text
):

    start_time = time.time()

    # --------------------------------------------------------
    # DETECT PDF TYPE
    # --------------------------------------------------------

    status_text.info(
        "🔍 Detecting PDF type..."
    )

    pdf_type = detect_pdf_type(
        pdf_path
    )

    # --------------------------------------------------------
    # COUNT PAGES
    # --------------------------------------------------------

    status_text.info(
        "📄 Counting PDF pages..."
    )

    total_pages = get_total_pages(
        pdf_path
    )

    if total_pages <= 0:

        raise ValueError(
            "The PDF contains no pages."
        )

    # --------------------------------------------------------
    # DETECT SPECIAL NEET STRUCTURE
    # --------------------------------------------------------

    neet_structure = False

    if pdf_type == "searchable":

        status_text.info(
            "🔎 Detecting PDF table structure..."
        )

        neet_structure = (
            detect_neet_structure(
                pdf_path
            )
        )

    # --------------------------------------------------------
    # CREATE BATCHES
    # --------------------------------------------------------

    batches = create_batches(
        total_pages,
        batch_size
    )

    total_batches = len(
        batches
    )

    # --------------------------------------------------------
    # MODE
    # --------------------------------------------------------

    if pdf_type == "scanned":

        mode_text = (
            "Scanned PDF / OCR"
        )

    elif neet_structure:

        mode_text = (
            "NEET Alotment Table / "
            "Fixed-Position Extraction"
        )

    else:

        mode_text = (
            "Searchable PDF / "
            "Generic Extraction"
        )

    status_text.success(
        f"Detected: **{mode_text}** | "
        f"Pages: **{total_pages:,}** | "
        f"Batches: **{total_batches}**"
    )

    # --------------------------------------------------------
    # DATA STORAGE
    # --------------------------------------------------------

    all_data = []

    completed_batches = 0

    # --------------------------------------------------------
    # PROCESS BATCHES
    # --------------------------------------------------------

    for batch_index, page_numbers in enumerate(
        batches,
        start=1
    ):

        batch_start = time.time()

        first_page = (
            page_numbers[0] + 1
        )

        last_page = (
            page_numbers[-1] + 1
        )

        status_text.info(
            f"⚙️ Processing Batch "
            f"{batch_index}/{total_batches} "
            f"| Pages {first_page}–{last_page}"
        )

        # ----------------------------------------------------
        # NEET SEARCHABLE
        # ----------------------------------------------------

        if (
            pdf_type == "searchable"
            and neet_structure
        ):

            records, errors = (
                process_neet_searchable_batch(
                    pdf_path,
                    page_numbers
                )
            )

            if records:

                df = pd.DataFrame(
                    records,
                    columns=NEET_COLUMNS
                )

                all_data.append(
                    df
                )

        # ----------------------------------------------------
        # GENERIC SEARCHABLE
        # ----------------------------------------------------

        elif pdf_type == "searchable":

            batch_data = (
                process_generic_searchable_batch(
                    pdf_path,
                    page_numbers
                )
            )

            all_data.extend(
                batch_data
            )

        # ----------------------------------------------------
        # OCR
        # ----------------------------------------------------

        else:

            batch_data = (
                process_ocr_batch(
                    pdf_path,
                    page_numbers,
                    dpi
                )
            )

            all_data.extend(
                batch_data
            )

        # ----------------------------------------------------
        # BATCH COMPLETE
        # ----------------------------------------------------

        completed_batches += 1

        batch_time = (
            time.time()
            - batch_start
        )

        elapsed = (
            time.time()
            - start_time
        )

        avg_batch_time = (
            elapsed
            / completed_batches
        )

        remaining_batches = (
            total_batches
            - completed_batches
        )

        eta_seconds = (
            remaining_batches
            * avg_batch_time
        )

        progress = (
            completed_batches
            / total_batches
        )

        progress_bar.progress(
            progress
        )

        if eta_seconds < 60:

            eta_text = (
                f"{eta_seconds:.0f} sec"
            )

        else:

            eta_text = (
                f"{eta_seconds / 60:.1f} min"
            )

        status_text.info(
            f"✅ Batch "
            f"{batch_index}/{total_batches} "
            f"completed | "
            f"Pages {first_page}–{last_page} | "
            f"Batch: {batch_time:.1f}s | "
            f"ETA: {eta_text}"
        )

        gc.collect()

    # --------------------------------------------------------
    # COMBINE
    # --------------------------------------------------------

    status_text.info(
        "📊 Combining extracted data..."
    )

    combined_df = combine_dataframes(
        all_data
    )

    del all_data

    gc.collect()

    total_time = (
        time.time()
        - start_time
    )

    progress_bar.progress(
        1.0
    )

    # --------------------------------------------------------
    # NO DATA
    # --------------------------------------------------------

    if combined_df is None:

        status_text.warning(
            f"⚠️ Extraction completed, "
            f"but no data was detected. "
            f"Time: {total_time:.1f}s"
        )

        return None, {

            "total_pages": total_pages,

            "total_batches": total_batches,

            "pdf_type": pdf_type,

            "mode": mode_text,

            "total_time": total_time,

            "rows": 0
        }

    # --------------------------------------------------------
    # COMPLETE
    # --------------------------------------------------------

    rows = len(
        combined_df
    )

    status_text.success(
        f"🎉 Extraction completed! "
        f"{rows:,} rows extracted in "
        f"{total_time:.1f} seconds."
    )

    return combined_df, {

        "total_pages": total_pages,

        "total_batches": total_batches,

        "pdf_type": pdf_type,

        "mode": mode_text,

        "total_time": total_time,

        "rows": rows
    }


# ============================================================
# APP UI
# ============================================================

st.title(
    "📄 Universal PDF Fast Extractor"
)

st.caption(
    "NEET fixed-position extraction + "
    "generic searchable PDF extraction + "
    "OCR • 150-page cloud-safe batching"
)


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    st.header(
        "⚙️ Settings"
    )

    batch_size = st.number_input(
        "Batch Size",
        min_value=1,
        max_value=150,
        value=150,
        step=10,
        help=(
            "Maximum 150 pages per batch. "
            "150 is recommended for cloud processing."
        )
    )

    dpi = st.slider(
        "OCR DPI",
        min_value=MIN_OCR_DPI,
        max_value=MAX_OCR_DPI,
        value=DEFAULT_OCR_DPI,
        step=10,
        help=(
            "Higher DPI can improve OCR accuracy "
            "but increases processing time."
        )
    )

    st.divider()

    st.subheader(
        "☁️ Cloud Mode"
    )

    st.write(
        "Batch size: **150 pages maximum**"
    )

    st.write(
        "OCR: **1 page at a time**"
    )

    st.write(
        "Multiprocessing: **OFF**"
    )

    st.write(
        "RAM optimization: **ON**"
    )

    st.divider()

    st.subheader(
        "🔧 System"
    )

    system_status = check_system()

    if system_status["tesseract"]:

        st.success(
            "Tesseract: OK"
        )

    else:

        st.error(
            "Tesseract: Not found"
        )

    if system_status["pdfplumber"]:

        st.success(
            "pdfplumber: OK"
        )

    else:

        st.error(
            "pdfplumber: Error"
        )

    if system_status["fitz"]:

        st.success(
            "PyMuPDF: OK"
        )

    else:

        st.error(
            "PyMuPDF: Error"
        )

    if system_status["pdf2image"]:

        st.success(
            "pdf2image: OK"
        )

    else:

        st.error(
            "pdf2image: Error"
        )


# ============================================================
# FILE UPLOAD
# ============================================================

uploaded_file = st.file_uploader(
    "📤 Upload PDF",
    type=["pdf"],
    help=(
        "Upload NEET allotment PDF, "
        "table PDF, text PDF or scanned PDF."
    )
)


# ============================================================
# PROCESS
# ============================================================

if uploaded_file is not None:

    st.success(
        f"📄 {uploaded_file.name}"
    )

    file_size_mb = (
        uploaded_file.size
        / (1024 * 1024)
    )

    st.write(
        f"File size: **{file_size_mb:.2f} MB**"
    )

    start_button = st.button(
        "🚀 Start Extraction",
        type="primary",
        use_container_width=True
    )

    if start_button:

        temp_dir = tempfile.mkdtemp(
            prefix="pdf_extractor_"
        )

        safe_filename = (
            os.path.basename(
                uploaded_file.name
            )
        )

        pdf_path = os.path.join(
            temp_dir,
            safe_filename
        )

        try:

            # ------------------------------------------------
            # SAVE PDF
            # ------------------------------------------------

            with open(
                pdf_path,
                "wb"
            ) as file:

                file.write(
                    uploaded_file.getbuffer()
                )

            # ------------------------------------------------
            # PROGRESS
            # ------------------------------------------------

            progress_bar = st.progress(
                0
            )

            status_text = st.empty()

            # ------------------------------------------------
            # EXTRACTION
            # ------------------------------------------------

            with st.spinner(
                "Processing PDF..."
            ):

                combined_df, stats = (
                    run_extraction(
                        pdf_path=pdf_path,
                        batch_size=int(
                            batch_size
                        ),
                        dpi=int(
                            dpi
                        ),
                        progress_bar=progress_bar,
                        status_text=status_text
                    )
                )

            # ------------------------------------------------
            # RESULTS
            # ------------------------------------------------

            if combined_df is not None:

                st.divider()

                st.subheader(
                    "📊 Extraction Summary"
                )

                col1, col2, col3, col4 = (
                    st.columns(4)
                )

                with col1:

                    st.metric(
                        "Pages",
                        f"{stats['total_pages']:,}"
                    )

                with col2:

                    st.metric(
                        "Batches",
                        f"{stats['total_batches']:,}"
                    )

                with col3:

                    st.metric(
                        "Rows",
                        f"{stats['rows']:,}"
                    )

                with col4:

                    st.metric(
                        "Time",
                        f"{stats['total_time']:.1f}s"
                    )

                st.info(
                    f"Extraction mode: "
                    f"**{stats['mode']}**"
                )

                # ------------------------------------------------
                # PREVIEW
                # ------------------------------------------------

                st.subheader(
                    "👀 Preview"
                )

                preview_rows = min(
                    100,
                    len(combined_df)
                )

                st.dataframe(
                    combined_df.head(
                        preview_rows
                    ),
                    use_container_width=True,
                    height=500
                )

                # ------------------------------------------------
                # EXCEL
                # ------------------------------------------------

                with st.spinner(
                    "Creating formatted Excel..."
                ):

                    excel_data = (
                        dataframe_to_excel(
                            combined_df
                        )
                    )

                output_name = (
                    os.path.splitext(
                        safe_filename
                    )[0]
                    + "_extracted.xlsx"
                )

                st.download_button(
                    label="⬇️ Download Excel",
                    data=excel_data,
                    file_name=output_name,
                    mime=(
                        "application/vnd.openxmlformats-"
                        "officedocument.spreadsheetml.sheet"
                    ),
                    type="primary",
                    use_container_width=True
                )

                st.success(
                    "✅ Excel file is ready."
                )

                # ------------------------------------------------
                # MEMORY RELEASE
                # ------------------------------------------------

                del excel_data
                del combined_df

                gc.collect()

            else:

                st.warning(
                    "⚠️ No extractable data was found."
                )

        except Exception as error:

            st.error(
                "❌ Extraction failed."
            )

            st.exception(
                error
            )

        finally:

            # ------------------------------------------------
            # CLEAN TEMP DIRECTORY
            # ------------------------------------------------

            try:

                if os.path.exists(
                    temp_dir
                ):

                    shutil.rmtree(
                        temp_dir,
                        ignore_errors=True
                    )

            except Exception:
                pass

            gc.collect()


# ============================================================
# FOOTER
# ============================================================

st.divider()

st.caption(
    "Universal PDF Fast Extractor • "
    "NEET fixed-position parser • "
    "Generic PDF parser • OCR • "
    "Cloud-safe 150-page architecture"
)
