# NEET Universal PDF → Excel Extractor

Streamlit application for extracting NEET counselling/allotment data from searchable tables, fixed-column searchable PDFs, and scanned PDFs.

## Output columns

1. Sr. No.
2. AIR
3. NEET Roll No.
4. CET Form No.
5. Name
6. G
7. Cat.
8. Quota
9. Code
10. College
11. PDF Page

## Extraction logic

### 1. Searchable PDFs
- PyMuPDF (`fitz`) reads text blocks while preserving fixed-column spacing.
- `pdfplumber` is also used for real table structures.
- Fixed-position parsing uses the supplied layout boundaries:
  - Name: 39
  - G: 73
  - Cat.: 76
  - Quota: 88
  - Code: 115
  - College: 120
- The raw line is sliced before whitespace cleanup so Quota values such as `OPEN`, `OPEN (W)`, `OPEN (EMD)`, `OPEN (W) (EMD)`, `DEF2`, `HOPEN`, and `HOPENW` are preserved when the source PDF retains alignment.

### 2. College-wise PDFs
A college/institute heading above the table/header is detected and carried forward across following pages until another college heading is found.

### 3. Scanned PDFs
- One page is converted at a time using Poppler/pdf2image.
- Tesseract OCR is used with English language data.
- Fixed-column parsing is attempted first, followed by a conservative OCR fallback.

## Streamlit Cloud deployment

Repository root should contain:

```text
app.py
requirements.txt
packages.txt
README.md
```

Set the Streamlit main file to:

```text
app.py
```

`packages.txt` installs the Linux system packages required by OCR/PDF rendering.

## Recommended settings

Searchable PDF:

- Batch Size: 150

Scanned PDF:

- Batch Size: 50–150
- OCR DPI: 180 initially
- Use 220–250 DPI for small/unclear text if necessary

## Important limitation

No universal PDF parser can guarantee perfect extraction from every possible PDF design. This application is optimized for NEET counselling/allotment layouts similar to the supplied scripts. PDFs with substantially different column structures may require an additional layout-specific rule.

## Local run

```bash
python -m pip install -r requirements.txt
streamlit run app.py
```

## Cloud-safe design

The application does not use multiprocessing. OCR converts only one page at a time, and processing is organized in batches of up to 150 pages to reduce memory pressure.
