# NEET Universal PDF → Excel Extractor

A Streamlit application for extracting NEET counselling/allotment data from different PDF layouts and exporting it to Excel.

## Features

This application combines multiple extraction methods:

### 1. Table PDF Extraction

Uses `pdfplumber` to detect and extract actual PDF tables.

Suitable for PDFs where the candidate information is arranged in a proper table.

### 2. Fixed-Column PDF Extraction

Uses PyMuPDF (`fitz`) to preserve searchable PDF text alignment.

This is particularly important for PDFs where the data is not technically a table but is aligned underneath fixed column headings.

The extractor preserves the important Quota alignment.

### 3. Scanned PDF OCR

For scanned/image PDFs:

- Converts PDF pages to images
- Uses Tesseract OCR
- Processes one page at a time
- Avoids multiprocessing
- Designed to be safer for Streamlit Cloud

### 4. College-Wise PDF Support

Some counselling PDFs have the college name above the column headers.

Example:

    ABC MEDICAL COLLEGE

    Sr. No.   AIR   NEET Roll No.   CET Form No.   Name   G   Cat.   Quota   Code

    1         ...                         ...        ...    ... OPEN   1103
    2         ...                         ...        ...    ... OBC    1103

The application detects the college heading and associates it with the extracted records.

The college context can also continue across subsequent pages until another college heading is detected.

### 5. Excel Export

The generated Excel file contains:

- Extracted Data
- Extraction Summary

The data sheet includes:

    Sr. No.
    AIR
    NEET Roll No.
    CET Form No.
    Name
    G
    Cat.
    Quota
    Code
    College
    PDF Page

## Important Quota Alignment

For the fixed-column PDF format, the extraction boundaries are:

    Name     = 39
    G        = 73
    Cat.     = 76
    Quota    = 88
    Code     = 115
    College  = 120

The raw PDF text is sliced before whitespace normalization.

This is important because Quota values may contain spaces and brackets.

Examples:

    OPEN
    OPEN (W)
    OPEN (EMD)
    OPEN (W) (EMD)
    DEF2
    HOPEN
    HOPENW

The application attempts to preserve these values rather than splitting them incorrectly.

## Processing Logic

The application first determines whether the PDF is searchable or scanned.

### Searchable PDF

The application uses:

    PyMuPDF
    pdfplumber
    Fixed-position extraction

The fixed-position parser is preferred when valid candidate rows are detected because it preserves the alignment required for fields such as Quota, Code and College.

### Scanned PDF

The application uses:

    pdf2image
    Poppler
    Tesseract OCR

Pages are processed individually to reduce memory usage.

## Cloud-Safe Design

The application is designed for Streamlit Community Cloud.

It intentionally does not use:

- multiprocessing
- CPU pools
- persistent local services
- Windows-specific executable paths

OCR is performed one page at a time.

PDF processing is divided into batches.

Maximum batch size:

    150 pages

Recommended batch size:

    150

For very large scanned PDFs, a smaller batch size can be used.

## Project Structure

Your GitHub repository should contain:

    app.py
    requirements.txt
    packages.txt
    README.md

Example:

    pdf-fast-extractor/
    │
    ├── app.py
    ├── requirements.txt
    ├── packages.txt
    └── README.md

## Streamlit Cloud Deployment

Push all four files to your GitHub repository.

Then create a new Streamlit Community Cloud application.

Set:

    Main file path: app.py

Streamlit Cloud will install Python dependencies from:

    requirements.txt

Linux system packages will be installed from:

    packages.txt

## Local Installation

Install the Python packages:

    python -m pip install -r requirements.txt

For scanned PDFs on Windows, Poppler and Tesseract must also be installed and available to the system.

## Recommended Settings

### Searchable PDF

    Batch Size: 150

OCR DPI is not important for a searchable PDF.

### Scanned PDF

Start with:

    Batch Size: 50–150
    OCR DPI: 180

If the scanned text is small or unclear, try:

    OCR DPI: 220

or:

    OCR DPI: 250

Higher DPI generally improves OCR quality but increases processing time and memory usage.

## Duplicate Removal

The application removes duplicate candidate records based on the extracted data fields rather than the PDF page number.

This helps prevent duplicate rows when the same candidate is detected through multiple extraction attempts.

## College Context

The application supports PDFs where the college name is:

- above the table
- above the column header
- present on the first page of a college section
- omitted from subsequent pages of the same college section

The current college context is carried forward while processing subsequent pages.

## Excel Formatting

The generated Excel workbook includes:

- Bold column headers
- Freeze panes
- Auto-filter
- Useful column widths
- Extraction Summary sheet

For very large datasets, the application automatically creates multiple data sheets when required by Excel's worksheet row limit.

## Troubleshooting

### No data extracted

The PDF may use a layout that is substantially different from the expected NEET candidate-row format.

The current fixed-row parser expects the beginning of a candidate row to resemble:

    Sr. No. AIR NEET Roll No. CET Form No.

For example:

    1 12543 1234567890 12345678 ...

### Quota is incorrectly extracted

For the fixed-layout PDF, Quota depends on the original character alignment.

The important boundaries are:

    QUOTA_START = 88
    CODE_START = 115

Do not normalize/collapse spaces before performing fixed-position slicing.

### College name is missing

If the PDF does not contain the college name in the candidate row or in a recognizable heading above the table, automatic college detection may not be possible.

### OCR is slow

OCR is significantly slower than searchable-text extraction.

Try:

    OCR DPI = 180

instead of:

    OCR DPI = 250

You can also reduce the batch size.

## Supported PDF Types

This application is optimized for:

- NEET counselling PDFs
- NEET allotment PDFs
- Merit lists
- Rank lists
- College-wise allotment lists
- Searchable table PDFs
- Searchable fixed-alignment PDFs
- Scanned/image PDFs
- PDFs where college headings appear above tables
- PDFs where college information continues across multiple pages

## Important Limitation

No PDF extraction program can guarantee perfect extraction from literally every possible PDF design.

Accuracy depends on the source PDF.

The strongest results are expected when:

- Text is digitally generated
- Columns are consistently aligned
- Candidate rows follow a consistent structure
- College headings are clearly separated
- Scanned PDFs are high quality

Highly unusual layouts, distorted scans, handwritten PDFs, rotated pages or PDFs with heavily fragmented text may require additional layout-specific rules.

## Main Technologies

Python

Streamlit

PyMuPDF

pdfplumber

Pandas

OpenPyXL

pdf2image

Poppler

Tesseract OCR

Pillow

## Output Example

The final Excel data follows this structure:

| Sr. No. | AIR | NEET Roll No. | CET Form No. | Name | G | Cat. | Quota | Code | College | PDF Page |
|--------:|----:|---------------|--------------|------|---|------|-------|------|---------|----------|
| 1 | 12543 | 1234567890 | 12345678 | Candidate Name | M | GM | OPEN | 1103 | ABC Medical College | 12 |

## Version

NEET Universal PDF → Excel Extractor

Designed for:

    GitHub
    Streamlit Community Cloud
    Large NEET counselling/allotment PDFs