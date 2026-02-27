# Excel to PDF Converter

A Python program that converts Excel files (.xls, .xlsx) to PDF format using Win32 COM (Windows only).

## Requirements

- **Windows OS** (required for Win32 COM)
- **Microsoft Excel** installed on the system
- Python with pywin32 library

## Installation

1. Install the required dependency:
```bash
pip install -r requirements.txt
```

## Usage

### Command Line
```bash
python excel_to_pdf.py <excel_file_path> [print_area] [pdf_output_path] [fit_to_one_page]
```

### Examples
```bash
# Basic conversion (fits to one page by default)
python excel_to_pdf.py boletin-diario_2026_02_20.xls

# With custom print area (fits to one page)
python excel_to_pdf.py data.xlsx "A1:Z50"

# With custom print area and PDF name (fits to one page)
python excel_to_pdf.py data.xlsx "A1:Z50" output.pdf

# Disable one-page fitting
python excel_to_pdf.py data.xlsx "A1:Z50" output.pdf false
```

### Programmatic Usage
```python
from excel_to_pdf import excel_to_pdf

# Basic conversion (fits to one page by default)
result = excel_to_pdf("input.xlsx", "output.pdf")

# With print area (fits to one page)
result = excel_to_pdf("input.xlsx", "A1:Z50", "output.pdf")

# Without one-page fitting
result = excel_to_pdf("input.xlsx", "A1:Z50", "output.pdf", False)

if result:
    print("Conversion successful!")
```

## Features

- Supports both .xls and .xlsx files
- Automatic PDF naming (same name as Excel file)
- High-quality PDF output using Excel's native export
- Preserves formatting, charts, and images from Excel
- **Custom print area control** - specify exact ranges (e.g., "A1:Z50")
- **One-page fitting** - automatically fits content to one page (default)
- Error handling and validation
- Command-line interface

## Dependencies

- pywin32: For Windows COM automation

## Notes

- **Windows only**: This solution requires Windows OS
- **Microsoft Excel required**: Excel must be installed on the system
- **High quality**: Uses Excel's native PDF export for best results
- **Preserves formatting**: Maintains all Excel formatting, charts, and images
- **Background operation**: Excel runs invisibly during conversion
- **Print area format**: Use Excel cell notation (e.g., "A1:Z50", "B2:G100")
- **Active sheet**: Always converts the active worksheet
- **One-page fitting**: Enabled by default, scales content to fit on one page
