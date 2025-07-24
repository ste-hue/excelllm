# ExcelLLM

Convert Excel files to LLM-friendly formats (JSON/Markdown) with a single command. Extract formulas, values, and cell metadata from .xlsx files.

## 🚀 Web Interface Quick Start

### One-Click Launch
```bash
python run_excelllm.py
```

This will:
- Install all dependencies automatically
- Start the web server at http://localhost:8080
- Open your browser to the interface

### Features
- 🎯 **Drag & Drop** - Drop Excel files directly onto the page
- 📊 **Visual Sheet Selection** - See all sheets with dimensions
- 📍 **Interactive Range Picker** - Click examples or type custom ranges
- 💾 **Multiple Export Formats** - JSON, Markdown, CSV, Plain Text
- 🚀 **Live Preview** - See data before downloading

## Installation

```bash
# Clone the repository
git clone https://github.com/yourusername/excelllm.git
cd excelllm

# Install with pip
pip install -e .

# Or install just the dependency
pip install openpyxl
```

## Usage

### Web Interface (Recommended)

1. **Start the server:**
   ```bash
   python run_excelllm.py
   ```

2. **Open your browser** at http://localhost:8080

3. **Upload your Excel file** by dragging it onto the page

4. **Select sheets and ranges** interactively

5. **Choose output format** and download

### Command Line

```bash
# Convert Excel to JSON (stdout)
python excelllm.py data.xlsx

# Save as JSON file
python excelllm.py data.xlsx -o output.json

# Convert to Markdown
python excelllm.py data.xlsx -f markdown -o output.md

# Convert to simple text
python excelllm.py data.xlsx -f text

# Process more rows per sheet (default: 200)
python excelllm.py data.xlsx --chunk 500

# Pretty-print JSON
python excelllm.py data.xlsx --pretty

# SELECTIVE FILTERING:

# Extract only specific sheets
python excelllm.py data.xlsx -s Sheet1 Sheet2

# Extract only specific cell ranges
python excelllm.py data.xlsx -r A1:D10

# Extract multiple ranges
python excelllm.py data.xlsx -r A1:C5,E1:G5,A10:C15

# Combine sheet and range filters
python excelllm.py data.xlsx -s Summary Details -r A1:F20

# Extract specific columns
python excelllm.py data.xlsx -r A:A,C:C,E:E

# Extract headers and first 10 rows
python excelllm.py data.xlsx -r A1:Z11
```

### Python API

```python
from excelllm import ExcelParser, ExcelFormatter

# Parse Excel file
parser = ExcelParser(chunk_size=200)
data = parser.parse_file('data.xlsx')

# Parse with selective filtering
data = parser.parse_file(
    'data.xlsx',
    sheets=['Sheet1', 'Summary'],  # Only these sheets
    ranges='A1:D10,F1:H10'         # Only these cell ranges
)

# Convert to different formats
formatter = ExcelFormatter()
json_output = formatter.to_json(data, pretty=True)
markdown_output = formatter.to_markdown(data)
text_output = formatter.to_simple_text(data)
```

## Output Formats

### JSON Format
```json
{
  "filename": "data.xlsx",
  "filters": {
    "sheets": ["Sheet1"],
    "ranges": "A1:D10"
  },
  "sheets": [
    {
      "name": "Sheet1",
      "cells": [
        {
          "address": "A1",
          "value": 100,
          "formula": null,
          "type": "n",
          "row": 1,
          "col": 1
        },
        {
          "address": "B1",
          "value": 200,
          "formula": "=A1*2",
          "type": "f",
          "row": 1,
          "col": 2
        }
      ],
      "dimensions": "10x5",
      "filtered": true,
      "truncated": false
    }
  ]
}
```

### Markdown Format
```markdown
# Excel File: data.xlsx

## Sheet: Sheet1
Dimensions: 10x5

| Cell | Value | Formula | Type |
|------|-------|---------|------|
| A1   | 100   |         | n    |
| B1   | 200   | =A1*2   | f    |
```

## Features

- **Formula Extraction**: Captures both formulas and calculated values
- **Multiple Formats**: JSON, Markdown, and plain text output
- **Selective Filtering**: Extract specific sheets and cell ranges
- **Chunking**: Process large files in manageable chunks
- **Merged Cells**: Detects and reports merged cell ranges
- **Minimal Dependencies**: Works with just `openpyxl`, falls back to XML parsing
- **CLI & API**: Use from command line or import as Python module

### Selective Filtering

ExcelLLM allows you to focus on specific parts of your Excel files:

- **Sheet Filtering**: Process only the sheets you need
- **Range Filtering**: Extract specific cell ranges (e.g., A1:D10)
- **Multiple Ranges**: Specify multiple ranges separated by commas
- **Column/Row Selection**: Extract entire columns (A:A) or rows
- **Combined Filters**: Use sheet and range filters together

This is particularly useful for:
- Extracting summary sections from large reports
- Focusing on specific KPI cells in dashboards
- Analyzing only formula cells
- Extracting headers and sample data
- Comparing specific ranges across sheets

## Limitations

- Only supports .xlsx files (not .xls)
- Default chunk size is 200 rows per sheet (configurable)
- Basic XML parser fallback has limited features

## Web Interface Details

### Manual Setup
If the launcher doesn't work, install dependencies manually:
```bash
pip install fastapi uvicorn[standard] openpyxl python-multipart
python excelllm_webapp.py
```

### API Endpoints
The web app also provides REST API:

```bash
# Upload and analyze
curl -X POST -F "file=@data.xlsx" http://localhost:8000/api/upload

# Process with filters
curl -X POST -F "file=@data.xlsx" \
     -F "sheets=Sheet1,Sheet2" \
     -F "range=A1:D10" \
     -F "format=json" \
     http://localhost:8000/api/process

# Download in any format
curl -X POST -F "file=@data.xlsx" \
     -F "format=csv" \
     http://localhost:8000/api/download -o output.csv
```

### Troubleshooting

**Port already in use?**
Edit `excelllm_webapp.py` and change the port in the last line.

**Can't see the interface?**
Make sure you're using a modern browser (Chrome, Firefox, Safari, Edge).

## License

MIT License# excelllm
