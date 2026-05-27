# XLSX to CSV Splitter

A simple, OS-agnostic tool to split Excel files with multiple worksheets into separate CSV files. Optimized for LLM ingestion with clean output and proper handling of edge cases.

## Features

- Split any XLSX file into one CSV per worksheet
- Automatic filename sanitization (handles special characters)
- Smart empty row detection (stops at consecutive empty rows)
- UTF-8 encoding for full Unicode support
- Optional row index column
- Select specific sheets or export all
- Works on Windows, macOS, and Linux

## Installation

1. Install Python 3.8 or later
2. Install the dependency:

```bash
pip install openpyxl
```

Or use the requirements file:

```bash
pip install -r requirements.txt
```

## Usage

### Basic Usage

Export all sheets from an Excel file:

```bash
python xlsx_to_csv.py your_file.xlsx
```

This creates a folder `your_file_csv/` with CSV files for each worksheet.

### Specify Output Directory

```bash
python xlsx_to_csv.py your_file.xlsx --output-dir ./output
```

### Export Specific Sheets

```bash
python xlsx_to_csv.py your_file.xlsx --sheet-names "Sheet1,Sheet2,Data"
```

### Include Row Index

Add a row number column (useful for LLM context):

```bash
python xlsx_to_csv.py your_file.xlsx --include-index
```

### Full Options

```bash
python xlsx_to_csv.py your_file.xlsx \
    --output-dir ./csv_output \
    --sheet-names "Sheet1,Sheet2" \
    --include-index \
    --max-empty-rows 5
```

## Options

| Option | Description | Default |
|--------|-------------|---------|
| `input` | Input XLSX file path (required) | - |
| `-o, --output-dir` | Output directory for CSV files | `<input_name>_csv` |
| `-s, --sheet-names` | Comma-separated sheet names to export | All sheets |
| `-i, --include-index` | Add row index as first column | False |
| `--max-empty-rows` | Stop after N consecutive empty rows | 3 |

## Examples

### LLM Ingestion Workflow

```bash
# Split Excel file
python xlsx_to_csv.py company_data.xlsx --output-dir ./llm_input

# Now you can feed individual CSVs to your LLM
# Each CSV has a clean name matching the original worksheet
```

### Batch Processing

```bash
# Process multiple Excel files (bash/zsh)
for file in *.xlsx; do
    python xlsx_to_csv.py "$file" --output-dir "./csv_output/$(basename "$file" .xlsx)"
done
```

### Windows PowerShell

```powershell
# Process all Excel files in current directory
Get-ChildItem *.xlsx | ForEach-Object {
    python xlsx_to_csv.py $_.Name --output-dir ".\csv_output\$($_.BaseName)"
}
```

## Output Format

- **Encoding**: UTF-8
- **Delimiter**: Comma
- **Line ending**: OS default
- **Empty values**: Empty strings (no quoted nulls)
- **Filenames**: Sanitized worksheet names (special chars replaced with `_`)

## Why This Tool for LLM Ingestion?

1. **Clean Data**: Removes trailing empty rows that waste tokens
2. **Predictable Structure**: Each CSV has consistent formatting
3. **UTF-8**: Handles international characters without encoding issues
4. **Named Files**: Easy to reference specific data sources in prompts
5. **No Formulas**: Exports values only (formulas evaluated)

## License

MIT License - Feel free to use and modify.
