# IFL Grade Extractor

A Python tool to extract student grades from IFL Excel spreadsheets and organize them by term and class code.

## Features

- 🔍 Automatically detects grade columns (A-F) and student ID columns (5-digit numbers)
- 📊 Groups grades by term ID and class code from filenames
- 📁 Organizes output into separate CSV files per term/class combination
- 🗂️ Moves unidentified files to `not-found/` directory
- ✅ Validates term IDs against a reference CSV file
- 📝 Comprehensive reporting and logging

## Requirements

- Python 3.14+
- pandas
- openpyxl

## Installation

```bash
# Clone the repository
git clone <repository-url>
cd extract_grades

# Install dependencies
pip install -e .
```

## Usage

### Basic Usage

Process all files in the default directory:

```bash
python extract_grades.py
```

This will:
- Read files from `~/Documents/IFL_CONSOLIDATED/`
- Match filenames against term IDs in `terms.csv`
- Extract class codes (e.g., EHSS-101, GESL-205)
- Output to `./extracted/grades_extract_{termid}_{classcode}.csv`

### Custom Paths

```bash
# Specify custom input directory
python extract_grades.py /path/to/xlsx/files

# Specify custom output directory
python extract_grades.py -o /path/to/output

# Use custom terms file
python extract_grades.py -t /path/to/terms.csv

# Verbose output
python extract_grades.py -v
```

### Process Single File

```bash
python extract_grades.py /path/to/single/file.xlsx
```

## Output Structure

### Successfully Processed Files

```
extracted/
├── grades_extract_2023T2E_EHSS-101.csv
├── grades_extract_2023T2E_GESL-205.csv
├── grades_extract_2022T4E_IEAP-1234.csv
└── ...
```

Each CSV contains:
- `filename` - Original xlsx filename (without extension)
- `student_id` - 5-digit zero-padded student ID
- `grade` - Letter grade (A, B, C, D, or F)

### Unidentified Files

Files that cannot be processed are moved to `not-found/`:

```
not-found/
├── file_without_termid.xlsx       (missing term ID)
├── file_without_classcode.xlsx    (missing class code)
└── file_without_grades.xlsx       (no grades found)
```

## Term ID Patterns

The tool recognizes various term ID formats:
- Modern: `251216E-T1AE`, `250916E-T4AE`
- Standard: `2023T2E`, `2022T4E`
- Special: `2022AT3E`, `2019-2020T3E`
- Complex: `2022T4ET4E`, `2023T2BT2E`

## Class Code Patterns

Class codes must match the pattern `[A-Z]{4}-[A-Z0-9]+`:
- `EHSS-101`
- `GESL-205A`
- `IEAP-1234`

## Terms File Format

`terms.csv` should contain:

```csv
termid,startdate
251216E-T1AE,2025-12-16 00:00:00.000
250916E-T4AE,2025-09-16 00:00:00.000
2023T2E,2023-02-13 00:00:00.000
...
```

## Grade Detection Logic

The tool automatically identifies grade and ID columns by:

1. **Grade Column**: >50% of values are valid grades (A-F), minimum 3 grades
2. **ID Column**: >50% of values are 4-5 digit numbers, minimum 3 IDs

## Example Output

```
Found 150 xlsx files to process
Loaded 51 term IDs from terms.csv

Wrote 25 records for 2023T2E EHSS-101 to grades_extract_2023T2E_EHSS-101.csv
Wrote 30 records for 2023T2E GESL-205 to grades_extract_2023T2E_GESL-205.csv
...

============================================================
SUMMARY
============================================================
  Files processed: 150
  Files with grades: 145
  Files without grades: 2
  Files without term ID: 1
  Files without class code: 2
  Files moved to not-found/: 5
  Total grade records: 3,542
  CSV files written: 142
  Output directory: ./extracted
  Not-found directory: ./not-found
```

## Development

Project structure:
```
extract_grades/
├── extract_grades.py      # Main script
├── terms.csv              # Term ID reference
├── pyproject.toml         # Dependencies
├── README.md              # This file
├── extracted/             # Output directory (generated)
└── not-found/             # Unidentified files (generated)
```

## License

MIT License

## Contributing

Contributions welcome! Please open an issue or submit a pull request.