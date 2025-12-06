#!/usr/bin/env python3
"""
Extract grades from IFL grade spreadsheets.

Searches each sheet for:
- A column containing letter grades (A-F)
- A column containing 5-digit student IDs

Processes files from ~/Documents/IFL_CONSOLIDATED/ and outputs
separate CSV files for each term ID to extracted/ folder.
"""

import os
import re
import csv
import argparse
import shutil
from pathlib import Path
from collections import defaultdict
import pandas as pd

# Valid letter grades
VALID_GRADES = {'A', 'B', 'C', 'D', 'F'}


def load_terms(terms_file='terms.csv'):
    """
    Load valid term IDs from terms.csv.
    Returns a set of term IDs.
    """
    terms = set()
    terms_path = Path(terms_file)

    if not terms_path.exists():
        print(f"WARNING: {terms_file} not found, continuing without term validation")
        return terms

    try:
        with open(terms_path, 'r') as f:
            reader = csv.DictReader(f)
            for row in reader:
                if 'termid' in row:
                    terms.add(row['termid'])
        print(f"Loaded {len(terms)} term IDs from {terms_file}")
    except Exception as e:
        print(f"ERROR loading terms: {e}")

    return terms


def extract_class_code_from_filename(filename):
    """
    Extract class code from filename.
    Class codes look like: EHSS-*, GESL-*, IEAP-*, etc.
    Returns class code if found, None otherwise.
    """
    # Pattern: 4 uppercase letters, hyphen, digits (and possibly letters)
    # Examples: EHSS-101, GESL-205A, IEAP-1234
    pattern = r'[A-Z]{4}-[A-Z0-9]+'

    match = re.search(pattern, filename.upper())
    if match:
        return match.group(0)

    return None


def extract_termid_from_filename(filename, valid_terms):
    """
    Extract term ID from filename by matching against valid term IDs.
    Returns term ID if found, None otherwise.

    Tries multiple matching strategies:
    1. Exact substring match
    2. Pattern-based extraction with validation
    """
    filename_upper = filename.upper()

    # Strategy 1: Check if any valid term ID appears in the filename
    for term in valid_terms:
        if term.upper() in filename_upper:
            return term

    # Strategy 2: Extract patterns that look like term IDs
    # Pattern examples: 251216E-T1AE, 2023T2E, 2022T4ET4E, 2019-2020T3E
    patterns = [
        r'\d{6}E?-T\d[A-Z]{2}',  # 251216E-T1AE
        r'\d{4}T\d[A-Z]?E?',      # 2023T2E, 2022T4E
        r'\d{4}[A-Z]?T\d[A-Z]?E', # 2022AT3E
        r'\d{4}-\d{4}T\d[A-Z]?E', # 2019-2020T3E
        r'\d{4}T\dT\dE',          # 2022T4ET4E
        r'\d{4}T\d[A-Z]T\dE',     # 2023T2BT2E
    ]

    for pattern in patterns:
        match = re.search(pattern, filename_upper)
        if match:
            candidate = match.group(0)
            # Validate against known terms
            for term in valid_terms:
                if term.upper() == candidate:
                    return term

    return None


def is_grade_column(series):
    """
    Check if a pandas Series looks like a grade column.
    Returns (is_grade_column, grade_values_count)
    """
    # Drop NaN and convert to string
    values = series.dropna().astype(str).str.strip().str.upper()
    
    if len(values) == 0:
        return False, 0
    
    # Count how many values are valid grades
    grade_count = sum(1 for v in values if v in VALID_GRADES)
    
    # Consider it a grade column if >50% are valid grades and at least 3 grades
    ratio = grade_count / len(values) if len(values) > 0 else 0
    
    return ratio > 0.5 and grade_count >= 3, grade_count


def is_id_column(series):
    """
    Check if a pandas Series looks like a student ID column.
    IDs are typically 5-digit numbers.
    """
    values = series.dropna()
    
    if len(values) == 0:
        return False, 0
    
    id_count = 0
    for v in values:
        # Convert to string and check if it's a 4-5 digit number
        s = str(v).strip()
        # Handle floats like 14354.0
        if '.' in s:
            s = s.split('.')[0]
        if re.match(r'^\d{4,5}$', s):
            id_count += 1
    
    ratio = id_count / len(values) if len(values) > 0 else 0
    return ratio > 0.5 and id_count >= 3, id_count


def find_grade_and_id_columns(df):
    """
    Find the grade column and ID column in a dataframe.
    Returns (grade_col_idx, id_col_idx) or (None, None) if not found.
    """
    grade_col = None
    grade_col_count = 0
    id_col = None
    id_col_count = 0
    
    for col_idx in range(len(df.columns)):
        col_data = df.iloc[:, col_idx]
        
        # Check for grade column
        is_grade, g_count = is_grade_column(col_data)
        if is_grade and g_count > grade_col_count:
            grade_col = col_idx
            grade_col_count = g_count
        
        # Check for ID column
        is_id, i_count = is_id_column(col_data)
        if is_id and i_count > id_col_count:
            id_col = col_idx
            id_col_count = i_count
    
    return grade_col, id_col


def extract_grades_from_sheet(df, sheet_name):
    """
    Extract (student_id, grade) pairs from a dataframe.
    Returns list of tuples or empty list if not a grades sheet.
    """
    grade_col, id_col = find_grade_and_id_columns(df)
    
    if grade_col is None or id_col is None:
        return []
    
    results = []
    
    for idx in range(len(df)):
        grade_val = df.iloc[idx, grade_col]
        id_val = df.iloc[idx, id_col]
        
        # Skip if either is NaN
        if pd.isna(grade_val) or pd.isna(id_val):
            continue
        
        # Clean grade
        grade = str(grade_val).strip().upper()
        if grade not in VALID_GRADES:
            continue
        
        # Clean ID - convert to 5-digit zero-padded string
        id_str = str(id_val).strip()
        if '.' in id_str:
            id_str = id_str.split('.')[0]
        
        if not re.match(r'^\d{4,5}$', id_str):
            continue
        
        # Zero-pad to 5 digits
        student_id = id_str.zfill(5)
        
        results.append((student_id, grade))
    
    return results


def process_xlsx_file(filepath):
    """
    Process a single xlsx file and return list of (student_id, grade) tuples.
    """
    try:
        xl = pd.ExcelFile(filepath)
    except Exception as e:
        print(f"ERROR: Could not open {filepath}: {e}")
        return []
    
    all_results = []
    
    for sheet_name in xl.sheet_names:
        try:
            df = pd.read_excel(filepath, sheet_name=sheet_name, header=None)
            results = extract_grades_from_sheet(df, sheet_name)
            
            if results:
                # Found grades in this sheet
                all_results.extend(results)
                # Usually we only want one sheet's grades, but some files might have multiple
                # For safety, break after finding the first sheet with grades
                break
                
        except Exception as e:
            print(f"WARNING: Error reading sheet '{sheet_name}' in {filepath}: {e}")
            continue
    
    return all_results


def process_directory_by_term(input_dir, output_dir, terms_file='terms.csv', verbose=False):
    """
    Process all xlsx files in input_dir, group by term ID and class code, and write separate CSV files.
    Outputs to output_dir/extracted/grades_extract_{termid}_{classcode}.csv
    Moves unidentified files to output_dir/not-found/
    """
    input_path = Path(input_dir)
    output_path = Path(output_dir) / 'extracted'
    notfound_path = Path(output_dir) / 'not-found'

    # Create output directories if needed
    output_path.mkdir(parents=True, exist_ok=True)
    notfound_path.mkdir(parents=True, exist_ok=True)

    # Load valid term IDs
    valid_terms = load_terms(terms_file)

    # Find all xlsx files
    xlsx_files = list(input_path.rglob('*.xlsx'))

    # Filter out temp files and __MACOSX
    xlsx_files = [f for f in xlsx_files
                  if not f.name.startswith('~$')
                  and not f.name.startswith('.')
                  and '__MACOSX' not in str(f)]

    print(f"Found {len(xlsx_files)} xlsx files to process")

    # Group records by (term_id, class_code) tuple
    records_by_key = defaultdict(list)

    total_records = 0
    files_with_grades = 0
    files_without_grades = []
    files_without_termid = []
    files_without_classcode = []
    files_moved_to_notfound = []

    # Process each file
    for xlsx_file in xlsx_files:
        file_stem = xlsx_file.stem

        # Extract term ID and class code from filename
        termid = extract_termid_from_filename(xlsx_file.name, valid_terms)
        class_code = extract_class_code_from_filename(xlsx_file.name)

        if verbose:
            print(f"Processing: {xlsx_file.name}")
            print(f"  -> Term: {termid}, Class: {class_code}")

        results = process_xlsx_file(xlsx_file)

        if results:
            files_with_grades += 1
            total_records += len(results)

            if termid and class_code:
                # Both term ID and class code found - good!
                key = (termid, class_code)
                for student_id, grade in results:
                    records_by_key[key].append({
                        'filename': file_stem,
                        'student_id': student_id,
                        'grade': grade
                    })

                if verbose:
                    print(f"  -> Found {len(results)} grades")
            else:
                # Missing either term ID or class code - move to not-found
                if not termid:
                    files_without_termid.append(xlsx_file.name)
                    reason = "no_termid"
                    if verbose:
                        print(f"  -> Found {len(results)} grades but no term ID match")
                else:
                    files_without_classcode.append(xlsx_file.name)
                    reason = "no_classcode"
                    if verbose:
                        print(f"  -> Found {len(results)} grades but no class code")

                # Move to not-found directory
                dest = notfound_path / xlsx_file.name
                try:
                    shutil.move(str(xlsx_file), str(dest))
                    files_moved_to_notfound.append((xlsx_file.name, reason))
                    if verbose:
                        print(f"  -> Moved to not-found/")
                except Exception as e:
                    print(f"  -> ERROR moving file: {e}")
        else:
            files_without_grades.append(xlsx_file.name)
            if verbose:
                print(f"  -> No grades found")

            # Move to not-found directory
            dest = notfound_path / xlsx_file.name
            try:
                shutil.move(str(xlsx_file), str(dest))
                files_moved_to_notfound.append((xlsx_file.name, "no_grades"))
                if verbose:
                    print(f"  -> Moved to not-found/")
            except Exception as e:
                print(f"  -> ERROR moving file: {e}")

    # Write separate CSV file for each (term_id, class_code) combination
    csv_files_written = []
    for (termid, class_code), records in sorted(records_by_key.items()):
        output_csv = output_path / f'grades_extract_{termid}_{class_code}.csv'

        with open(output_csv, 'w', newline='') as f:
            writer = csv.DictWriter(f, fieldnames=['filename', 'student_id', 'grade'])
            writer.writeheader()
            writer.writerows(records)

        csv_files_written.append((termid, class_code, len(records), output_csv))
        print(f"Wrote {len(records)} records for {termid} {class_code} to {output_csv.name}")

    # Summary
    print(f"\n{'='*60}")
    print(f"SUMMARY")
    print(f"{'='*60}")
    print(f"  Files processed: {len(xlsx_files)}")
    print(f"  Files with grades: {files_with_grades}")
    print(f"  Files without grades: {len(files_without_grades)}")
    print(f"  Files without term ID: {len(files_without_termid)}")
    print(f"  Files without class code: {len(files_without_classcode)}")
    print(f"  Files moved to not-found/: {len(files_moved_to_notfound)}")
    print(f"  Total grade records: {total_records}")
    print(f"  CSV files written: {len(csv_files_written)}")
    print(f"  Output directory: {output_path}")
    print(f"  Not-found directory: {notfound_path}")

    if files_moved_to_notfound:
        print(f"\nMoved to not-found/ directory:")
        # Group by reason
        by_reason = defaultdict(list)
        for filename, reason in files_moved_to_notfound:
            by_reason[reason].append(filename)

        for reason, filenames in by_reason.items():
            reason_label = {
                'no_termid': 'No term ID match',
                'no_classcode': 'No class code',
                'no_grades': 'No grades found'
            }.get(reason, reason)
            print(f"\n  {reason_label} ({len(filenames)} files):")
            for f in filenames[:5]:
                print(f"    - {f}")
            if len(filenames) > 5:
                print(f"    ... and {len(filenames) - 5} more")

    return records_by_key


def process_single_file(filepath, output_dir):
    """Process a single file and write CSV."""
    filepath = Path(filepath)
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)
    
    file_stem = filepath.stem
    results = process_xlsx_file(filepath)
    
    if not results:
        print(f"No grades found in {filepath}")
        return []
    
    records = []
    for student_id, grade in results:
        records.append({
            'filename': file_stem,
            'student_id': student_id,
            'grade': grade
        })
    
    # Write CSV
    output_csv = output_path / 'grades_extract.csv'
    
    # Append if file exists, otherwise create with header
    file_exists = output_csv.exists()
    
    with open(output_csv, 'a', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=['filename', 'student_id', 'grade'])
        if not file_exists:
            writer.writeheader()
        writer.writerows(records)
    
    print(f"Extracted {len(records)} grades from {filepath.name}")
    print(f"Output: {output_csv}")
    
    return records


def main():
    parser = argparse.ArgumentParser(
        description='Extract grades from IFL xlsx files and organize by term ID'
    )
    parser.add_argument(
        'input',
        nargs='?',
        default=os.path.expanduser('~/Documents/IFL_CONSOLIDATED'),
        help='Input xlsx file or directory (default: ~/Documents/IFL_CONSOLIDATED)'
    )
    parser.add_argument(
        '-o', '--output',
        default='.',
        help='Output directory for extracted/ folder (default: current directory)'
    )
    parser.add_argument(
        '-t', '--terms',
        default='terms.csv',
        help='Path to terms.csv file (default: terms.csv)'
    )
    parser.add_argument(
        '-v', '--verbose',
        action='store_true',
        help='Verbose output'
    )

    args = parser.parse_args()

    input_path = Path(args.input).expanduser()

    if not input_path.exists():
        print(f"Error: {args.input} does not exist")
        return 1

    if input_path.is_file():
        # Single file mode - process and output to extracted/ folder
        print("Single file mode - processing one xlsx file")
        process_single_file(input_path, args.output)
    elif input_path.is_dir():
        # Directory mode - group by term ID
        print(f"Directory mode - processing all xlsx files in {input_path}")
        process_directory_by_term(
            input_path,
            args.output,
            terms_file=args.terms,
            verbose=args.verbose
        )
    else:
        print(f"Error: {args.input} is not a valid file or directory")
        return 1

    return 0


if __name__ == '__main__':
    exit(main())
