#!/usr/bin/env python3
"""
Extract only formula cells from Excel files using ExcelLLM
Useful for auditing spreadsheet calculations
"""

import sys
import json
from excelllm import ExcelParser, ExcelFormatter
from collections import defaultdict

def extract_formulas(filename, sheets=None, ranges=None):
    """Extract all formula cells from the specified file/sheets/ranges."""

    # Parse the Excel file
    parser = ExcelParser()
    data = parser.parse_file(filename, sheets=sheets, ranges=ranges)

    # Collect all formulas
    formulas_by_sheet = defaultdict(list)
    total_formulas = 0

    for sheet in data['sheets']:
        sheet_name = sheet['name']

        for cell in sheet['cells']:
            if cell.get('formula'):
                formulas_by_sheet[sheet_name].append({
                    'address': cell['address'],
                    'formula': cell['formula'],
                    'value': cell['value'],
                    'row': cell['row'],
                    'col': cell['col']
                })
                total_formulas += 1

    return formulas_by_sheet, total_formulas, data

def analyze_formulas(formulas_by_sheet):
    """Analyze formula patterns and dependencies."""

    analysis = {
        'formula_types': defaultdict(int),
        'cell_references': set(),
        'functions_used': defaultdict(int)
    }

    import re

    for sheet, formulas in formulas_by_sheet.items():
        for f in formulas:
            formula = f['formula']

            # Count formula types
            if formula.startswith('=SUM'):
                analysis['formula_types']['SUM'] += 1
            elif formula.startswith('=AVERAGE'):
                analysis['formula_types']['AVERAGE'] += 1
            elif formula.startswith('=IF'):
                analysis['formula_types']['IF'] += 1
            elif formula.startswith('=VLOOKUP'):
                analysis['formula_types']['VLOOKUP'] += 1

            # Extract cell references
            cell_refs = re.findall(r'[A-Z]+\d+', formula)
            analysis['cell_references'].update(cell_refs)

            # Extract function names
            functions = re.findall(r'=?([A-Z]+)\(', formula)
            for func in functions:
                analysis['functions_used'][func] += 1

    return analysis

def main():
    if len(sys.argv) < 2:
        print("Usage: python extract_formulas.py <excel_file> [--sheets Sheet1 Sheet2] [--ranges A1:D10]")
        sys.exit(1)

    filename = sys.argv[1]
    sheets = None
    ranges = None

    # Parse command line arguments
    i = 2
    while i < len(sys.argv):
        if sys.argv[i] == '--sheets' and i + 1 < len(sys.argv):
            sheets = []
            i += 1
            while i < len(sys.argv) and not sys.argv[i].startswith('--'):
                sheets.append(sys.argv[i])
                i += 1
            i -= 1
        elif sys.argv[i] == '--ranges' and i + 1 < len(sys.argv):
            ranges = sys.argv[i + 1]
            i += 1
        i += 1

    # Extract formulas
    print(f"Extracting formulas from: {filename}")
    if sheets:
        print(f"Sheets: {', '.join(sheets)}")
    if ranges:
        print(f"Ranges: {ranges}")
    print("-" * 50)

    formulas_by_sheet, total_formulas, data = extract_formulas(filename, sheets, ranges)

    # Display results
    print(f"\nFound {total_formulas} formulas in {len(formulas_by_sheet)} sheets\n")

    for sheet_name, formulas in formulas_by_sheet.items():
        print(f"\n📊 Sheet: {sheet_name} ({len(formulas)} formulas)")
        print("-" * 40)

        # Group by formula type for better readability
        formula_groups = defaultdict(list)
        for f in formulas:
            # Simple grouping by first function
            if '(' in f['formula']:
                func = f['formula'].split('(')[0].replace('=', '')
                formula_groups[func].append(f)
            else:
                formula_groups['Other'].append(f)

        for func_type, group in sorted(formula_groups.items()):
            print(f"\n  {func_type} formulas:")
            for f in group[:5]:  # Show max 5 per type
                print(f"    {f['address']}: {f['formula']}")
                if f['value'] != f['formula']:
                    print(f"             → {f['value']}")

            if len(group) > 5:
                print(f"    ... and {len(group) - 5} more {func_type} formulas")

    # Analyze formulas
    print("\n\n📈 Formula Analysis")
    print("=" * 50)

    analysis = analyze_formulas(formulas_by_sheet)

    print("\nFunctions used:")
    for func, count in sorted(analysis['functions_used'].items(), key=lambda x: x[1], reverse=True):
        print(f"  {func}: {count}")

    print(f"\nUnique cells referenced: {len(analysis['cell_references'])}")

    # Save to JSON for further processing
    output_file = filename.replace('.xlsx', '_formulas.json')
    output_data = {
        'source_file': filename,
        'filters': {
            'sheets': sheets,
            'ranges': ranges
        },
        'total_formulas': total_formulas,
        'formulas': dict(formulas_by_sheet),
        'analysis': {
            'functions_used': dict(analysis['functions_used']),
            'unique_references': len(analysis['cell_references'])
        }
    }

    with open(output_file, 'w') as f:
        json.dump(output_data, f, indent=2)

    print(f"\n💾 Formula data saved to: {output_file}")

if __name__ == "__main__":
    main()
