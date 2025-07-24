#!/usr/bin/env python3
"""
Example usage of ExcelLLM
"""

from excelllm import ExcelParser, ExcelFormatter
import json

def main():
    # Example 1: Basic usage
    print("=== Example 1: Basic Excel parsing ===")

    # Parse an Excel file
    parser = ExcelParser(chunk_size=200)

    # You would use your own Excel file here
    # data = parser.parse_file('your_file.xlsx')

    # For this example, let's create sample data
    sample_data = {
        "filename": "example.xlsx",
        "sheets": [
            {
                "name": "Sales",
                "cells": [
                    {"address": "A1", "value": "Product", "formula": None, "type": "s"},
                    {"address": "B1", "value": "Q1 Sales", "formula": None, "type": "s"},
                    {"address": "C1", "value": "Q2 Sales", "formula": None, "type": "s"},
                    {"address": "D1", "value": "Total", "formula": None, "type": "s"},
                    {"address": "A2", "value": "Widget A", "formula": None, "type": "s"},
                    {"address": "B2", "value": 1000, "formula": None, "type": "n"},
                    {"address": "C2", "value": 1500, "formula": None, "type": "n"},
                    {"address": "D2", "value": 2500, "formula": "=B2+C2", "type": "f"},
                ],
                "dimensions": "2x4"
            }
        ]
    }

    # Example 2: Convert to different formats
    print("\n=== Example 2: Format conversions ===")

    formatter = ExcelFormatter()

    # Convert to JSON
    print("\nJSON output:")
    json_output = formatter.to_json(sample_data, pretty=True)
    print(json_output)

    # Convert to Markdown
    print("\n\nMarkdown output:")
    markdown_output = formatter.to_markdown(sample_data)
    print(markdown_output)

    # Convert to simple text
    print("\n\nText output:")
    text_output = formatter.to_simple_text(sample_data)
    print(text_output)

    # Example 3: Processing specific cells
    print("\n\n=== Example 3: Processing specific cells ===")

    # Find all cells with formulas
    formula_cells = []
    for sheet in sample_data['sheets']:
        for cell in sheet['cells']:
            if cell.get('formula'):
                formula_cells.append({
                    'sheet': sheet['name'],
                    'address': cell['address'],
                    'formula': cell['formula'],
                    'value': cell['value']
                })

    print(f"\nFound {len(formula_cells)} cells with formulas:")
    for cell in formula_cells:
        print(f"  {cell['sheet']}!{cell['address']}: {cell['formula']} = {cell['value']}")

    # Example 4: Chunked processing for large files
    print("\n\n=== Example 4: Handling large files ===")

    # When parsing large files, the parser will chunk automatically
    large_parser = ExcelParser(chunk_size=100)

    # Check if data was truncated
    for sheet in sample_data['sheets']:
        if sheet.get('truncated'):
            print(f"Sheet '{sheet['name']}' was truncated at {len(sheet['cells'])} cells")
            print(f"Total rows in sheet: {sheet.get('total_rows', 'unknown')}")

    # Example 5: Export for LLM consumption
    print("\n\n=== Example 5: Prepare for LLM ===")

    # Create a simplified version for LLM
    llm_data = {
        "summary": f"Excel file with {len(sample_data['sheets'])} sheets",
        "sheets": []
    }

    for sheet in sample_data['sheets']:
        sheet_summary = {
            "name": sheet['name'],
            "cell_count": len(sheet['cells']),
            "has_formulas": any(cell.get('formula') for cell in sheet['cells']),
            "sample_data": sheet['cells'][:5]  # First 5 cells as sample
        }
        llm_data['sheets'].append(sheet_summary)

    print("\nLLM-friendly summary:")
    print(json.dumps(llm_data, indent=2))

    # Example 6: Error handling
    print("\n\n=== Example 6: Error handling ===")

    try:
        # This would fail with a non-existent file
        # data = parser.parse_file('non_existent.xlsx')
        print("Always wrap file parsing in try-except blocks")
    except Exception as e:
        print(f"Error: {e}")

    print("\n\nDone! Use these examples as templates for your own Excel processing needs.")


if __name__ == "__main__":
    main()
