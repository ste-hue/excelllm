#!/usr/bin/env python3
"""
Example of selective Excel parsing with ExcelLLM
Shows how to extract specific sheets and cell ranges
"""

from excelllm import ExcelParser, ExcelFormatter
import json

def main():
    print("=== ExcelLLM Selective Parsing Examples ===\n")

    # Create a parser
    parser = ExcelParser(chunk_size=200)
    formatter = ExcelFormatter()

    # Example 1: Parse only specific sheets
    print("Example 1: Parse only specific sheets")
    print("-" * 40)

    # If you have a file with sheets: Sales, Inventory, Customers, Finance
    # You can parse only Sales and Finance:
    # data = parser.parse_file('company_data.xlsx', sheets=['Sales', 'Finance'])

    # For demonstration, let's show the concept:
    print("Command: excelllm company_data.xlsx -s Sales Finance")
    print("This would extract only the 'Sales' and 'Finance' sheets\n")

    # Example 2: Parse specific cell ranges
    print("\nExample 2: Parse specific cell ranges")
    print("-" * 40)

    # Extract only a specific range from all sheets
    # data = parser.parse_file('data.xlsx', ranges='A1:D10')

    print("Command: excelllm data.xlsx -r A1:D10")
    print("This extracts only cells in the A1:D10 range from all sheets\n")

    # Example 3: Multiple ranges
    print("\nExample 3: Parse multiple cell ranges")
    print("-" * 40)

    # Extract multiple ranges
    # data = parser.parse_file('data.xlsx', ranges='A1:C5,E1:G5,A10:C15')

    print("Command: excelllm data.xlsx -r A1:C5,E1:G5,A10:C15")
    print("This extracts three different ranges from all sheets\n")

    # Example 4: Combine sheet and range filters
    print("\nExample 4: Combine sheet and range filters")
    print("-" * 40)

    # Extract specific ranges from specific sheets only
    # data = parser.parse_file('data.xlsx', sheets=['Summary', 'Details'], ranges='A1:F20')

    print("Command: excelllm data.xlsx -s Summary Details -r A1:F20")
    print("This extracts the A1:F20 range from only 'Summary' and 'Details' sheets\n")

    # Example 5: Common use cases
    print("\nExample 5: Common selective parsing patterns")
    print("-" * 40)

    print("\n1. Extract headers and first 10 rows:")
    print("   excelllm file.xlsx -r A1:Z11")

    print("\n2. Extract summary cells from corner of sheet:")
    print("   excelllm file.xlsx -r A1:E5")

    print("\n3. Extract specific columns (e.g., A and C):")
    print("   excelllm file.xlsx -r A:A,C:C")

    print("\n4. Extract a dashboard sheet with specific KPI cells:")
    print("   excelllm file.xlsx -s Dashboard -r B2:B10,D2:D10,F2:F10")

    # Example 6: Programmatic usage
    print("\n\nExample 6: Programmatic usage with filters")
    print("-" * 40)

    # Simulated data to show the structure
    sample_data = {
        "filename": "financial_report.xlsx",
        "filters": {
            "sheets": ["Q1_Summary", "Q1_Details"],
            "ranges": "A1:D10,F1:H10"
        },
        "sheets": [
            {
                "name": "Q1_Summary",
                "cells": [
                    {"address": "A1", "value": "Metric", "formula": None, "type": "s"},
                    {"address": "B1", "value": "Value", "formula": None, "type": "s"},
                    {"address": "A2", "value": "Revenue", "formula": None, "type": "s"},
                    {"address": "B2", "value": 1500000, "formula": None, "type": "n"},
                    {"address": "A3", "value": "Costs", "formula": None, "type": "s"},
                    {"address": "B3", "value": 1000000, "formula": None, "type": "n"},
                    {"address": "A4", "value": "Profit", "formula": None, "type": "s"},
                    {"address": "B4", "value": 500000, "formula": "=B2-B3", "type": "f"},
                ],
                "dimensions": "10x8",
                "filtered": True
            }
        ]
    }

    print("\nPython code:")
    print("""
    from excelllm import ExcelParser, ExcelFormatter

    # Parse with filters
    parser = ExcelParser()
    data = parser.parse_file(
        'financial_report.xlsx',
        sheets=['Q1_Summary', 'Q1_Details'],
        ranges='A1:D10,F1:H10'
    )

    # Convert to markdown for report
    formatter = ExcelFormatter()
    markdown_report = formatter.to_markdown(data)
    print(markdown_report)
    """)

    # Example 7: Extract formulas only
    print("\n\nExample 7: Extract only cells with formulas")
    print("-" * 40)

    print("\nPython code to filter formula cells:")
    print("""
    # Parse the file
    data = parser.parse_file('spreadsheet.xlsx', ranges='A1:Z100')

    # Extract only cells with formulas
    formula_cells = []
    for sheet in data['sheets']:
        for cell in sheet['cells']:
            if cell.get('formula'):
                formula_cells.append({
                    'sheet': sheet['name'],
                    'cell': cell['address'],
                    'formula': cell['formula'],
                    'value': cell['value']
                })

    # Display results
    for fc in formula_cells:
        print(f"{fc['sheet']}!{fc['cell']}: {fc['formula']} = {fc['value']}")
    """)

    # Example 8: Extract for specific analysis
    print("\n\nExample 8: Extract data for specific analysis")
    print("-" * 40)

    print("\nScenario: You want to analyze only the summary rows (row 1 and rows 50-55)")
    print("from sheets 'Sales_2023' and 'Sales_2024':")
    print("\nCommand: excelllm sales.xlsx -s Sales_2023 Sales_2024 -r A1:Z1,A50:Z55")

    print("\n\nScenario: Extract only the totals column (column G) from all sheets:")
    print("Command: excelllm report.xlsx -r G:G")

    print("\n\nScenario: Extract a specific dashboard grid (B2:F6) from multiple region sheets:")
    print("Command: excelllm regions.xlsx -s North South East West -r B2:F6")

    # Show sample output structure
    print("\n\nSample filtered output structure:")
    print("-" * 40)
    print(json.dumps(sample_data, indent=2))

    print("\n\nTips for selective parsing:")
    print("-" * 30)
    print("1. Use sheet filters to focus on relevant data")
    print("2. Use range filters to extract specific data sections")
    print("3. Combine both for precise data extraction")
    print("4. Multiple ranges can be specified with commas")
    print("5. Column ranges (A:A) extract entire columns")
    print("6. The filters are shown in the output for traceability")


if __name__ == "__main__":
    main()
