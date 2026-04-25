# Excel Compare Python Project

A small Python utility for comparing two `.xlsx` workbooks and generating a diff report workbook.

## Installation

Install the required dependency:

```bash
python -m pip install -r requirements.txt
```

## Usage

Compare two Excel files and generate a diff workbook:

```bash
python compare_excels.py left.xlsx right.xlsx --output-file diff.xlsx
```

The generated `diff.xlsx` workbook includes:

- `SUMMARY`: sheet-by-sheet status and difference counts
- `Diff: <SheetName>`: changed cells and values for each compared sheet

## Notes

- Supports `.xlsx` files only
- Compares each sheet cell-by-cell
- Missing sheets are reported in the output workbook

The `skills` folder is reserved for adding custom skills later.