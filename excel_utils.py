import logging
from datetime import datetime
from pathlib import Path
import openpyxl

def read_excel_data(filepath):
    """Read headers and data from Excel file, ignoring truly empty rows."""
    filepath = Path(filepath)
    if not filepath.exists():
        raise FileNotFoundError(f"Excel file not found: {filepath}")

    wb = openpyxl.load_workbook(filepath)
    sheet = wb.active
    headers = [
        cell.value.strftime("%Y-%m-%d") if isinstance(cell.value, datetime)
        else str(cell.value).strip()
        for cell in sheet[1] if cell.value
    ]
    # Only include rows with at least one non-empty, non-whitespace value
    data = [
        ["" if val is None else val for val in row]
        for row in sheet.iter_rows(min_row=2, values_only=True)
        if any(val and str(val).strip() for val in row)
    ]
    logging.info(f"Read {len(data)} rows from {filepath.name}: {data}")
    return headers, data