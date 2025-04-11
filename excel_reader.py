# excel_reader.py
import openpyxl
from config import logging

def read_excel_data(filepath):
    """Read headers and data from the Excel file."""
    try:
        wb = openpyxl.load_workbook(filepath)
        sheet = wb.active
        headers = [str(cell.value).strip() for cell in sheet[1] if cell.value]
        data = []
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if any(row):  # Skip empty rows
                cleaned_row = [val if val is not None else "" for val in row]
                data.append(cleaned_row)
        logging.info(f"Read {len(headers)} headers and {len(data)} rows from Excel.")
        return headers, data
    except Exception as e:
        logging.error(f"Error reading Excel file: {e}")
        raise