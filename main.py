import logging
from logging_config import configure_logging
from driver_utils import terminate_chrome_processes, initialize_driver
from excel_utils import read_excel_data
from form_utils import get_form_headers, fill_google_form
from matching_utils import match_headers
from config import EXCEL_FILE
from pathlib import Path
import openpyxl
import sys

def main():
    """Main function to orchestrate the automation process."""
    # Configure logging
    configure_logging()
    driver = None
    try:
        # Load Excel file
        filepath = Path(EXCEL_FILE)
        wb = openpyxl.load_workbook(filepath)
        sheet = wb.active
        
        # Read headers and data from Excel
        excel_headers, rows = read_excel_data(EXCEL_FILE)
        logging.info(f"Total rows to process: {len(rows)}")

        # Check if "Note" column exists; if not, add it
        note_column = None
        for col_idx, cell in enumerate(sheet[1], start=1):
            if cell.value and cell.value.lower() == "note":  # Case-insensitive check
                note_column = col_idx
                break
        if not note_column:
            note_column = len(excel_headers) + 1
            sheet.cell(row=1, column=note_column).value = "Note"
            logging.info("Added 'Note' column to Excel file at column {note_column}")

        # Exit if no data rows to process
        if not rows:
            logging.info("No data rows to process in Excel file")
            wb.save(filepath)
            return

        # Initialize WebDriver and get Google Form headers
        terminate_chrome_processes()
        driver = initialize_driver()
        form_headers = get_form_headers(driver)
        header_mapping, unmatched_headers = match_headers(excel_headers, form_headers)

        # Process each row
        for idx, row in enumerate(rows, start=2):
            # Check if row is already marked as inserted
            note_cell = sheet.cell(row=idx, column=note_column).value
            if note_cell == "Inserted":
                logging.info(f"Row {idx} already inserted, skipping")
                continue

            logging.info(f"Processing row {idx}: {row}")
            # Attempt to fill and submit the Google Form
            success = fill_google_form(driver, row, excel_headers, header_mapping)

            if success:
                # Mark row as inserted in "Note" column
                sheet.cell(row=idx, column=note_column).value = "Inserted"
                logging.info(f"Row {idx} processed successfully")
            else:
                # Log error, update "Note" column, save Excel, and exit
                error_message = f"Failed to insert row {idx}"
                sheet.cell(row=idx, column=note_column).value = error_message
                logging.error(f"{error_message} - Row data: {row}")
                
                # Save Excel file before exiting
                wb.save(filepath)
                logging.info(f"Excel file saved with error note for row {idx}")
                
                # Close driver and exit program
                driver.quit()
                sys.exit(1)

            # Save Excel file after each successful row
            wb.save(filepath)
            logging.info(f"Excel file saved after processing row {idx}")

    except Exception as e:
        # Handle unexpected errors, save Excel, and exit
        logging.error(f"Main process error: {e}")
        wb.save(filepath)
        logging.info("Excel file saved due to critical error")
        if driver:
            driver.quit()
        sys.exit(1)

    finally:
        # Ensure WebDriver is closed and Excel file is saved
        if driver:
            try:
                driver.quit()
                logging.info("WebDriver closed")
            except Exception as e:
                logging.error(f"Error closing WebDriver: {e}")
        wb.save(filepath)
        logging.info("Excel file saved in finally block")

if __name__ == "__main__":
    main()