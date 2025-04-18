import logging
from logging_config import configure_logging
from driver_utils import terminate_chrome_processes, initialize_driver
from excel_utils import read_excel_data
from form_utils import get_form_headers, fill_google_form
from matching_utils import match_headers
from config import EXCEL_FILE

def main():
    """Main function to orchestrate the automation process."""
    configure_logging()
    driver = None
    try:
        terminate_chrome_processes()
        driver = initialize_driver()
        form_headers = get_form_headers(driver)
        excel_headers, rows = read_excel_data(EXCEL_FILE)
        logging.info(f"Total rows to process: {len(rows)}")
        header_mapping, unmatched_headers = match_headers(excel_headers, form_headers)

        if not rows:
            logging.info("No data rows to process in Excel file")
            return

        for idx, row in enumerate(rows, start=2):
            logging.info(f"Processing row {idx}: {row}")
            success = fill_google_form(driver, row, excel_headers, header_mapping)
            logging.info(f"Row {idx} processed {'successfully' if success else 'with errors'}")
            # break  # Break after processing one row to avoid duplicates

    except Exception as e:
        logging.error(f"Main process error: {e}")
    finally:
        if driver:
            try:
                driver.quit()
                logging.info("WebDriver closed")
            except Exception as e:
                logging.error(f"Error closing WebDriver: {e}")

if __name__ == "__main__":
    main()