# utils.py
from datetime import datetime
from config import logging
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import WebDriverException

def scroll_into_view(driver, element):
    """Scroll an element into view smoothly."""
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
    time.sleep(0.2)

def clear_inputs(driver):
    """Clear all form inputs, including text, date, textarea, dropdowns, checkboxes, radio buttons, and file uploads."""
    logging.debug("Attempting to clear all form inputs.")

    try:
        # Step 1: Reset the form using JavaScript (works for most inputs)
        driver.execute_script("""
            var forms = document.getElementsByTagName('form');
            for (var i = 0; i < forms.length; i++) {
                forms[i].reset();
            }
        """)
        logging.debug("Executed JavaScript form reset.")

        # Step 2: Clear text, date, and textarea inputs explicitly
        text_inputs = driver.find_elements(By.XPATH, "//input[@type='text'] | //input[@type='date'] | //textarea")
        for field in text_inputs:
            try:
                scroll_into_view(driver, field)
                field.clear()
                if field.get_attribute("value"):
                    logging.warning(f"Failed to clear input: {field.get_attribute('outerHTML')[:50]}...")
                else:
                    logging.debug(f"Cleared text/textarea input: {field.get_attribute('outerHTML')[:50]}...")
            except WebDriverException as e:
                logging.warning(f"Error clearing text input: {e}")

        # Step 3: Reset dropdowns to first option (often blank)
        dropdowns = driver.find_elements(By.XPATH, "//div[@role='listbox'] | //select")
        for dropdown in dropdowns:
            try:
                scroll_into_view(driver, dropdown)
                dropdown.click()
                # Select the first option (assumed to be default/blank)
                first_option = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, "//div[@role='option'][1] | //option[1]"))
                )
                first_option.click()
                logging.debug(f"Reset dropdown to first option.")
            except WebDriverException as e:
                logging.warning(f"Error resetting dropdown: {e}")

        # Step 4: Uncheck checkboxes and radio buttons
        checked_inputs = driver.find_elements(By.XPATH, "//input[@type='checkbox' and @checked] | //input[@type='radio' and @checked]")
        for input_elem in checked_inputs:
            try:
                scroll_into_view(driver, input_elem)
                driver.execute_script("arguments[0].checked = false;", input_elem)
                logging.debug(f"Unchecked input: {input_elem.get_attribute('outerHTML')[:50]}...")
            except WebDriverException as e:
                logging.warning(f"Error unchecking input: {e}")

        # Step 5: Remove selected files (Google Forms specific)
        remove_file_buttons = driver.find_elements(By.XPATH, "//div[@aria-label='Remove']")
        for button in remove_file_buttons:
            try:
                scroll_into_view(driver, button)
                button.click()
                logging.debug("Removed selected file.")
            except WebDriverException as e:
                logging.warning(f"Error removing file: {e}")

        # Step 6: Verify form is cleared
        remaining_inputs = driver.find_elements(By.XPATH, "//input[@type='text' or @type='date'] | //textarea")
        for input_elem in remaining_inputs:
            if input_elem.get_attribute("value"):
                logging.warning(f"Input still has value after clearing: {input_elem.get_attribute('value')}")
        remaining_checked = driver.find_elements(By.XPATH, "//input[@type='checkbox' and @checked] | //input[@type='radio' and @checked]")
        if remaining_checked:
            logging.warning(f"Found {len(remaining_checked)} checked inputs after clearing.")

        logging.info("All form inputs cleared successfully.")
    except Exception as e:
        logging.error(f"Failed to clear inputs: {e}")
        # Fallback: Reload the form
        logging.info("Reloading form as fallback.")
        driver.get(driver.current_url)
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, "//form")))
        logging.debug("Form reloaded.")

def parse_date(value):
    """Convert Excel date to MM/DD/YYYY format."""
    if not value:
        return None
    if isinstance(value, datetime):
        return value.strftime("%m/%d/%Y")
    if isinstance(value, str):
        for fmt in ["%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y"]:
            try:
                return datetime.strptime(value, fmt).strftime("%m/%d/%Y")
            except ValueError:
                continue
        logging.warning(f"Invalid date format: {value}")
        return None