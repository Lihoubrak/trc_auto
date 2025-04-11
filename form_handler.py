# form_handler.py
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from retrying import retry
from config import logging, GOOGLE_FORM_URL, TEMP_DIR, SIMILARITY_THRESHOLD
from utils import scroll_into_view, clear_inputs, parse_date
from file_downloader import download_google_drive_file
from header_matcher import fuzzy_match_value

def get_form_headers(driver):
    """Fetch headers (question labels) from the Google Form."""
    try:
        driver.get(GOOGLE_FORM_URL)
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, "//span[@class='M7eMe']")))
        header_elements = driver.find_elements(By.XPATH, "//span[@class='M7eMe']")
        headers = [elem.text.strip() for elem in header_elements if elem.text.strip()]
        if not headers:
            raise ValueError("No form headers found.")
        logging.info(f"Found form headers: {headers}")
        return headers
    except Exception as e:
        logging.error(f"Error fetching form headers: {e}")
        raise

def fill_date_field(driver, form_header, value):
    """Fill a date field in the form."""
    try:
        date_value = parse_date(value)
        if not date_value:
            return False

        month, day, year = date_value.split("/")
        date_input_xpath = f"//span[@class='M7eMe' and normalize-space(.)='{form_header}']/ancestor::div[@role='listitem']//input[@type='date']"
        date_inputs = driver.find_elements(By.XPATH, date_input_xpath)
        if date_inputs:
            scroll_into_view(driver, date_inputs[0])
            date_inputs[0].send_keys(f"{month}{day}{year}")
            logging.info(f"Filled date field '{form_header}' with '{date_value}'")
            return True

        date_container_xpath = f"//span[@class='M7eMe' and normalize-space(.)='{form_header}']/ancestor::div[@role='listitem']"
        date_container = driver.find_element(By.XPATH, date_container_xpath)
        month_input = date_container.find_element(By.XPATH, ".//input[@aria-label='Month']")
        day_input = date_container.find_element(By.XPATH, ".//input[@aria-label='Day of the month']")
        year_input = date_container.find_element(By.XPATH, ".//input[@aria-label='Year']")
        scroll_into_view(driver, month_input)
        month_input.clear()
        month_input.send_keys(month)
        day_input.clear()
        day_input.send_keys(day)
        year_input.clear()
        year_input.send_keys(year)
        logging.info(f"Filled date field '{form_header}' with '{date_value}'")
        return True
    except Exception as e:
        logging.warning(f"Failed to fill date field '{form_header}': {e}")
        return False

@retry(stop_max_attempt_number=3, wait_fixed=2000)
def upload_file_via_iframe(driver, wait, file_path):
    """Upload a file through Google's iframe picker."""
    try:
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.XPATH, "//iframe[contains(@src, 'docs.google.com/picker')]")))
        file_input = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='file']")))
        file_input.send_keys(file_path)
        insert_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@role='dialog']//span[text()='Insert']/ancestor::button")))
        insert_btn.click()
        driver.switch_to.default_content()
        wait.until(EC.presence_of_element_located((By.XPATH, "//div[@role='list' and @aria-label='Selected files']//div")))
        logging.info(f"Uploaded file '{file_path}' successfully.")
        return True
    except Exception as e:
        logging.warning(f"File upload failed: {e}")
        driver.switch_to.default_content()
        raise

def fill_file_upload(driver, form_header, file_url):
    """Handle file upload for a form field."""
    import os
    try:
        if not file_url or "drive.google.com" not in file_url:
            logging.warning(f"Invalid file URL for '{form_header}': {file_url}")
            return False

        downloaded_file = download_google_drive_file(file_url)
        if not downloaded_file:
            return False

        button_xpath = f"//span[@class='M7eMe' and normalize-space(.)='{form_header}']/ancestor::div[@role='listitem']//div[@role='button' and contains(@aria-label, 'Add File')]"
        upload_button = driver.find_element(By.XPATH, button_xpath)
        scroll_into_view(driver, upload_button)
        upload_button.click()

        if upload_file_via_iframe(driver, WebDriverWait(driver, 15), downloaded_file):
            logging.info(f"File uploaded for '{form_header}'")
            os.remove(downloaded_file)
            return True
        return False
    except Exception as e:
        logging.warning(f"Failed to upload file for '{form_header}': {e}")
        return False

def check_required_fields(driver):
    """Ensure all required fields are filled."""
    try:
        required_fields = driver.find_elements(By.XPATH, "//span[@class='NPEfkd' and text()='*']")
        for field in required_fields:
            parent = field.find_element(By.XPATH, "./ancestor::div[@role='listitem']")
            inputs = parent.find_elements(By.XPATH, ".//input[@type='text' or @type='date'] | .//div[@role='listbox'] | .//div[@role='list' and @aria-label='Selected files']")
            if not inputs:
                logging.warning("Required field has no input element.")
                return False
            for inp in inputs:
                if inp.get_attribute("value") == "" and "Selected files" not in inp.get_attribute("aria-label"):
                    logging.warning("Required field is empty.")
                    return False
        logging.debug("All required fields appear filled.")
        return True
    except Exception as e:
        logging.warning(f"Error checking required fields: {e}")
        return False

@retry(stop_max_attempt_number=3, wait_fixed=2000)
def submit_form(driver):
    """Submit the form with retries."""
    try:
        submit_btn = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Submit')]/ancestor::div[@role='button']"))
        )
        scroll_into_view(driver, submit_btn)
        if not submit_btn.is_enabled():
            logging.warning("Submit button disabled, likely missing required fields.")
            return False
        submit_btn.click()
        WebDriverWait(driver, 15).until(EC.url_contains("formResponse"))
        logging.info("✅ Form submitted successfully!")
        return True
    except Exception as e:
        logging.error(f"Form submission failed: {e}")
        return False

def fill_google_form(driver, row, headers, header_mapping):
    try:
        driver.get(GOOGLE_FORM_URL)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//form")))
        clear_inputs(driver)

        for excel_header, value in zip(headers, row):
            if excel_header not in header_mapping or value is None:
                continue

            form_header = header_mapping[excel_header]
            value_str = str(value).strip() if value != "" else ""
            logging.info(f"Processing field '{form_header}' with value '{value_str}'")

            # Check if the field is a date field
            date_input_xpath = (f"//span[@class='M7eMe' and normalize-space(.)='{form_header}']/ancestor::div[@role='listitem']//"
                               f"input[@type='date' or @aria-label='Month' or @aria-label='Day of the month' or @aria-label='Year']")
            date_inputs = driver.find_elements(By.XPATH, date_input_xpath)
            if date_inputs:
                if fill_date_field(driver, form_header, value):
                    continue

            # Check if the field is a file upload
            file_button_xpath = f"//span[@class='M7eMe' and normalize-space(.)='{form_header}']/ancestor::div[@role='listitem']//div[@role='button' and contains(@aria-label, 'Add File')]"
            file_buttons = driver.find_elements(By.XPATH, file_button_xpath)
            if file_buttons:
                if fill_file_upload(driver, form_header, value_str):
                    continue

            try:
                # Try text input field
                input_xpath = f"//span[@class='M7eMe' and normalize-space(.)='{form_header}']/ancestor::div[@role='listitem']//input[@type='text']"
                inputs = driver.find_elements(By.XPATH, input_xpath)
                if inputs:
                    scroll_into_view(driver, inputs[0])
                    inputs[0].clear()
                    inputs[0].send_keys(value_str)
                    logging.info(f"Filled text field '{form_header}' with '{value_str}'")
                    continue

                # Try dropdown
                dropdown_xpath = f"//span[@class='M7eMe' and normalize-space(.)='{form_header}']/ancestor::div[@role='listitem']//div[@role='listbox']"
                dropdowns = driver.find_elements(By.XPATH, dropdown_xpath)
                if dropdowns:
                    scroll_into_view(driver, dropdowns[0])
                    dropdowns[0].click()
                    time.sleep(0.5)
                    options = driver.find_elements(By.XPATH, "//div[@role='option']")
                    for opt in options:
                        if value_str.lower() in opt.text.lower():
                            scroll_into_view(driver, opt)
                            opt.click()
                            logging.info(f"Selected dropdown option '{opt.text}' for '{form_header}'")
                            break
                    continue

                # Try checkbox
                checkbox_xpath = f"//div[@role='listitem']//span[normalize-space(.)='{value_str}']/ancestor::label"
                checkboxes = driver.find_elements(By.XPATH, checkbox_xpath)
                if checkboxes:
                    scroll_into_view(driver, checkboxes[0])
                    checkboxes[0].click()
                    logging.info(f"Checked checkbox '{form_header}' with value '{value_str}'")
                    continue

                logging.warning(f"No matching input found for field '{form_header}' with value '{value_str}'")

            except Exception as e:
                logging.warning(f"Failed to process field '{form_header}' with value '{value_str}': {e}")
                continue

        try:
            submit_xpath = "//span[contains(text(), 'Submit')]/ancestor::div[@role='button']"
            submit_btn = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, submit_xpath)))
            scroll_into_view(driver, submit_btn)
            submit_btn.click()
            WebDriverWait(driver, 10).until(EC.url_contains("formResponse"))
            logging.info("✅ Form submitted successfully.")
        except TimeoutException:
            logging.error("❌ Submit button not clickable or form submission failed.")
        except Exception as e:
            logging.error(f"❌ Submit failed: {e}")

    except Exception as e:
        logging.error(f"Error filling form: {e}")