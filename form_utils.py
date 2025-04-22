import logging
import sys
from datetime import datetime
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from bs4 import BeautifulSoup
import re
import unicodedata
from retrying import retry
import time
from image_utils import download_google_drive_image
import os
# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)]
)

# Ensure UTF-8 encoding
if sys.stdout.encoding.lower() != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

def scroll_into_view(driver, element):
    """Scroll element into view smoothly."""
    driver.execute_script(
        "arguments[0].scrollIntoView({block: 'center', behavior: 'smooth'});",
        element
    )

def parse_date(value):
    """Convert date to MM/DD/YYYY format."""
    if not value:
        return None
    if isinstance(value, datetime):
        return value.strftime("%m/%d/%Y")
    
    date_formats = ["%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y"]
    for fmt in date_formats:
        try:
            return datetime.strptime(str(value), fmt).strftime("%m/%d/%Y")
        except ValueError:
            continue
    logging.warning(f"Invalid date format: {value}")
    return None

def normalize_text(text):
    """Normalize text for comparison."""
    text = unicodedata.normalize("NFC", text).replace("\u200b", "").replace("\u00a0", " ")
    soup = BeautifulSoup(text, "html.parser")
    text = soup.get_text(separator=" ").strip()
    return re.sub(r"\s+", " ", text)

def get_form_headers(driver, config):
    """Retrieve and normalize Google Form headers."""
    try:
        driver.get(config["GOOGLE_FORM_URL"])
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//span[@class='M7eMe']")))
        headers = [
            normalize_text(elem.text)
            for elem in driver.find_elements(By.XPATH, "//span[@class='M7eMe']")
            if elem.text.strip()
        ]
        if not headers:
            raise ValueError("No headers found in form")
        logging.info(f"Retrieved {len(headers)} form headers")
        return headers
    except Exception as e:
        logging.error(f"Failed to fetch form headers: {e}")
        raise

def handle_file_upload(driver, form_header, value, form_header_cleaned):
    """Handle file upload fields for Google Forms with a single URL."""
    url = str(value).strip() if value else ""
    if not url:
        logging.warning(f"No valid Google Drive URL provided for '{form_header}': {value}")
        return False

    if "drive.google.com" not in url:
        logging.warning(f"Invalid Google Drive URL for '{form_header}': {url}")
        return False

    temp_dir = "images"
    try:
        start_time = time.time()
        temp_file_path = download_google_drive_image(url, temp_dir=temp_dir)
        download_duration = time.time() - start_time

        # Validate downloaded file
        if not temp_file_path or not os.path.exists(temp_file_path):
            logging.error(f"Failed to download file for '{form_header}': {url}")
            return False
        if os.path.getsize(temp_file_path) == 0:
            logging.error(f"Downloaded file is empty for '{form_header}': {temp_file_path}")
            return False
        if not os.path.dirname(temp_file_path).endswith(temp_dir):
            logging.error(f"File not in expected temp directory '{temp_dir}': {temp_file_path}")
            return False

        # Verify file accessibility
        with open(temp_file_path, 'rb') as f:
            f.read(1)
        logging.info(f"Download took {download_duration:.2f} seconds for URL: {url}")
        file_name = os.path.basename(temp_file_path)

    except Exception as e:
        logging.error(f"Error downloading file for '{form_header}': {url}, Error: {e}")
        return False

    # Upload file
    @retry(stop_max_attempt_number=3, wait_fixed=1000)
    def upload_single_file(temp_file_path, file_name):
        try:
            upload_btn_xpath = (
                f"//span[@class='M7eMe' and contains(normalize-space(.), '{form_header_cleaned[:50]}')]"
                f"/ancestor::div[@role='listitem']//div[@role='button' and contains(@class, 'uArJ5e') and contains(@aria-label, 'Add file')]"
            )
            upload_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, upload_btn_xpath))
            )
            scroll_into_view(driver, upload_button)
            driver.execute_script("arguments[0].click();", upload_button)

            picker_dialog_xpath = "//div[contains(@class, 'fFW7wc XKSfm-Sx9Kwc picker-dialog') and not(contains(@style, 'display: none'))]"
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, picker_dialog_xpath))
            )

            iframe_xpath = f"{picker_dialog_xpath}//iframe[contains(@src, 'docs.google.com/picker')]"
            WebDriverWait(driver, 10).until(
                EC.frame_to_be_available_and_switch_to_it((By.XPATH, iframe_xpath))
            )

            file_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//input[@type='file']"))
            )
            file_input.send_keys(temp_file_path)
            driver.switch_to.default_content()

            file_list_xpath = (
                f"//span[@class='M7eMe' and contains(normalize-space(.), '{form_header_cleaned[:50]}')]"
                f"/ancestor::div[@role='listitem']//div[@role='listitem']//div[contains(text(), '{file_name}')]"
            )
            file_element = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, file_list_xpath))
            )
            displayed_file_name = file_element.text.strip()
            if file_name.lower() not in displayed_file_name.lower():
                logging.warning(f"File name mismatch for '{form_header}': expected '{file_name}', got '{displayed_file_name}'")
                return False

            logging.info(f"Successfully uploaded file '{file_name}' for '{form_header}'")
            return True

        except TimeoutException as te:
            logging.error(f"Timeout during file upload for '{form_header}': {te}")
            driver.switch_to.default_content()
            raise
        except Exception as e:
            logging.error(f"Error during file upload for '{form_header}': {e}")
            driver.switch_to.default_content()
            raise

    try:
        if upload_single_file(temp_file_path, file_name):
            logging.info(f"File uploaded successfully for '{form_header}'")
            return True
        else:
            logging.error(f"File upload failed for '{form_header}'")
            return False
    except Exception as e:
        logging.error(f"File upload failed for '{form_header}' after retries: {e}")
        return False

def handle_date_field(driver, form_header, value, form_header_cleaned):
    """Handle date input fields."""
    date_value = parse_date(value)
    if not date_value:
        logging.warning(f"Invalid date value for '{form_header}': {value}")
        return False
    
    day, month, year = date_value.split("/")
    date_input_xpath = (
        f"//*[contains(normalize-space(.), '{form_header_cleaned[:50]}')]"
        f"/ancestor::div[@role='listitem']//input[@type='date']"
    )
    try:
        date_inputs = driver.find_elements(By.XPATH, date_input_xpath)
        if date_inputs:
            scroll_into_view(driver, date_inputs[0])
            date_inputs[0].send_keys(f"{month}{day}{year}")
        else:
            date_container = driver.find_element(
                By.XPATH, f"//*[contains(normalize-space(.), '{form_header_cleaned[:50]}')]/ancestor::div[@role='listitem']"
            )
            inputs = {
                "month": date_container.find_element(By.XPATH, ".//input[@aria-label='Month']"),
                "day": date_container.find_element(By.XPATH, ".//input[@aria-label='Day of the month']"),
                "year": date_container.find_element(By.XPATH, ".//input[@aria-label='Year']")
            }
            for field, val in zip(inputs, [month, day, year]):
                scroll_into_view(driver, inputs[field])
                inputs[field].clear()
                inputs[field].send_keys(val)
        logging.info(f"Filled date field '{form_header}' with '{date_value}'")
        return True
    except Exception as e:
        logging.error(f"Error filling date field '{form_header}': {e}")
        return False

def handle_checkbox_field(driver, form_header, value, form_header_cleaned):
    """Handle checkbox fields."""
    checkbox_values = [v.strip() for v in str(value).split(",") if v.strip()]
    success = True
    for val in checkbox_values:
        checkbox_xpath = (
            f"//*[contains(normalize-space(.), '{form_header_cleaned[:50]}')]"
            f"/ancestor::div[@role='listitem']//div[@role='checkbox' and @data-answer-value='{val}']"
        )
        try:
            checkbox = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, checkbox_xpath)))
            scroll_into_view(driver, checkbox)
            if checkbox.get_attribute("aria-checked") != "true":
                driver.execute_script("arguments[0].click();", checkbox)
                if checkbox.get_attribute("aria-checked") != "true":
                    logging.warning(f"Failed to check checkbox '{val}' for '{form_header}'")
                    success = False
            logging.info(f"Checked checkbox '{val}' for '{form_header}'")
        except Exception as e:
            logging.error(f"Error checking checkbox '{val}' for '{form_header}': {e}")
            success = False
    return success

def handle_dropdown_field(driver, form_header, value, form_header_cleaned):
    """Handle dropdown fields."""
    dropdown_xpath = (
        f"//*[contains(normalize-space(.), '{form_header_cleaned[:50]}')]"
        f"/ancestor::div[@role='listitem']//div[@role='listbox']"
    )
    try:
        dropdown = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, dropdown_xpath)))
        # scroll_into_view(driver, dropdown)
        dropdown.click()
        option_xpath = f"//div[@role='option' and contains(normalize-space(.), '{str(value)[:30]}')]"
        option = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, option_xpath)))
        scroll_into_view(driver, option)
        driver.execute_script("arguments[0].click();", option)
        selected = driver.find_element(By.XPATH, dropdown_xpath).text
        if str(value) not in selected:
            logging.warning(f"Failed to select dropdown option '{value}' for '{form_header}'")
            return False
        logging.info(f"Selected dropdown option '{value}' for '{form_header}'")
        return True
    except Exception as e:
        logging.error(f"Error selecting dropdown option for '{form_header}': {e}")
        return False

def handle_text_field(driver, form_header, value, form_header_cleaned):
    print("...form_header_cleaned",form_header_cleaned)
    """Handle text and textarea fields."""
    xpath = (
        f"//*[contains(normalize-space(.), '{form_header_cleaned[:50]}')]"
        f"/ancestor::div[@role='listitem']//*[(self::input[@type='text' or @type='number'] or self::textarea)]"
    )
    try:
        element = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, xpath)))
        scroll_into_view(driver, element)
        element.clear()
        element.send_keys(str(value))
        if element.get_attribute("value") != str(value):
            driver.execute_script("arguments[0].value = arguments[1];", element, str(value))
            if element.get_attribute("value") != str(value):
                logging.warning(f"Failed to fill text field '{form_header}' with '{value}'")
                return False
        logging.info(f"Filled text field '{form_header}' with '{value}'")
        return True
    except Exception as e:
        logging.error(f"Error filling text field '{form_header}': {e}")
        return False

def fill_field(driver, form_header, value, form_header_cleaned):
    """Fill a single form field based on header content."""
    try:
        field_configs = {
            "file_upload": {
                "keywords": ["Picture of Damage Cable", "Picture of drawing in google map"],
                "handler": handle_file_upload
            },
            "date": {
                "keywords": ["Date of Damage", "Finished Date of Repairing"],
                "handler": handle_date_field
            },
            "checkbox": {
                "keywords": ["Email", "Lat/Long"],
                "handler": handle_checkbox_field
            },
            "dropdown": {
                "keywords": [
                    "Requested Company", "Type of Infrastructure",
                    "Overhead or Underground", "ខេត្ត/ក្រុង"
                ],
                "handler": handle_dropdown_field
            },
            "text": {
                "keywords": [
                    "Repair for company/customers", "Number of cable * Core",
                    "Starting Address", "Ending Address", "Start: Lat ,Long",
                    "End: Lat ,Long", "Length of replacement broken cable",
                    "Cable Incident"
                ],
                "handler": handle_text_field
            }
        }

        for config in field_configs.values():
            if any(keyword in form_header_cleaned for keyword in config["keywords"]):
                return config["handler"](driver, form_header, value, form_header_cleaned)

        logging.warning(f"Unknown field type for header: {form_header}")
        return False

    except Exception as e:
        logging.error(f"Error filling field '{form_header}': {e}")
        return False

def fill_google_form(driver, row, headers, header_mapping, config):
    """Fill and submit a Google Form for a single row of data."""
    try:
        driver.get(config["GOOGLE_FORM_URL"])
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//form")))
        logging.info("Google Form loaded")

        # Handle email checkbox
        try:
            checkbox_xpath = '//div[.//span[text()="Email"]]/following::div[@role="checkbox"][1]'
            checkbox = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH, checkbox_xpath)))
            scroll_into_view(driver, checkbox)
            if checkbox.get_attribute("aria-checked") != "true":
                checkbox.click()
                logging.info("Checked 'Email' checkbox")
        except (TimeoutException, NoSuchElementException):
            logging.info("No email checkbox found")

        # Fill form fields
        fields_filled = True
        for excel_header, value in zip(headers, row):
            if excel_header not in header_mapping or not value:
                logging.info(f"Skipping field: {excel_header}")
                continue
            form_header = header_mapping[excel_header]
            form_header_cleaned = normalize_text(form_header)
            logging.info(f"Processing field: {form_header}")
            if not fill_field(driver, form_header, value, form_header_cleaned):
                fields_filled = False

        # Submit form
        if fields_filled:
            submit_btn = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//span[text()='Submit']/ancestor::div[@role='button']"))
            )
            scroll_into_view(driver, submit_btn)
            driver.execute_script("arguments[0].click();", submit_btn)
            WebDriverWait(driver, 10).until(EC.url_contains("formResponse"))
            logging.info("Form submitted successfully")
            return True
        else:
            logging.warning("Form not submitted due to field errors")
            return False

    except Exception as e:
        logging.error(f"Form filling failed: {e}")
        return False