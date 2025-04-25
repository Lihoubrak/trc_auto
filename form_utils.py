import logging
import sys
import os
import time
from datetime import datetime
from pathlib import Path
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from bs4 import BeautifulSoup
import re
import unicodedata
from retrying import retry

# Assuming image_utils is a custom module
from image_utils import download_google_drive_image

# Configure logging
log_handlers = [logging.FileHandler("app.log", encoding="utf-8")]

# Only add StreamHandler if running in a console environment
if hasattr(sys, 'stdout') and sys.stdout is not None and hasattr(sys.stdout, 'encoding'):
    log_handlers.append(logging.StreamHandler(sys.stdout))

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=log_handlers
)
logger = logging.getLogger(__name__)

# Safely handle UTF-8 encoding for stdout
if hasattr(sys, 'stdout') and sys.stdout is not None and hasattr(sys.stdout, 'encoding'):
    try:
        if sys.stdout.encoding.lower() != 'utf-8':
            sys.stdout.reconfigure(encoding='utf-8')
    except Exception as e:
        logger.warning(f"Failed to reconfigure stdout to UTF-8: {e}")

# Field configuration
FIELD_TYPES = {
    "date": {
        "keywords": ["Date of Damage", "Finished Date of Repairing"],
        "handler": "handle_date_field"
    },
    "checkbox": {
        "keywords": ["Email", "Lat/Long"],
        "handler": "handle_checkbox_field"
    },
    "dropdown": {
        "keywords": [
            "Requested Company", "Type of Infrastructure",
            "Overhead or Underground", "ខេត្ត/ក្រុង"
        ],
        "handler": "handle_dropdown_field"
    },
    "text": {
        "keywords": [
            "Repair for company/customers", "Starting Address", "Ending Address",
            "Start: Lat ,Long", "End: Lat ,Long", "Length of replacement broken cable",
            "Cable Incident"
        ],
        "handler": "handle_text_field"
    }
}

def scroll_into_view(driver, element):
    """Scroll an element into view smoothly."""
    driver.execute_script(
        "arguments[0].scrollIntoView({block: 'center', behavior: 'smooth'});", element
    )
    time.sleep(0.1)

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
    logger.warning(f"Invalid date format: {value}")
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
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//span[@class='M7eMe']"))
        )
        headers = [
            normalize_text(elem.text)
            for elem in driver.find_elements(By.XPATH, "//span[@class='M7eMe']")
            if elem.text.strip()
        ]
        if not headers:
            raise ValueError("No headers found in form")
        logger.info(f"Retrieved {len(headers)} form headers")
        return headers
    except Exception as e:
        logger.error(f"Failed to fetch form headers: {e}")
        raise

def handle_date_field(driver, form_header, value, form_header_cleaned):
    """Handle date input fields."""
    date_value = parse_date(value)
    if not date_value:
        logger.warning(f"Invalid date value for '{form_header}': {value}")
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
                By.XPATH,
                f"//*[contains(normalize-space(.), '{form_header_cleaned[:50]}')]"
                f"/ancestor::div[@role='listitem']"
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
        logger.info(f"Filled date field '{form_header}' with '{date_value}'")
        return True
    except Exception as e:
        logger.error(f"Error filling date field '{form_header}': {e}")
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
            checkbox = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, checkbox_xpath))
            )
            scroll_into_view(driver, checkbox)
            if checkbox.get_attribute("aria-checked") != "true":
                driver.execute_script("arguments[0].click();", checkbox)
                if checkbox.get_attribute("aria-checked") != "true":
                    logger.warning(f"Failed to check checkbox '{val}' for '{form_header}'")
                    success = False
            logger.info(f"Checked checkbox '{val}' for '{form_header}'")
        except Exception as e:
            logger.error(f"Error checking checkbox '{val}' for '{form_header}': {e}")
            success = False
    return success

def handle_dropdown_field(driver, form_header, value, form_header_cleaned):
    """Handle dropdown fields."""
    dropdown_xpath = (
        f"//*[contains(normalize-space(.), '{form_header_cleaned[:50]}')]"
        f"/ancestor::div[@role='listitem']//div[@role='listbox']"
    )
    try:
        dropdown = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, dropdown_xpath))
        )
        dropdown.click()
        option_xpath = f"//div[@role='option' and contains(normalize-space(.), '{str(value)[:30]}')]"
        option = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, option_xpath))
        )
        scroll_into_view(driver, option)
        driver.execute_script("arguments[0].click();", option)
        selected = driver.find_element(By.XPATH, dropdown_xpath).text
        if str(value) not in selected:
            logger.warning(f"Failed to select dropdown option '{value}' for '{form_header}'")
            return False
        logger.info(f"Selected dropdown option '{value}' for '{form_header}'")
        return True
    except Exception as e:
        logger.error(f"Error selecting dropdown option for '{form_header}': {e}")
        return False

def handle_text_field(driver, form_header, value, form_header_cleaned):
    """Handle text and textarea fields."""
    xpath = (
        f"//*[contains(normalize-space(.), '{form_header_cleaned[:50]}')]"
        f"/ancestor::div[@role='listitem']//*[(self::input[@type='text' or @type='number'] or self::textarea)]"
    )
    try:
        element = WebDriverWait(driver, 3).until(
            EC.element_to_be_clickable((By.XPATH, xpath))
        )
        scroll_into_view(driver, element)
        element.clear()
        element.send_keys(str(value))
        print(f"Type of value: {type(value)}")
        if element.get_attribute("value") != str(value):
            driver.execute_script("arguments[0].value = arguments[1];", element, str(value))
            if element.get_attribute("value") != str(value):
                logger.warning(f"Failed to fill text field '{form_header}' with '{value}'")
                return False
        logger.info(f"Filled text field '{form_header}' with '{value}'")
        return True
    except Exception as e:
        logger.error(f"Error filling text field '{form_header}': {e}")
        return False

def fill_form_field(driver, form_header, value, form_header_cleaned):
    """Fill a single form field based on header content, excluding file uploads."""
    for field_type, config in FIELD_TYPES.items():
        if any(keyword in form_header_cleaned for keyword in config["keywords"]):
            handler = globals()[config["handler"]]
            return handler(driver, form_header, value, form_header_cleaned)
    logger.warning(f"Unknown field type for header: {form_header}")
    return False

@retry(stop_max_attempt_number=3, wait_fixed=2000)
def upload_file(driver, form_header, form_header_cleaned, temp_file_path):
    """Attempt to upload a file with retries."""
    file_name = os.path.basename(temp_file_path)
    try:
        driver.switch_to.default_content()
        # upload_btn_xpath = (
        #     f"//span[@class='M7eMe' and contains(normalize-space(.), '{form_header_cleaned[:50]}')]"
        #     f"/ancestor::div[@role='listitem']//div[@role='button' and contains(@aria-label, 'Add File')]"
        # )
#         upload_btn_xpath = (
#     f"//span[@class='M7eMe' and contains(normalize-space(.), '{form_header_cleaned[:50]}')]"
#     f"/ancestor::div[@role='listitem']//div[@role='button' and "
#     f"(@aria-label='Add File' or contains(@class, 'uArJ5e') or contains(@class, 'cd29Sd') or @jsname='mWZCyf')]"
# )
        upload_btn_xpath = (
            f"//span[@class='M7eMe' and contains(normalize-space(.), '{form_header_cleaned[:50]}')]"
            f"/ancestor::div[@role='listitem']//div[@role='button' and ("
            f"@aria-label='Add File' or "
            f"contains(@class, 'uArJ5e') or "
            f"contains(@class, 'cd29Sd') or "
            f".//span[contains(@class, 'NPEfkd') and contains(@class, 'RveJvd') and contains(@class, 'snByac') and contains(., 'Add File')]"
            f")]"
        )
        upload_button = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, upload_btn_xpath))
        )
        scroll_into_view(driver, upload_button)
        driver.execute_script("arguments[0].click();", upload_button)
        logger.info(f"Clicked 'Add File' button for '{form_header}'")

        picker_dialog_xpath = "//div[contains(@class, 'picker-dialog') and not(contains(@style, 'display: none'))]"
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, picker_dialog_xpath))
        )

        iframe_xpath = f"{picker_dialog_xpath}//iframe[contains(@src, 'docs.google.com/picker')]"
        iframes = WebDriverWait(driver, 5).until(
            EC.presence_of_all_elements_located((By.XPATH, iframe_xpath))
        )
        if not iframes:
            raise Exception("No iframe found for file picker")

        iframe = iframes[-1]
        iframe_id = iframe.get_attribute("id")
        logger.info(f"Switching to iframe with id: {iframe_id}")
        driver.switch_to.frame(iframe)

        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.XPATH, "//input[@type='file']"))
        )
        file_input = driver.find_element(By.XPATH, "//input[@type='file']")
        file_input.send_keys(temp_file_path)
        logger.info(f"Sent file path to file input: {temp_file_path}")

        driver.switch_to.default_content()
        time.sleep(2)

        file_list_xpath = (
            f"//span[@class='M7eMe' and contains(normalize-space(.), '{form_header_cleaned[:50]}')]"
            f"/ancestor::div[@role='listitem']//div[@role='listitem']//div[contains(text(), '{file_name}')]"
        )
        file_element = WebDriverWait(driver, 45).until(
            EC.presence_of_element_located((By.XPATH, file_list_xpath))
        )
        displayed_file_name = file_element.text.strip()

        if file_name.lower() in displayed_file_name.lower():
            logger.info(f"File name matched: expected '{file_name}', got '{displayed_file_name}'")
            return True
        logger.warning(f"File name mismatch: expected '{file_name}', got '{displayed_file_name}'")
        return False
    except TimeoutException as te:
        logger.error(f"Timeout during file upload attempt for '{form_header}': {te}")
        driver.switch_to.default_content()
        raise
    except Exception as e:
        logger.error(f"Error during file upload attempt for '{form_header}': {e}")
        driver.switch_to.default_content()
        raise

def fill_google_form(driver, row, headers, header_mapping, config):
    """Fill and submit a Google Form for one row of data."""
    temp_dir = Path("images")
    temp_dir.mkdir(exist_ok=True)
    temp_files = []
    fields_filled = True

    try:
        driver.get(config["GOOGLE_FORM_URL"])
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, "//form")))
        logger.info("Google Form loaded successfully")

        # Handle email checkbox once
        try:
            checkbox_xpath = '//div[.//span[text()="Email"]]/following::div[@role="checkbox"][1]'
            checkbox = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, checkbox_xpath))
            )
            scroll_into_view(driver, checkbox)
            if checkbox.get_attribute("aria-checked") != "true":
                checkbox.click()
                logger.info("Checked 'Email' collection checkbox")
            else:
                logger.info("Email checkbox already checked, skipping")
        except TimeoutException:
            logger.info("No email checkbox found — skipping")
        except Exception as e:
            logger.error(f"Error handling email checkbox: {e}")

        # Process each header
        for excel_header, value in zip(headers, row):
            if excel_header not in header_mapping:
                logger.info(f"Skipping empty or unmapped field: {excel_header}")
                continue
            form_header = header_mapping[excel_header]
            form_header_cleaned = normalize_text(form_header)
            logger.info(f"Processing field: {form_header}")

            # Handle file upload fields
            if "Picture of Damage Cable" in form_header or "Picture of drawing in google map" in form_header:
                if isinstance(value, str) and "drive.google.com" in value:
                    start_time = time.time()
                    temp_file_path = download_google_drive_image(value,driver, temp_dir=str(temp_dir))
                    download_duration = time.time() - start_time
                    logger.info(f"Download took {download_duration:.2f} seconds for URL: {value}")

                    if temp_file_path:
                        temp_files.append(temp_file_path)
                        try:
                            if upload_file(driver, form_header, form_header_cleaned, temp_file_path):
                                logger.info(f"Successfully uploaded file for '{form_header}'")
                            else:
                                logger.error(f"Failed to upload file for '{form_header}'")
                                fields_filled = False
                        except Exception as e:
                            logger.error(f"Failed to upload file for '{form_header}' after retries: {e}")
                            fields_filled = False
                    else:
                        logger.warning(f"Failed to download image from Google Drive for '{form_header}': {value}")
                        fields_filled = False
                else:
                    logger.warning(f"Invalid Google Drive URL for image field '{form_header}': {value}")
                    fields_filled = False
                time.sleep(0.5)
                continue

            # Handle special case for "Number of cable * Core"
            if "Number of cable * Core" in form_header_cleaned:
                xpath_text_other = (
                    f"//div[@role='heading' and contains(., '{form_header_cleaned.split()[0]}')]"
                    f"/ancestor::div[@role='listitem']//input[@type='text' or @type='number']"
                )
                try:
                    input_elements = driver.find_elements(By.XPATH, xpath_text_other)
                    if input_elements:
                        scroll_into_view(driver, input_elements[0])
                        input_elements[0].clear()
                        input_elements[0].send_keys(str(value))
                        time.sleep(0.5)
                        if input_elements[0].get_attribute("value") == str(value):
                            logger.info(f"Filled 'Number of cable * Core' with value: {value}")
                        else:
                            logger.warning(f"Failed to fill 'Number of cable * Core' with value: {value}")
                            fields_filled = False
                    else:
                        logger.warning("No input element found for 'Number of cable * Core'")
                        fields_filled = False
                except Exception as e:
                    logger.error(f"Error filling 'Number of cable * Core': {e}")
                    fields_filled = False
                continue

            # Fill other fields
            if not fill_form_field(driver, form_header, value, form_header_cleaned):
                logger.warning(f"Failed to fill field '{form_header}' with value '{value}'")
                fields_filled = False
            time.sleep(0.2)

        # Submit the form if all fields were filled successfully
        if fields_filled:
             time.sleep(2)
        try:
            submit_btn = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//span[text()='Submit']/ancestor::div[@role='button']"))
            )
            scroll_into_view(driver, submit_btn)
            if submit_btn.get_attribute("aria-disabled") == "true":
                logger.error("Submit button is disabled, likely due to unfilled required fields")
                # Ghi log các trường bắt buộc còn trống
                required_fields = driver.find_elements(By.XPATH, "//*[contains(@aria-required, 'true')]")
                for field in required_fields:
                    value = field.get_attribute("value") or field.text
                    if not value:
                        logger.warning(f"Required field empty: {field.get_attribute('aria-label')}")
                return False
            driver.execute_script("arguments[0].click();", submit_btn)
            WebDriverWait(driver, 60).until(EC.url_contains("formResponse"))
            logger.info("Form submitted successfully")
            return True
        except Exception as e:
            logger.error(f"Form submission failed: {e}", exc_info=True)
            return False

    except Exception as e:
        logger.error(f"Error while filling the form: {e}")
        fields_filled = False

    finally:
        driver.switch_to.default_content()
        for f in temp_files:
            try:
                os.remove(f)
                logger.info(f"Deleted temp file: {f}")
            except Exception as e:
                logger.warning(f"Failed to delete temp file: {f} - {e}")

    return fields_filled