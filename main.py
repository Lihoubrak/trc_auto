import logging
import os
import shutil
from datetime import datetime
import tempfile
import time
from typing import List, Tuple, Dict, Optional

import openpyxl
import psutil
import requests
from fuzzywuzzy import fuzz
from retrying import retry
from selenium import webdriver
from selenium.common.exceptions import WebDriverException, TimeoutException, NoSuchElementException, ElementClickInterceptedException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

try:
    from webdriver_manager.chrome import ChromeDriverManager
    USE_WEBDRIVER_MANAGER = True
except ImportError:
    USE_WEBDRIVER_MANAGER = False

# Configuration
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger(__name__)

GOOGLE_FORM_URL = "https://docs.google.com/forms/d/e/1FAIpQLSeqWvnn4KIru5BYd6aNVTCvaej6KvPWdbK0tN3piOgU8u8ftg/viewform?usp=dialog"
EXCEL_FILE = "MAINTENANCE CABLE REQUEST TO VTC.xlsx"
SIMILARITY_THRESHOLD = 80
TEMP_DIR = tempfile.mkdtemp()
MAX_FILE_SIZE = 10 * 1024 * 1024  # 10MB limit
USER_DATA_DIR = r"C:\Users\KHC\AppData\Local\Google\Chrome\User Data"
PROFILE_DIR = "Profile 1"

class FormAutomationError(Exception):
    """Custom exception for form automation errors."""
    pass

def kill_chrome_processes() -> None:
    """Terminate all Chrome processes to prevent profile conflicts."""
    try:
        for proc in psutil.process_iter(['pid', 'name']):
            if proc.info['name'].lower() == "chrome.exe":
                proc.kill()
                logger.debug(f"Killed Chrome process: {proc.pid}")
    except Exception as e:
        logger.warning(f"Error killing Chrome processes: {e}")

def setup_driver() -> webdriver.Chrome:
    """Initialize Chrome WebDriver with user profile to avoid login."""
    options = Options()
    options.add_argument(f"--user-data-dir={USER_DATA_DIR}")
    options.add_argument(f"--profile-directory={PROFILE_DIR}")
    options.add_argument("--start-maximized")
    options.add_argument("--disable-notifications")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/135.0.7049.85 Safari/537.36"
    )

    try:
        service = ChromeDriverManager().install() if USE_WEBDRIVER_MANAGER else 'chromedriver.exe'
        driver = webdriver.Chrome(service=webdriver.chrome.service.Service(service), options=options)
        driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
            "source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
        })
        logger.info(f"WebDriver initialized with profile {USER_DATA_DIR}/{PROFILE_DIR}")
        return driver
    except Exception as e:
        logger.error(f"Failed to initialize WebDriver: {e}")
        raise FormAutomationError("WebDriver setup failed")

def download_google_drive_file(url: str, output_path: str) -> Optional[str]:
    """Download a file from Google Drive and ensure it exists."""
    try:
        file_id = None
        if "id=" in url:
            file_id = url.split("id=")[1].split("&")[0]
        elif "/file/d/" in url:
            file_id = url.split("/file/d/")[1].split("/")[0]

        if not file_id:
            logger.warning(f"Invalid Google Drive URL: {url}")
            return None

        download_url = f"https://drive.google.com/uc?export=download&id={file_id}"
        response = requests.get(download_url, stream=True, timeout=30)

        if response.status_code != 200:
            logger.warning(f"Failed to download file from {url}: Status {response.status_code}")
            return None

        content_length = response.headers.get("Content-Length")
        if content_length and int(content_length) > MAX_FILE_SIZE:
            logger.warning(f"File exceeds size limit: {content_length} bytes")
            return None

        with open(output_path, "wb") as f:
            for chunk in response.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)

        # Verify file exists and is non-empty
        if not os.path.exists(output_path) or os.path.getsize(output_path) == 0:
            logger.warning(f"File not fully downloaded: {output_path}")
            return None

        logger.info(f"Downloaded file to {output_path}")
        return output_path
    except Exception as e:
        logger.warning(f"Error downloading file from {url}: {e}")
        return None

def read_excel_data(filepath: str) -> Tuple[List[str], List[List[str]]]:
    """Read headers and data from Excel file."""
    try:
        wb = openpyxl.load_workbook(filepath, read_only=True)
        sheet = wb.active
        headers = [
            cell.value.strftime("%Y-%m-%d") if isinstance(cell.value, datetime)
            else str(cell.value).strip()
            for cell in sheet[1] if cell.value is not None
        ]
        data = [
            ["" if val is None else val for val in row]
            for row in sheet.iter_rows(min_row=2, values_only=True) if any(row)
        ]
        logger.info(f"Read {len(data)} rows from {filepath}")
        return headers, data
    except Exception as e:
        logger.error(f"Error reading Excel file {filepath}: {e}")
        raise FormAutomationError("Excel file reading failed")

def get_form_headers(driver: webdriver.Chrome) -> List[str]:
    """Retrieve headers from the Google Form."""
    try:
        driver.get(GOOGLE_FORM_URL)
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, "//span[@class='M7eMe']")))
        headers = [
            elem.text.strip()
            for elem in driver.find_elements(By.XPATH, "//span[@class='M7eMe']")
            if elem.text.strip()
        ]
        if not headers:
            raise ValueError("No form headers found")
        logger.info(f"Retrieved {len(headers)} form headers")
        return headers
    except Exception as e:
        logger.error(f"Error fetching form headers: {e}")
        raise FormAutomationError("Form headers retrieval failed")

def match_headers(excel_headers: List[str], form_headers: List[str]) -> Tuple[List[str], List[str], Dict[str, str]]:
    """Match Excel headers to form headers using fuzzy matching."""
    matched, unmatched, mapping = [], [], {}
    for excel_header in excel_headers:
        best_match, best_score = None, 0
        for form_header in form_headers:
            score = fuzz.token_sort_ratio(excel_header, form_header)
            if score > best_score and score >= SIMILARITY_THRESHOLD:
                best_match, best_score = form_header, score
        if best_match:
            matched.append(excel_header)
            mapping[excel_header] = best_match
        else:
            unmatched.append(excel_header)
    logger.info(f"Matched {len(matched)} headers, unmatched: {unmatched}")
    return matched, unmatched, mapping

def scroll_into_view(driver: webdriver.Chrome, element) -> None:
    """Scroll an element into view smoothly."""
    try:
        driver.execute_script(
            "arguments[0].scrollIntoView({block: 'center', behavior: 'smooth'});",
            element
        )
        time.sleep(0.2)  # Brief pause for scroll to settle
    except Exception as e:
        logger.warning(f"Error scrolling element into view: {e}")

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

def parse_date(value: any) -> Optional[str]:
    """Parse date into MM/DD/YYYY format."""
    if not value or isinstance(value, str) and "drive.google.com" in value:
        return None
    if isinstance(value, datetime):
        return value.strftime("%m/%d/%Y")
    if isinstance(value, str):
        for fmt in ["%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y"]:
            try:
                return datetime.strptime(value, fmt).strftime("%m/%d/%Y")
            except ValueError:
                continue
    logger.warning(f"Unsupported date format: {value}")
    return None

def fill_date_field(driver: webdriver.Chrome, form_header: str, value: any) -> bool:
    """Fill a date field."""
    try:
        date_value = parse_date(value)
        if not date_value:
            return False

        month, day, year = date_value.split("/")
        xpath = f"//span[@class='M7eMe' and normalize-space(.)='{form_header}']/ancestor::div[@role='listitem']"
        container = driver.find_element(By.XPATH, xpath)

        date_inputs = container.find_elements(By.XPATH, ".//input[@type='date']")
        if date_inputs:
            scroll_into_view(driver, date_inputs[0])
            date_inputs[0].send_keys(f"{month}{day}{year}")
            logger.info(f"Filled date field '{form_header}' with {date_value}")
            return True

        month_input = container.find_element(By.XPATH, ".//input[@aria-label='Month']")
        day_input = container.find_element(By.XPATH, ".//input[@aria-label='Day of the month']")
        year_input = container.find_element(By.XPATH, ".//input[@aria-label='Year']")

        for input_elem, val in [(month_input, month), (day_input, day), (year_input, year)]:
            scroll_into_view(driver, input_elem)
            input_elem.clear()
            input_elem.send_keys(val)

        logger.info(f"Filled date field '{form_header}' with {date_value}")
        return True
    except NoSuchElementException:
        logger.warning(f"Date field '{form_header}' not found")
        return False
    except Exception as e:
        logger.warning(f"Error filling date field '{form_header}': {e}")
        return False

@retry(stop_max_attempt_number=3, wait_fixed=5000)
def upload_file(driver: webdriver.Chrome, file_path: str) -> bool:
    """Upload a file through Google's file picker iframe and check completion."""
    try:
        logger.info(f"Attempting to upload file: {file_path}")
        WebDriverWait(driver, 30).until(
            EC.frame_to_be_available_and_switch_to_it((By.XPATH, "//iframe[contains(@src, 'docs.google.com/picker')]"))
        )
        file_input = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, "//input[@type='file']"))
        )
        file_input.send_keys(file_path)
        logger.info(f"File path sent to input: {file_path}")
        insert_btn = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, "//div[@role='dialog']//span[text()='Insert']/ancestor::button"))
        )
        driver.execute_script("arguments[0].click();", insert_btn)
        driver.switch_to.default_content()
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, "//div[@role='list' and @aria-label='Selected files']//div"))
        )
        logger.info(f"File '{file_path}' uploaded successfully")
        return True
    except TimeoutException:
        logger.warning(f"Timeout waiting for file upload to complete: {file_path}")
        driver.switch_to.default_content()
        return False
    except Exception as e:
        logger.warning(f"File upload failed: {e}")
        driver.switch_to.default_content()
        return False

def fill_file_upload(driver: webdriver.Chrome, form_header: str, file_url: str) -> bool:
    """Handle file upload for a form field and verify completion."""
    try:
        if not file_url or "drive.google.com" not in file_url:
            logger.warning(f"Invalid file URL for '{form_header}': {file_url}")
            return False

        temp_file = os.path.join(TEMP_DIR, f"image_{int(time.time())}.jpg")
        downloaded_file = download_google_drive_file(file_url, temp_file)
        if not downloaded_file:
            logger.warning(f"Failed to download file for '{form_header}'")
            return False

        # Ensure file is ready (no wait time, just a check)
        if not os.path.exists(downloaded_file) or os.path.getsize(downloaded_file) == 0:
            logger.warning(f"File not ready for upload: {downloaded_file}")
            return False
        logger.info(f"File {downloaded_file} is ready for upload")

        # Click "Add File" button with retry for interception
        button_xpath = f"//span[@class='M7eMe' and normalize-space(.)='{form_header}']/ancestor::div[@role='listitem']//div[@role='button' and contains(@aria-label, 'Add File')]"
        for _ in range(3):
            try:
                upload_button = WebDriverWait(driver, 30).until(
                    EC.element_to_be_clickable((By.XPATH, button_xpath))
                )
                scroll_into_view(driver, upload_button)
                driver.execute_script("arguments[0].click();", upload_button)
                logger.info(f"Clicked 'Add File' button for '{form_header}'")
                break
            except ElementClickInterceptedException as e:
                logger.warning(f"Add File button click intercepted: {e}")
                return False

        # Upload the file and check completion
        return upload_file(driver, downloaded_file)

    except TimeoutException:
        logger.warning(f"Timeout during file upload for '{form_header}'")
        return False
    except Exception as e:
        logger.warning(f"Error uploading file for '{form_header}': {e}")
        return False
    finally:
        if os.path.exists(temp_file):
            try:
                os.remove(temp_file)
                logger.debug(f"Cleaned up temp file: {temp_file}")
            except Exception as e:
                logger.warning(f"Failed to delete temp file {temp_file}: {e}")
def fill_google_form(driver, row, headers, header_mapping, is_first_submission: bool = False):
    try:
        driver.get(GOOGLE_FORM_URL)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//form")))
        
        # Clear inputs only for the first submission
        if is_first_submission:
            logger.info("Clearing inputs for the first submission")
            clear_inputs(driver)

        for excel_header, value in zip(headers, row):
            if excel_header not in header_mapping or value is None:
                continue

            form_header = header_mapping[excel_header]
            value_str = str(value).strip()
            logging.info(f"Processing field '{form_header}' with value '{value_str}'")

            # Wait for the field header to be present
            try:
                # Use partial match to handle long headers and special characters
                header_xpath = (
                    f"//span[@class='M7eMe' and contains(normalize-space(.), '{form_header[:50]}')]"
                )
                WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, header_xpath))
                )
                logger.debug(f"Found header: {form_header[:50]}...")
            except TimeoutException:
                logging.warning(f"Field header '{form_header}' not found in form")
                continue

            # Check if the field is a date field
            date_input_xpath = (
                f"//span[@class='M7eMe' and contains(normalize-space(.), '{form_header[:50]}')]/ancestor::div[@role='listitem']//"
                f"input[@type='date' or @aria-label='Month' or @aria-label='Day of the month' or @aria-label='Year']"
            )
            date_inputs = driver.find_elements(By.XPATH, date_input_xpath)
            if date_inputs:
                if fill_date_field(driver, form_header, value):
                    continue
            # Check if the field is a file upload by looking for the "Add File" button
            file_button_xpath = (
                f"//span[@class='M7eMe' and contains(normalize-space(.), '{form_header[:50]}')]/ancestor::div[@role='listitem']//"
                f"div[@role='button' and contains(@aria-label, 'Add File')]"
            )
            file_buttons = driver.find_elements(By.XPATH, file_button_xpath)
            if file_buttons and value_str.startswith("http"):
                try:
                    scroll_into_view(driver, file_buttons[0])
                    if fill_file_upload(driver, form_header, value_str):
                        logging.info(f"✅ Uploaded file for '{form_header}'")
                        continue
                except Exception as e:
                    logging.warning(f"⚠️ Error uploading file for '{form_header}': {e}")
                    continue

            try:
                # Try text input
                input_xpath = (
                    f"//span[@class='M7eMe' and contains(normalize-space(.), '{form_header[:50]}')]/ancestor::div[@role='listitem']//"
                    f"input[@type='text' or @type='number']"
                )
                inputs = driver.find_elements(By.XPATH, input_xpath)
                if inputs:
                    # Wait for the input to be interactable
                    WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH, input_xpath))
                    )
                    scroll_into_view(driver, inputs[0])
                    inputs[0].clear()
                    inputs[0].send_keys(value_str)
                    logging.info(f"Filled text field '{form_header}' with '{value_str}'")
                    continue

                # Try textarea
                textarea_xpath = (
                    f"//span[@class='M7eMe' and contains(normalize-space(.), '{form_header[:50]}')]/ancestor::div[@role='listitem']//textarea"
                )
                textareas = driver.find_elements(By.XPATH, textarea_xpath)
                if textareas:
                    WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH, textarea_xpath))
                    )
                    scroll_into_view(driver, textareas[0])
                    textareas[0].clear()
                    textareas[0].send_keys(value_str)
                    logging.info(f"Filled textarea field '{form_header}' with '{value_str}'")
                    continue

                # Try dropdown
                dropdown_xpath = (
                    f"//span[@class='M7eMe' and contains(normalize-space(.), '{form_header[:50]}')]/ancestor::div[@role='listitem']//div[@role='listbox']"
                )
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

def main():
    """Main function to orchestrate form automation."""
    driver = None
    try:
        kill_chrome_processes()
        driver = setup_driver()
        form_headers = get_form_headers(driver)
        excel_headers, rows = read_excel_data(EXCEL_FILE)
        matched_headers, unmatched_headers, header_mapping = match_headers(excel_headers, form_headers)

        if unmatched_headers:
            logger.warning(f"Unmatched Excel headers: {unmatched_headers}")
        logger.info(f"Header mapping: {header_mapping}")

        for idx, row in enumerate(rows, start=2):
            logger.info(f"Processing row {idx}")
            # Clear inputs only for the first row
            fill_google_form(driver, row, excel_headers, header_mapping, is_first_submission=(idx == 2))
            logger.info(f"Row {idx} processed successfully")

    except FormAutomationError as e:
        logger.error(f"Automation failed: {e}")
    except Exception as e:
        logger.error(f"Unexpected error: {e}")
    finally:
        if driver:
            driver.quit()
            logger.info("WebDriver closed")
        if os.path.exists(TEMP_DIR):
            shutil.rmtree(TEMP_DIR, ignore_errors=True)
            logger.info("Temporary directory cleaned up")

if __name__ == "__main__":
    main()