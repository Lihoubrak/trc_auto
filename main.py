import time
import logging
import openpyxl
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from fuzzywuzzy import fuzz
import psutil
from pathlib import Path
import tempfile
import requests
import re
import os
import logging
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from retrying import retry
# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[logging.StreamHandler()]
)

# Constants
GOOGLE_FORM_URL = "https://docs.google.com/forms/d/e/1FAIpQLSeqWvnn4KIru5BYd6aNVTCvaej6KvPWdbK0tN3piOgU8u8ftg/viewform"
EXCEL_FILE = "test copy.xlsx"
CHROMEDRIVER_PATH = r"D:\trc_auto\chromedriver.exe"
SIMILARITY_THRESHOLD = 80
USER_DATA_DIR = r"C:\Users\user\AppData\Local\Google\Chrome\User Data"
PROFILE_DIR = "Profile 5"

# Check for webdriver_manager
try:
    from webdriver_manager.chrome import ChromeDriverManager
    USE_WEBDRIVER_MANAGER = True
except ImportError:
    USE_WEBDRIVER_MANAGER = False
    logging.warning("webdriver_manager not installed. Using static chromedriver path.")

def terminate_chrome_processes():
    """Terminate all running Chrome processes."""
    try:
        for proc in psutil.process_iter(['pid', 'name']):
            if proc.info['name'].lower() == "chrome.exe":
                proc.kill()
                logging.info(f"Terminated Chrome process PID: {proc.pid}")
    except Exception as e:
        logging.error(f"Error terminating Chrome processes: {e}")

def initialize_driver():
    """Initialize and configure Chrome WebDriver."""
    options = webdriver.ChromeOptions()
    options.add_argument(f"--user-data-dir={USER_DATA_DIR}")
    options.add_argument(f"--profile-directory={PROFILE_DIR}")
    options.add_argument("--start-maximized")
    options.add_argument("--disable-notifications")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/135.0.0.0 Safari/537.36"
    )

    try:
        service = (
            webdriver.chrome.service.Service(ChromeDriverManager().install())
            if USE_WEBDRIVER_MANAGER
            else webdriver.chrome.service.Service(CHROMEDRIVER_PATH)
        )
        driver = webdriver.Chrome(service=service, options=options)
        driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
            "source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
        })
        logging.info("WebDriver initialized successfully")
        return driver
    except Exception as e:
        logging.error(f"Failed to initialize WebDriver: {e}")
        raise

def read_excel_data(filepath):
    """Read headers and data from an Excel file."""
    try:
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

        data = [
            ["" if val is None else val for val in row]
            for row in sheet.iter_rows(min_row=2, values_only=True) if any(row)
        ]
        logging.info(f"Read {len(data)} rows from {filepath.name}")
        return headers, data
    except Exception as e:
        logging.error(f"Error reading Excel file: {e}")
        raise

def get_form_headers(driver):
    """Retrieve headers from the Google Form."""
    try:
        driver.get(GOOGLE_FORM_URL)
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//span[@class='M7eMe']"))
        )
        headers = [
            elem.text.strip()
            for elem in driver.find_elements(By.XPATH, "//span[@class='M7eMe']")
            if elem.text.strip()
        ]
        if not headers:
            raise ValueError("No form headers found")
        logging.info(f"Retrieved {len(headers)} form headers")
        return headers
    except Exception as e:
        logging.error(f"Error fetching form headers: {e}")
        raise

def match_headers(excel_headers, form_headers):
    """Match Excel headers to form headers using fuzzy matching."""
    mapping = {}
    unmatched = []
    for excel_header in excel_headers:
        best_match, best_score = None, 0
        for form_header in form_headers:
            score = fuzz.token_sort_ratio(excel_header, form_header)
            if score > best_score and score >= SIMILARITY_THRESHOLD:
                best_match, best_score = form_header, score
        if best_match:
            mapping[excel_header] = best_match
            logging.info(f"Matched '{excel_header}' to '{best_match}' (score: {best_score})")
        else:
            unmatched.append(excel_header)
            logging.warning(f"No match for Excel header: '{excel_header}'")

    if unmatched:
        logging.warning(f"Unmatched headers: {unmatched}")
    return mapping, unmatched

def scroll_into_view(driver, element):
    """Scroll an element into view."""
    driver.execute_script("arguments[0].scrollIntoView({block: 'center', behavior: 'smooth'});", element)
    time.sleep(0.2)

def parse_date(value):
    """Parse various date formats into MM/DD/YYYY."""
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
    logging.warning(f"Unsupported date format: {value}")
    return None

def fill_date_field(driver, form_header, date_value):
    """Fill a date field in the Google Form."""
    try:
        date_value = parse_date(date_value)
        if not date_value:
            return False

        month, day, year = date_value.split("/")
        date_input_xpath = (
            f"//span[@class='M7eMe' and normalize-space(.)='{form_header}']"
            f"/ancestor::div[@role='listitem']//input[@type='date']"
        )
        date_inputs = driver.find_elements(By.XPATH, date_input_xpath)
        if date_inputs:
            scroll_into_view(driver, date_inputs[0])
            date_inputs[0].send_keys(f"{month}{day}{year}")
            logging.info(f"Filled date field '{form_header}' with '{month}{day}{year}'")
            return True

        date_container_xpath = (
            f"//span[@class='M7eMe' and normalize-space(.)='{form_header}']"
            f"/ancestor::div[@role='listitem']"
        )
        date_container = driver.find_element(By.XPATH, date_container_xpath)
        inputs = {
            "month": date_container.find_element(By.XPATH, ".//input[@aria-label='Month']"),
            "day": date_container.find_element(By.XPATH, ".//input[@aria-label='Day of the month']"),
            "year": date_container.find_element(By.XPATH, ".//input[@aria-label='Year']")
        }

        for field, value in zip(inputs, [month, day, year]):
            scroll_into_view(driver, inputs[field])
            inputs[field].clear()
            inputs[field].send_keys(value)
        logging.info(f"Filled date field '{form_header}' with '{month}/{day}/{year}'")
        return True

    except NoSuchElementException:
        logging.warning(f"Date field '{form_header}' not found")
        return False
    except Exception as e:
        logging.error(f"Error filling date field '{form_header}': {e}")
        return False
import re
import logging
import requests
import tempfile
import os
from urllib.parse import urlparse

def download_google_drive_image(google_drive_link, temp_dir=r"C:\Users\user\Desktop\trc_auto"):
    try:
        # Validate Google Drive link format
        valid_patterns = [
            r'https?://drive\.google\.com/file/d/([-\w]{25,})/',
            r'https?://drive\.google\.com/uc\?id=([-\w]{25,})',
            r'https?://drive\.google\.com/open\?id=([-\w]{25,})'
        ]
        file_id = None
        for pattern in valid_patterns:
            match = re.search(pattern, google_drive_link)
            if match:
                file_id = match.group(1)
                break
        
        if not file_id:
            logging.error(f"Invalid or unsupported Google Drive link: {google_drive_link}")
            return None

        # Validate file ID format (strict alphanumeric check, typical length 25-40)
        if not re.match(r'^[a-zA-Z0-9_-]{25,40}$', file_id):
            logging.error(f"Suspicious file ID format: {file_id}")
            return None

        # Construct download URL
        download_url = f"https://drive.google.com/uc?export=download&id={file_id}"

        # Verify the domain before downloading
        parsed_url = urlparse(download_url)
        if parsed_url.netloc != 'drive.google.com':
            logging.error(f"Invalid domain in download URL: {parsed_url.netloc}")
            return None

        # Download the file with additional headers and timeout
        headers = {'User-Agent': 'Mozilla/5.0 (compatible; SafeDownloader/1.0)'}
        response = requests.get(download_url, stream=True, headers=headers, timeout=10)
        
        # Check for successful response and content type
        if response.status_code != 200:
            logging.error(f"Failed to download file from {download_url}: Status {response.status_code}")
            return None
        
        content_type = response.headers.get('Content-Type', '')
        if not content_type.startswith('image/'):
            logging.error(f"Unexpected content type: {content_type}")
            return None

        # Ensure the specified directory exists
        os.makedirs(temp_dir, exist_ok=True)

        # Create a temporary file in the specified directory
        temp_file = tempfile.NamedTemporaryFile(
            delete=False, 
            suffix='.jpg',
            dir=temp_dir
        )
        
        # Write content with size limit (e.g., 10MB max)
        max_size = 10 * 1024 * 1024  # 10MB
        downloaded_size = 0
        for chunk in response.iter_content(chunk_size=8192):
            if chunk:
                downloaded_size += len(chunk)
                if downloaded_size > max_size:
                    temp_file.close()
                    os.unlink(temp_file.name)
                    logging.error(f"File exceeds maximum size limit: {max_size} bytes")
                    return None
                temp_file.write(chunk)
        temp_file.close()
        
        logging.info(f"Downloaded image to {temp_file.name}")
        return temp_file.name

    except requests.exceptions.Timeout:
        logging.error(f"Download timed out for link: {google_drive_link}")
        return None
    except requests.exceptions.RequestException as e:
        logging.error(f"Network error downloading Google Drive image: {e}")
        return None
    except Exception as e:
        logging.error(f"Error downloading Google Drive image: {e}")
        return None
def fill_form_field(driver, form_header, value, field_type="text"):
    """Fill a specific form field based on its type."""
    try:
        value_str = str(value).strip()
        if not value_str:
            return False

        if field_type == "date":
            return fill_date_field(driver, form_header, value)

        xpath_map = {
            "text": (
                f"//span[@class='M7eMe' and contains(normalize-space(.), '{form_header[:50]}')]"
                f"/ancestor::div[@role='listitem']//input[@type='text' or @type='number']"
            ),
            "textarea": (
                f"//span[@class='M7eMe' and contains(normalize-space(.), '{form_header[:50]}')]"
                f"/ancestor::div[@role='listitem']//textarea"
            ),
            "dropdown": (
                f"//span[@class='M7eMe' and contains(normalize-space(.), '{form_header[:50]}')]"
                f"/ancestor::div[@role='listitem']//div[@role='listbox']"
            ),
            "checkbox": (
                f"//div[@role='listitem']//span[normalize-space(.)='{value_str}']/ancestor::label"
            ),
        }

        xpath = xpath_map.get(field_type)
        if not xpath:
            logging.warning(f"Unsupported field type: {field_type}")
            return False

        elements = driver.find_elements(By.XPATH, xpath)
        if not elements:
            logging.warning(f"No elements found for form header '{form_header}' and type '{field_type}'")
            return False

        scroll_into_view(driver, elements[0])

        if field_type in ["text", "textarea"]:
            elements[0].clear()
            elements[0].send_keys(value_str)
            logging.info(f"Filled {field_type} field '{form_header}' with '{value_str}'")
        elif field_type == "dropdown":
            elements[0].click()
            time.sleep(0.5)
            options = driver.find_elements(By.XPATH, "//div[@role='option']")
            for opt in options:
                if value_str.lower() in opt.text.lower():
                    scroll_into_view(driver, opt)
                    opt.click()
                    logging.info(f"Selected dropdown option '{opt.text}' for '{form_header}'")
                    return True
            logging.warning(f"No matching dropdown option for '{form_header}' value '{value_str}'")
            return False
        elif field_type == "checkbox":
            elements[0].click()
            logging.info(f"Checked checkbox '{form_header}' with value '{value_str}'")

        return True

    except Exception as e:
        logging.error(f"Error filling form field '{form_header}': {e}")
        return False



def fill_google_form(driver, row, headers, header_mapping):
    """Fill and submit a Google Form for one row of data."""
    temp_files = []
    fields_filled = True

    try:
        # Step 1: Open the form
        driver.get(GOOGLE_FORM_URL)
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, "//form")))
        logging.info("Google Form loaded successfully")

        # Step 2: Optional Email Checkbox
        try:
            checkbox_xpath = '//div[.//span[text()="Email"]]/following::div[@role="checkbox"][1]'
            checkbox = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, checkbox_xpath)))
            scroll_into_view(driver, checkbox)
            checkbox.click()
            logging.info("Checked 'Email' collection checkbox")
        except TimeoutException:
            logging.info("No email checkbox found — skipping")

        # Step 3: Fill form fields
        for excel_header, value in zip(headers, row):
            if excel_header not in header_mapping or not value:
                logging.info(f"Skipping empty or unmapped field: {excel_header}")
                continue

            form_header = header_mapping[excel_header]
            logging.info(f"Processing field: {form_header}")

            # Image upload handling
            is_image_field = (
                "រូបភាពនៃស្ថានភាពការខូចខាតនៃខ្សែកាប្លិ៍ដោយមាន lat/long <10MB (Picture of Damage Cable with lat/long <10MB):"
                in form_header or "រូបភាពនៃគំនូសនៅលើ Google Map ដែលមានចំណុចចាប់ផ្តើមនិងបញ្ចប់ (Picture of drawing in google map with start and end lat/long):" in form_header
            )
            if is_image_field:
                if isinstance(value, str) and "drive.google.com" in value:
                    temp_file_path = download_google_drive_image(value)
                    if temp_file_path:
                        temp_files.append(temp_file_path)
                        uploaded = False
                        file_name = os.path.basename(temp_file_path)

                        @retry(stop_max_attempt_number=3, wait_fixed=2000)
                        def attempt_upload():
                            nonlocal uploaded
                            try:
                                # Locate the "Add File" button for the specific header
                                upload_btn_xpath = (
                                    f"//span[@class='M7eMe' and contains(normalize-space(.), '{form_header}')]/ancestor::div[@role='listitem']//"
                                    f"div[@role='button' and contains(@aria-label, 'Add File')]"
                                )
                                upload_button = WebDriverWait(driver, 15).until(
                                    EC.element_to_be_clickable((By.XPATH, upload_btn_xpath))
                                )
                                scroll_into_view(driver, upload_button)
                                logging.info(f"Located 'Add File' button for '{form_header}'")

                                # Click the "Add File" button to open a new iframe
                                driver.execute_script("arguments[0].click();", upload_button)
                                logging.info("Clicked 'Add File' button")

                                # Wait for the parent div to become visible and contain the iframe
                                picker_dialog_xpath = "//div[contains(@class, 'fFW7wc XKSfm-Sx9Kwc picker-dialog') and not(contains(@style, 'display: none'))]"
                                WebDriverWait(driver, 20).until(
                                    EC.presence_of_element_located((By.XPATH, picker_dialog_xpath))
                                )
                                logging.info("Picker dialog div is visible")

                                # Locate the iframe within the picker dialog
                                iframe_xpath = f"{picker_dialog_xpath}//iframe[contains(@src, 'docs.google.com/picker')]"
                                WebDriverWait(driver, 20).until(
                                    EC.frame_to_be_available_and_switch_to_it((By.XPATH, iframe_xpath))
                                )
                                logging.info("Switched to new file picker iframe")

                                # Wait for the file input in the iframe
                                file_input = WebDriverWait(driver, 15).until(
                                    EC.presence_of_element_located((By.XPATH, "//input[@type='file']"))
                                )
                                file_input.send_keys(temp_file_path)
                                logging.info(f"Sent file path to input: {temp_file_path}")

                                # Switch back to default content before checking for dialog closure
                                driver.switch_to.default_content()

                                # Wait for the picker dialog to disappear (iframe closed)
                                WebDriverWait(driver, 30).until(
                                    EC.invisibility_of_element_located((By.XPATH, picker_dialog_xpath))
                                )
                                logging.info("Picker dialog and iframe closed")

                                # Verify the file appears in the form's file list
                                file_list_xpath = (
                                    f"//span[@class='M7eMe' and contains(normalize-space(.), '{form_header}')]/ancestor::div[@role='listitem']"
                                    f"//div[@role='listitem']//div[contains(@class, 'ZXoVYe ybj8pf') and contains(text(), '{file_name}')]"
                                )
                                WebDriverWait(driver, 20).until(
                                    EC.presence_of_element_located((By.XPATH, file_list_xpath))
                                )
                                logging.info(f"File '{file_name}' confirmed in form's file list")
                                return True

                            except TimeoutException as te:
                                logging.error(f"Timeout during file upload attempt for '{form_header}': {te}")
                                driver.switch_to.default_content()
                                raise
                            except Exception as e:
                                logging.error(f"Error during file upload attempt for '{form_header}': {e}")
                                driver.switch_to.default_content()
                                raise

                        try:
                            uploaded = attempt_upload()
                        except Exception as e:
                            logging.error(f"Failed to upload file '{file_name}' for '{form_header}' after retries: {e}")
                            uploaded = False

                        if not uploaded:
                            logging.error(f"Failed to upload file '{file_name}' for '{form_header}'")
                            fields_filled = False
                        else:
                            logging.info(f"Successfully uploaded and verified file '{file_name}' for '{form_header}'")
                    else:
                        logging.warning(f"Failed to download image from Google Drive for '{form_header}': {value}")
                        fields_filled = False
                else:
                    logging.warning(f"Invalid Google Drive URL for image field '{form_header}': {value}")
                    fields_filled = False
                continue

            # Normal fields: text, dropdown, etc.
            filled = False
            for field_type in ["date", "text", "textarea", "dropdown", "checkbox"]:
                if fill_form_field(driver, form_header, value, field_type):
                    filled = True
                    break
            if not filled:
                logging.warning(f"No matching input found for '{form_header}'")
                fields_filled = False

        # Step 4: Submit form
        try:
            submit_btn = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//span[text()='Submit']/ancestor::div[@role='button']"))
            )
            scroll_into_view(driver, submit_btn)
            driver.execute_script("arguments[0].click();", submit_btn)
            WebDriverWait(driver, 15).until(EC.url_contains("formResponse"))
            logging.info("Form submitted successfully")
        except Exception as e:
            logging.error(f"Form submission failed: {e}")
            fields_filled = False

    except Exception as e:
        logging.error(f"Error while filling the form: {e}")
        fields_filled = False

    finally:
        for f in temp_files:
            try:
                os.remove(f)
                logging.info(f"Deleted temp file: {f}")
            except Exception as e:
                logging.warning(f"Failed to delete temp file: {f} - {e}")

    return fields_filled


def main():
    """Main function to orchestrate the automation process."""
    driver = None
    try:
        terminate_chrome_processes()
        driver = initialize_driver()
        form_headers = get_form_headers(driver)
        excel_headers, rows = read_excel_data(EXCEL_FILE)
        header_mapping, unmatched_headers = match_headers(excel_headers, form_headers)
        for idx, row in enumerate(rows, start=2):
            logging.info(f"Processing row {idx}")
            success = fill_google_form(driver, row, excel_headers, header_mapping)
            logging.info(f"Row {idx} processed {'successfully' if success else 'with errors'}")
            time.sleep(1)

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