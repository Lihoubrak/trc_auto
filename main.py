import time
import os
import logging
import openpyxl
import tempfile
import shutil
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from fuzzywuzzy import fuzz
import psutil
import gdown

# Configure logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

# Constants
GOOGLE_FORM_URL = "https://docs.google.com/forms/d/e/1FAIpQLSeqWvnn4KIru5BYd6aNVTCvaej6KvPWdbK0tN3piOgU8u8ftg/viewform"
EXCEL_FILE = "MAINTENANCE CABLE REQUEST TO VTC.xlsx"  # Update to your Excel file path
SIMILARITY_THRESHOLD = 80
TEMP_DIR = tempfile.mkdtemp()
MAX_FILE_SIZE = 10 * 1024 * 1024  # 10MB limit
CHROMEDRIVER_PATH = r"D:\trc_auto\chromedriver.exe"  # Adjust if needed

# Optional: Check for webdriver_manager
try:
    from webdriver_manager.chrome import ChromeDriverManager
    USE_WEBDRIVER_MANAGER = True
except ImportError:
    USE_WEBDRIVER_MANAGER = False

def kill_chrome():
    """Terminate all running Chrome processes."""
    for proc in psutil.process_iter(attrs=['pid', 'name']):
        if proc.info['name'].lower() == "chrome.exe":
            try:
                proc.kill()
                logging.info(f"Killed Chrome process: {proc.pid}")
            except Exception as e:
                logging.warning(f"Could not kill Chrome process {proc.pid}: {e}")

def setup_driver():
    """Initialize and configure Chrome WebDriver."""
    from selenium.webdriver.chrome.options import Options
    options = Options()
    user_data_dir = r"C:\Users\KHC\AppData\Local\Google\Chrome\User Data"
    profile_dir = "Profile 1"
    options.add_argument(f"--user-data-dir={user_data_dir}")
    options.add_argument(f"--profile-directory={profile_dir}")
    options.add_argument("--start-maximized")
    options.add_argument("--disable-notifications")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/135.0.0.0 Safari/537.36"
    )

    try:
        if USE_WEBDRIVER_MANAGER:
            driver = webdriver.Chrome(service=webdriver.chrome.service.Service(ChromeDriverManager().install()), options=options)
        else:
            driver = webdriver.Chrome(service=webdriver.chrome.service.Service(CHROMEDRIVER_PATH), options=options)
    except Exception as e:
        logging.error(f"Failed to initialize WebDriver: {e}")
        raise

    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
    })
    return driver

def download_google_drive_file(url, output_path):
    """Download an image from a Google Drive URL using gdown."""
    try:
        if not url or "drive.google.com" not in url:
            logging.warning(f"Invalid or missing Google Drive URL: '{url}'")
            return None

        gdown.download(url=url, output=output_path, quiet=False, fuzzy=True)
        
        if not os.path.exists(output_path) or os.path.getsize(output_path) == 0:
            logging.warning(f"Failed to download file to {output_path}")
            return None

        file_size = os.path.getsize(output_path)
        if file_size > MAX_FILE_SIZE:
            logging.warning(f"File at {url} exceeds {MAX_FILE_SIZE/1024/1024}MB limit")
            os.remove(output_path)
            return None

        logging.info(f"Downloaded file to {output_path}")
        return output_path
    except Exception as e:
        logging.warning(f"Error downloading file from {url}: {e}")
        return None

def read_excel_data(filepath):
    """Read headers and data from an Excel file."""
    try:
        wb = openpyxl.load_workbook(filepath)
        sheet = wb.active
        headers = []
        for cell in sheet[1]:
            if cell.value is None:
                continue
            if isinstance(cell.value, datetime):
                headers.append(cell.value.strftime("%Y-%m-%d"))
            else:
                headers.append(str(cell.value).strip())

        data = []
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if any(row):
                cleaned_row = [val if val is not None else "" for val in row]
                data.append(cleaned_row)
        logging.info(f"Read {len(data)} rows from Excel file")
        return headers, data
    except Exception as e:
        logging.error(f"Error reading Excel file: {e}")
        raise

def get_form_headers_and_file_fields(driver):
    """Retrieve headers and identify file upload fields from the Google Form."""
    try:
        driver.get(GOOGLE_FORM_URL)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//span[@class='M7eMe']")))
        
        # Get all form headers
        header_elements = driver.find_elements(By.XPATH, "//span[@class='M7eMe']")
        headers = [elem.text.strip() for elem in header_elements if elem.text.strip()]
        if not headers:
            raise ValueError("No form headers found")
        
        # Identify file upload fields by checking for "Add File" buttons
        file_upload_headers = []
        for header in headers:
            button_xpath = (
                f"//span[@class='M7eMe' and contains(normalize-space(.), '{header[:50]}')]/ancestor::div[@role='listitem']"
                f"//div[@role='button' and contains(@aria-label, 'Add File')]"
            )
            if driver.find_elements(By.XPATH, button_xpath):
                file_upload_headers.append(header)
        
        logging.info(f"Form headers: {headers}")
        logging.info(f"Detected file upload fields: {file_upload_headers}")
        return headers, file_upload_headers
    except Exception as e:
        logging.error(f"Error fetching form headers: {e}")
        raise

def match_headers(excel_headers, form_headers):
    """Match Excel headers to form headers using fuzzy matching."""
    matched_headers, mapping, unmatched = [], {}, []
    for excel_header in excel_headers:
        best_match, best_score = None, 0
        for form_header in form_headers:
            score = fuzz.token_sort_ratio(excel_header, form_header)
            if score > best_score and score >= SIMILARITY_THRESHOLD:
                best_match, best_score = form_header, score
        if best_match:
            matched_headers.append(excel_header)
            mapping[excel_header] = best_match
            logging.info(f"Fuzzy matched: '{excel_header}' → '{best_match}' (score: {best_score})")
        else:
            unmatched.append(excel_header)
            logging.warning(f"No match for Excel header: '{excel_header}'")
    
    logging.info(f"Header mappings: {mapping}")
    if unmatched:
        logging.warning(f"Unmatched Excel headers: {unmatched}")
    return matched_headers, unmatched, mapping

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
                parsed_date = datetime.strptime(value, fmt)
                return parsed_date.strftime("%m/%d/%Y")
            except ValueError:
                continue
        logging.warning(f"Unsupported date format: {value}")
    return None

def fill_date_field(driver, form_header, date_value):
    """Fill a date field in the form."""
    try:
        date_value = parse_date(date_value)
        if not date_value:
            return False

        month, day, year = date_value.split("/")

        date_input_xpath = (
            f"//span[@class='M7eMe' and normalize-space(.)='{form_header}']/ancestor::div[@role='listitem']//input[@type='date']"
        )
        date_inputs = driver.find_elements(By.XPATH, date_input_xpath)
        if date_inputs:
            scroll_into_view(driver, date_inputs[0])
            date_for_input = f"{month}{day}{year}"
            date_inputs[0].send_keys(date_for_input)
            logging.info(f"Filled date field '{form_header}' with '{date_for_input}'")
            return True

        date_container_xpath = (
            f"//span[@class='M7eMe' and normalize-space(.)='{form_header}']/ancestor::div[@role='listitem']"
        )
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
        logging.info(f"Filled date field '{form_header}' with '{month}/{day}/{year}'")
        return True

    except NoSuchElementException:
        logging.warning(f"Date field '{form_header}' not found or unsupported")
        return False
    except Exception as e:
        logging.warning(f"Error filling date field '{form_header}': {e}")
        return False

def fill_file_upload(driver, form_header, file_urls, max_files):
    """Upload multiple files to a form field from Google Drive URLs."""
    uploaded_files = []
    file_urls = [url for url in file_urls if url and "drive.google.com" in url][:max_files]

    for idx, file_url in enumerate(file_urls, 1):
        temp_file = os.path.join(TEMP_DIR, f"image_{form_header.replace(' ', '_')}_{idx}_{int(time.time())}.jpg")
        try:
            # Download the file
            downloaded_file = download_google_drive_file(file_url, temp_file)
            if not downloaded_file or not os.path.exists(downloaded_file):
                logging.warning(f"Failed to download file from '{file_url}'")
                continue

            # Wait until file is fully written
            start_time = time.time()
            while time.time() - start_time < 15:
                if os.path.exists(downloaded_file) and os.path.getsize(downloaded_file) > 0:
                    break
                time.sleep(0.5)
            else:
                logging.warning(f"File not ready for upload: {downloaded_file}")
                continue

            # Find "Add File" button
            button_xpath = (
                f"//span[@class='M7eMe' and contains(normalize-space(.), '{form_header}')]/ancestor::div[@role='listitem']"
                f"//div[@role='button' and contains(@aria-label, 'Add File')]"
            )
            try:
                add_file_btn = WebDriverWait(driver, 20).until(
                    EC.element_to_be_clickable((By.XPATH, button_xpath))
                )
                scroll_into_view(driver, add_file_btn)
                driver.execute_script("arguments[0].click();", add_file_btn)
                logging.info(f"Clicked 'Add File' button for '{form_header}' (file {idx})")
            except Exception as e:
                logging.warning(f"Failed to click 'Add File' button for '{form_header}' (file {idx}): {e}")
                continue

            # Switch to file picker iframe
            try:
                WebDriverWait(driver, 20).until(
                    EC.frame_to_be_available_and_switch_to_it((By.XPATH, "//iframe[contains(@src, 'docs.google.com/picker')]"))
                )
                logging.info(f"Switched to file picker iframe for '{form_header}' (file {idx})")
            except Exception as e:
                logging.warning(f"Failed to switch to file picker iframe for '{form_header}' (file {idx}): {e}")
                driver.switch_to.default_content()
                continue

            # Upload the file
            try:
                file_input = WebDriverWait(driver, 20).until(
                    EC.presence_of_element_located((By.XPATH, "//input[@type='file']"))
                )
                file_input.send_keys(downloaded_file)
                logging.info(f"File sent to input: {downloaded_file} for '{form_header}' (file {idx})")

                # Wait for the file to appear in the picker
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//div[@class='docs-uploads-container']"))
                )

                # Click "Insert" button
                insert_btn = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//div[@role='dialog']//span[text()='Insert']/ancestor::button"))
                )
                driver.execute_script("arguments[0].click();", insert_btn)
                logging.info(f"Clicked Insert button for '{form_header}' (file {idx})")
                uploaded_files.append(downloaded_file)
            except Exception as e:
                logging.warning(f"File upload failed for '{form_header}' (file {idx}): {e}")
                driver.switch_to.default_content()
                continue
            finally:
                driver.switch_to.default_content()

        except Exception as e:
            logging.error(f"Unexpected error during file upload for '{form_header}' (file {idx}): {e}")
            driver.switch_to.default_content()
            continue

    return uploaded_files

def fill_google_form(driver, row, headers, header_mapping, file_upload_headers):
    """Fill and submit a Google Form for one row, handling multiple file upload fields."""
    try:
        driver.get(GOOGLE_FORM_URL)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//form")))

        fields_filled = True
        all_uploaded_files = []
        file_upload_urls = {header: [] for header in file_upload_headers}

        # Handle email collection checkbox
        try:
            checkbox_xpath = '//div[.//span[text()="Email"]]/following::div[@role="checkbox"][1]'
            checkbox = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, checkbox_xpath)))
            scroll_into_view(driver, checkbox)
            checkbox.click()
            logging.info("Checked 'Email' collection checkbox")
        except TimeoutException:
            logging.warning("Email collection checkbox not found or not clickable")

        # Collect file URLs for each file upload field
        for excel_header, value in zip(headers, row):
            if excel_header not in header_mapping or not value:
                continue
            form_header = header_mapping[excel_header]
            value_str = str(value).strip()

            # Check if the form header is a file upload field
            if form_header in file_upload_headers and value_str.startswith("http"):
                file_upload_urls[form_header].append(value_str)
                logging.debug(f"Added URL to '{form_header}': {value_str}")
                continue

            # Handle non-file fields (text, date, etc.)
            date_input_xpath = (
                f"//span[@class='M7eMe' and contains(normalize-space(.), '{form_header[:50]}')]/ancestor::div[@role='listitem']//"
                f"input[@type='date' or @aria-label='Month' or @aria-label='Day of the month' or @aria-label='Year']"
            )
            if driver.find_elements(By.XPATH, date_input_xpath):
                if not fill_date_field(driver, form_header, value):
                    fields_filled = False
                continue

            try:
                input_xpath = (
                    f"//span[@class='M7eMe' and contains(normalize-space(.), '{form_header[:50]}')]/ancestor::div[@role='listitem']//"
                    f"input[@type='text' or @type='number']"
                )
                inputs = driver.find_elements(By.XPATH, input_xpath)
                if inputs:
                    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, input_xpath)))
                    scroll_into_view(driver, inputs[0])
                    inputs[0].clear()
                    inputs[0].send_keys(value_str)
                    logging.info(f"Filled text field '{form_header}' with '{value_str}'")
                    continue

                textarea_xpath = (
                    f"//span[@class='M7eMe' and contains(normalize-space(.), '{form_header[:50]}')]/ancestor::div[@role='listitem']//textarea"
                )
                textareas = driver.find_elements(By.XPATH, textarea_xpath)
                if textareas:
                    WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, textarea_xpath)))
                    scroll_into_view(driver, textareas[0])
                    textareas[0].clear()
                    textareas[0].send_keys(value_str)
                    logging.info(f"Filled textarea field '{form_header}' with '{value_str}'")
                    continue

                dropdown_xpath = (
                    f"//span[@class='M7eMe' and contains(normalize-space(.), '{form_header[:50]}')]/ancestor::div[@role='listitem']//div[@role='listbox']"
                )
                dropdowns = driver.find_elements(By.XPATH, dropdown_xpath)
                if dropdowns:
                    scroll_into_view(driver, dropdowns[0])
                    dropdowns[0].click()
                    time.sleep(0.5)
                    options = driver.find_elements(By.XPATH, "//div[@role='option']")
                    matched = False
                    for opt in options:
                        if value_str.lower() in opt.text.lower():
                            scroll_into_view(driver, opt)
                            opt.click()
                            logging.info(f"Selected dropdown option '{opt.text}' for '{form_header}'")
                            matched = True
                            break
                    if not matched:
                        logging.warning(f"No matching dropdown option for '{form_header}' value '{value_str}'")
                        fields_filled = False
                    continue

                checkbox_xpath = (
                    f"//div[@role='listitem']//span[normalize-space(.)='{value_str}']/ancestor::label"
                )
                checkboxes = driver.find_elements(By.XPATH, checkbox_xpath)
                if checkboxes:
                    scroll_into_view(driver, checkboxes[0])
                    checkboxes[0].click()
                    logging.info(f"Checked checkbox '{form_header}' with value '{value_str}'")
                    continue

                logging.warning(f"No matching input found for field '{form_header}' with value '{value_str}'")
                fields_filled = False

            except Exception as e:
                logging.warning(f"Failed to process field '{form_header}' with value '{value_str}': {e}")
                fields_filled = False

        # Handle file uploads for detected file upload fields
        for idx, form_header in enumerate(file_upload_headers):
            if not file_upload_urls[form_header]:
                continue
            # First file upload field: up to 5 files; second: 1 file
            max_files = 5 if idx == 0 else 1
            uploaded_files = fill_file_upload(driver, form_header, file_upload_urls[form_header], max_files)
            if uploaded_files:
                logging.info(f"Completed upload to '{form_header}'")
                all_uploaded_files.extend(uploaded_files)
            else:
                logging.warning(f"Failed to upload to '{form_header}'")
                fields_filled = False

        # Submit the form
        try:
            submit_btn = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Submit')]/ancestor::div[@role='button']")))
            scroll_into_view(driver, submit_btn)
            submit_btn.click()
            WebDriverWait(driver, 10).until(EC.url_contains("formResponse"))
            logging.info("✅ Form submitted successfully")
        except Exception as e:
            logging.error(f"❌ Form submission failed: {e}")
            fields_filled = False

        # Clean up uploaded files
        for file_path in all_uploaded_files:
            try:
                if os.path.exists(file_path):
                    os.remove(file_path)
                    logging.info(f"Removed temporary file: {file_path}")
            except Exception as e:
                logging.warning(f"Failed to remove temporary file {file_path}: {e}")

        # Refresh the form to ensure a clean state for the next submission
        driver.get(GOOGLE_FORM_URL)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//form")))
        logging.info("🔄 Form page refreshed for next submission")

        return fields_filled

    except Exception as e:
        logging.error(f"Error filling form: {e}")
        return False

def main():
    """Main function to orchestrate the automation process."""
    driver = None
    try:
        kill_chrome()
        driver = setup_driver()
        form_headers, file_upload_headers = get_form_headers_and_file_fields(driver)
        excel_headers, rows = read_excel_data(EXCEL_FILE)
        matched_headers, unmatched_headers, header_mapping = match_headers(excel_headers, form_headers)

        for idx, row in enumerate(rows, start=2):
            logging.info(f"⏳ Processing row {idx}")
            fill_google_form(driver, row, excel_headers, header_mapping, file_upload_headers)
            time.sleep(1)  # Brief pause between submissions

    except Exception as e:
        logging.error(f"Unhandled error in main: {e}")
    finally:
        if driver:
            try:
                driver.quit()
                logging.info("🛑 Driver closed")
            except:
                pass
        if os.path.exists(TEMP_DIR):
            try:
                shutil.rmtree(TEMP_DIR)
                logging.info("🧹 Cleaned up temporary directory")
            except Exception as e:
                logging.warning(f"Failed to clean up temporary directory: {e}")

if __name__ == "__main__":
    main()