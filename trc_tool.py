import time
import os
import logging
import openpyxl
import requests
import tempfile
import shutil
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from fuzzywuzzy import fuzz
from pynput.keyboard import Key, Controller

try:
    from webdriver_manager.chrome import ChromeDriverManager
    USE_WEBDRIVER_MANAGER = True
except ImportError:
    USE_WEBDRIVER_MANAGER = False

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

GOOGLE_FORM_URL = "https://docs.google.com/forms/d/e/1FAIpQLSeqWvnn4KIru5BYd6aNVTCvaej6KvPWdbK0tN3piOgU8u8ftg/viewform?usp=dialog"
EXCEL_FILE = "MAINTENANCE CABLE REQUEST TO VTC.xlsx"
SIMILARITY_THRESHOLD = 80
TEMP_DIR = tempfile.mkdtemp()

def setup_driver():
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
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/135.0.0.0 Safari/537.36")

    try:
        if USE_WEBDRIVER_MANAGER:
            driver = webdriver.Chrome(service=webdriver.chrome.service.Service(ChromeDriverManager().install()), options=options)
        else:
            driver = webdriver.Chrome(service=webdriver.chrome.service.Service('chromedriver.exe'), options=options)
    except Exception as e:
        logging.error(f"Failed to initialize WebDriver: {e}")
        raise

    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
    })
    return driver

def download_google_drive_file(url, output_path):
    try:
        file_id = None
        if "id=" in url:
            file_id = url.split("id=")[1].split("&")[0]
        elif "/file/d/" in url:
            file_id = url.split("/file/d/")[1].split("/")[0]

        if not file_id:
            logging.warning(f"Invalid Google Drive URL: {url}")
            return None

        download_url = f"https://drive.google.com/uc?export=download&id={file_id}"
        response = requests.get(download_url, stream=True)

        if response.status_code != 200:
            logging.warning(f"Failed to download file from {url}: Status {response.status_code}")
            return None

        content_length = response.headers.get("Content-Length")
        if content_length and int(content_length) > 10 * 1024 * 1024:
            logging.warning(f"File at {url} exceeds 10MB limit")
            return None

        with open(output_path, "wb") as f:
            for chunk in response.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)
        logging.info(f"Downloaded file to {output_path}")
        return output_path
    except Exception as e:
        logging.warning(f"Error downloading file from {url}: {e}")
        return None

def read_excel_data(filepath):
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
        return headers, data
    except Exception as e:
        logging.error(f"Error reading Excel file: {e}")
        raise

def get_form_headers(driver):
    try:
        driver.get(GOOGLE_FORM_URL)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//span[@class='M7eMe']")))
        header_elements = driver.find_elements(By.XPATH, "//span[@class='M7eMe']")
        headers = [elem.text.strip() for elem in header_elements if elem.text.strip()]
        if not headers:
            raise ValueError("No form headers found.")
        return headers
    except Exception as e:
        logging.error(f"Error fetching form headers: {e}")
        raise

def match_headers(excel_headers, form_headers):
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
        else:
            unmatched.append(excel_header)
    return matched_headers, unmatched, mapping

def scroll_into_view(driver, element):
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
    time.sleep(0.3)

def clear_inputs(driver):
    inputs = driver.find_elements(By.XPATH, "//input[@type='text'] | //input[@type='date']")
    for field in inputs:
        try:
            scroll_into_view(driver, field)
            field.clear()
        except:
            continue

def parse_date(value):
    if not value:
        return None
    if isinstance(value, datetime):
        return value.strftime("%m/%d/%Y")
    if isinstance(value, str):
        try:
            parsed_date = datetime.strptime(value, "%Y-%m-%d")
            return parsed_date.strftime("%m/%d/%Y")
        except ValueError:
            try:
                parsed_date = datetime.strptime(value, "%m/%d/%Y")
                return parsed_date.strftime("%m/%d/%Y")
            except ValueError:
                try:
                    parsed_date = datetime.strptime(value, "%d/%m/%Y")
                    return parsed_date.strftime("%m/%d/%Y")
                except ValueError:
                    logging.warning(f"Unsupported date format: {value}")
                    return None
    return None

def fill_date_field(driver, form_header, date_value):
    try:
        date_value = parse_date(date_value)
        if not date_value:
            return False

        month, day, year = date_value.split("/")

        date_input_xpath = f"//span[@class='M7eMe' and normalize-space(.)='{form_header}']/ancestor::div[@role='listitem']//input[@type='date']"
        date_inputs = driver.find_elements(By.XPATH, date_input_xpath)
        if date_inputs:
            scroll_into_view(driver, date_inputs[0])
            date_for_input = f"{month}{day}{year}"
            date_inputs[0].send_keys(date_for_input)
            logging.info(f"Filled date field '{form_header}' with '{date_for_input}'")
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
        logging.info(f"Filled date field '{form_header}' with '{month}/{day}/{year}'")
        return True

    except NoSuchElementException:
        logging.warning(f"Date field '{form_header}' not found or unsupported.")
        return False
    except Exception as e:
        logging.warning(f"Error filling date field '{form_header}': {e}")
        return False

def fill_file_upload(driver, form_header, file_url):
    try:
        if not file_url or "drive.google.com" not in file_url:
            logging.warning(f"Invalid or missing Google Drive URL: '{file_url}'")
            return False

        temp_file = os.path.join(TEMP_DIR, f"image_{int(time.time() * 1000)}.jpg")
        downloaded_file = download_google_drive_file(file_url, temp_file)
        if not downloaded_file:
            return False

        # Locate and click the "Add File" button for the specific header
        button_xpath = (f"//span[@class='M7eMe' and normalize-space(.)='{form_header}']/ancestor::div[@role='listitem']//"
                       f"div[@role='button' and contains(@aria-label, 'Add File')]")
        upload_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, button_xpath))
        )
        scroll_into_view(driver, upload_button)
        upload_button.click()
        logging.info(f"Clicked 'Add File' button for '{form_header}'")

        # Wait for the modal to appear
        modal_xpath = "//div[contains(@class, 'VfPpkd-ksKsZd-mWPk3d')]"
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, modal_xpath))
        )
        logging.info(f"Modal dialog appeared for '{form_header}'")

        # Locate the "Browse" button within the modal
        browse_button_xpath = "//div[contains(@class, 'VfPpkd-ksKsZd-mWPk3d')]//button[@jsname='PX1Pzd']//span[@jsname='V67aGc' and text()='Browse']/ancestor::button"
        browse_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, browse_button_xpath))
        )
        browse_button.click()
        logging.info(f"Clicked 'Browse' button for '{form_header}'")

        # Use pynput to input the file path in the File Explorer dialog
        keyboard = Controller()
        time.sleep(1)  # Wait for the File Explorer dialog to open
        keyboard.type(downloaded_file)
        logging.info(f"Typed file path '{downloaded_file}' into File Explorer for '{form_header}'")
        keyboard.press(Key.enter)
        keyboard.release(Key.enter)
        logging.info(f"Pressed Enter to confirm file selection for '{form_header}'")

        # Wait for the uploaded file to appear in the "Selected files" list
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, f"//span[@class='M7eMe' and normalize-space(.)='{form_header}']/ancestor::div[@role='listitem']//div[@role='list' and @aria-label='Selected files']//div"))
        )
        logging.info(f"File upload confirmed for '{form_header}'")
        return True

    except TimeoutException:
        logging.warning(f"File upload for '{file_url}' took too long or elements not found.")
        return False
    except Exception as e:
        logging.warning(f"Error uploading file '{file_url}' to field '{form_header}': {e}")
        return False
    finally:
        if os.path.exists(temp_file):
            try:
                os.remove(temp_file)
            except:
                pass

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

def main():
    driver = None
    try:
        driver = setup_driver()
        form_headers = get_form_headers(driver)
        excel_headers, rows = read_excel_data(EXCEL_FILE)
        matched, unmatched, mapping = match_headers(excel_headers, form_headers)

        if unmatched:
            logging.warning(f"Unmatched headers in Excel: {unmatched}")
        logging.info(f"Matched Headers Mapping: {mapping}")

        for idx, row in enumerate(rows, start=2):
            logging.info(f"⏳ Processing row {idx}")
            fill_google_form(driver, row, excel_headers, mapping)
            time.sleep(1)

    except Exception as e:
        logging.error(f"Unhandled error in main: {e}")
    finally:
        if driver:
            driver.quit()
            logging.info("🛑 Driver closed.")
        if os.path.exists(TEMP_DIR):
            try:
                shutil.rmtree(TEMP_DIR)
            except:
                pass

if __name__ == "__main__":
    main()