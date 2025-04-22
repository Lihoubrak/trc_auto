import logging
import os
import time
from datetime import datetime
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from retrying import retry
from image_utils import download_google_drive_image
import re

# Set up logging configuration at the start
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),  # Output to console
        logging.FileHandler('form_fill.log')  # Output to file
    ]
)

# Define field type mapping for specific headers
FIELD_TYPE_MAPPING = {
    "កាលបរិច្ឆេទនៃការកាត់ផ្តាច់-Date of Damage:": "date",
    "កាលបរិច្ឆេទគ្រោងបញ្ចប់ការថែទាំ-Finished Date of Repairing:": "date",
    "រូបភាពនៃស្ថានភាពការខូចខាតនៃខ្សែកាប្លិ៍ដោយមាន lat/long <10MB (Picture of Damage Cable with lat/long <10MB):": "file",
    "រូបភាពនៃគំនូសនៅលើ Google Map ដែលមានចំណុចចាប់ផ្តើមនិងបញ្ចប់ (Picture of drawing in google map with start and end lat/long):": "file",
    "Email": "checkbox",
    "ស្នើចុះថែទាំខ្សែក្នុងវង្វង់ចំណុចចាប់ផ្តើមនិងបញ្ចប់ Lat/Long ខាងលើ ដើម្បី៖": "checkbox",
    "ក្រុមហ៊ុនស្នើសុំ-Requested Company:": "dropdown",
    "ប្រភេទហេដ្ឋារចនាសម្ព័ន្ធ-Type of Infrastructure:": "dropdown",
    "ប្រភេទនៃការជួសជុល-Overhead or Underground:": "dropdown",
    "ខេត្ត/ក្រុង": "dropdown",
    "ជួសជុលជូនក្រុមហ៊ុនឬអតិថិជន-Repair for company/customers:": "text",
    "ចំនួនខ្សែកាប្លិ៍xចំនួន\nបណ្តូលខ្សែកាប្លិ៍ -Number of cable * Core:": "text",
    "អាសយដ្ឋានចំនុចចាប់ផ្តើម(លេខផ្លូវ ភូមិ ឃុំ/សង្កាត់ ស្រុក/ខណ្ឌ ខេត្ត/រាជធានី)-Starting Address:": "text",
    "អាសយដ្ឋានចំណុចបញ្ចប់ (លេខផ្លូវ ភូមិ ឃុំ/សង្កាត់ ស្រុក/ខណ្ឌ ខេត្ត/រាជធានី)  -Ending Address:": "text",
    "ចំណុចចាប់ផ្តើម រយៈទទឹងនិងបណ្តោយ (Start: Lat ,Long)": "text",
    "ចំណុចបញ្ចប់ រយៈទទឹងនិងបណ្តោយ (End: Lat ,Long)": "text",
    "ប្រវែងខ្សែកាប្លិ៍ដែលត្រូវជួសថ្មីគឺត្រូវតិចជាង២០០ម៉ែត្រ-Length of replacement broken cable must be less than 200m (ករណីប្រវែងខ្សែច្រើនជាង ២០០ម៉ែត្រក្រុមហ៊ុនត្រូវអនុវត្តនីតិវិធីសាងសង់ថ្មី- If length of cable is larger than 200m,company has to follow the new build procedure):": "text",
    "មូលហេតុនៃការកាត់ផ្តាច់\nCable Incident": "textarea"
}

def get_form_headers(driver, config):
    """Retrieve headers from the Google Form."""
    try:
        driver.get(config.get("GOOGLE_FORM_URL"))
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

def scroll_into_view(driver, element):
    """Scroll an element into view."""
    driver.execute_script("arguments[0].scrollIntoView({block: 'center', behavior: 'smooth'});", element)
    time.sleep(0.1)  # Reduced from 0.2s for faster execution
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
        print("...date_value",date_value)
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
    
@retry(stop_max_attempt_number=3, wait_fixed=500)  # Reduced from 1000ms for faster retries
def fill_form_field(driver, form_header, value, field_type="text"):
    """Fill a specific form field based on its type, optimized for speed and reliability."""
    try:
        value_str = str(value).strip()
        if not value_str:
            logging.info(f"Skipping empty value for field '{form_header}'")
            return False

        if field_type == "date":
            return fill_date_field(driver, form_header, value)

        form_header_cleaned = re.sub(r'\s+', ' ', form_header.replace("\xa0", " ")).strip()
        logging.debug(f"Processing field: {form_header_cleaned} with type: {field_type}")

        xpath_map = {
            "text": (
                f"//span[@class='M7eMe' and contains(normalize-space(.), '{form_header_cleaned[:50]}')]"
                f"/ancestor::div[@role='listitem']//input[@type='text' or @type='number']"
            ),
            "textarea": (
                f"//span[@class='M7eMe' and contains(normalize-space(.), '{form_header_cleaned[:50]}')]"
                f"/ancestor::div[@role='listitem']//textarea"
            ),
            "dropdown": (
                f"//span[@class='M7eMe' and contains(normalize-space(.), '{form_header_cleaned[:50]}')]"
                f"/ancestor::div[@role='listitem']//div[@role='listbox']"
            ),
            "checkbox": (
                f"//span[@class='M7eMe' and contains(normalize-space(.), '{form_header_cleaned[:50]}')]"
                f"/ancestor::div[@role='listitem']//div[@role='checkbox']"
            ),
        }

        xpath = xpath_map.get(field_type)
        if not xpath:
            logging.warning(f"Unsupported field type: {field_type}")
            return False

        if field_type == "checkbox":
            checkbox_values = [v.strip() for v in value_str.split(",") if v.strip()]
            if not checkbox_values:
                logging.warning(f"No checkbox values provided for '{form_header}'")
                return False

            success = True
            for val in checkbox_values:
                checkbox_xpath = (
                    f"//span[@class='M7eMe' and contains(normalize-space(.), '{form_header_cleaned[:50]}')]"
                    f"/ancestor::div[@role='listitem']//div[@role='checkbox' and @data-answer-value='{val}']"
                )
                checkbox = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, checkbox_xpath))
                )
                scroll_into_view(driver, checkbox)
                if checkbox.get_attribute("aria-checked") != "true":
                    driver.execute_script("arguments[0].click();", checkbox)
                    time.sleep(0.2)  # Reduced from 0.5s
                    if checkbox.get_attribute("aria-checked") != "true":
                        logging.warning(f"Failed to check checkbox '{val}' for '{form_header}'")
                        success = False
                logging.info(f"Checked checkbox '{val}' for '{form_header}'")
            return success

        elements = driver.find_elements(By.XPATH, xpath)
        if not elements:
            logging.warning(f"No elements found for form header '{form_header}' and type '{field_type}'")
            return False

        scroll_into_view(driver, elements[0])
        if field_type in ["text", "textarea"]:
            elements[0].clear()
            elements[0].send_keys(value_str)
            time.sleep(0.1)  # Reduced from 0.5s
            if elements[0].get_attribute("value") == value_str:
                logging.info(f"Filled {field_type} field '{form_header}' with '{value_str}'")
                return True
            else:
                logging.warning(f"Failed to fill {field_type} field '{form_header}' with '{value_str}'")
                return False
        elif field_type == "dropdown":
            elements[0].click()
            time.sleep(0.2)  # Reduced from 0.5s
            option_xpath = f"//div[@role='option' and contains(normalize-space(.), '{value_str[:30]}')]"
            options = driver.find_elements(By.XPATH, option_xpath)
            if options:
                scroll_into_view(driver, options[0])
                driver.execute_script("arguments[0].click();", options[0])  # Use JS click for reliability
                time.sleep(0.2)  # Reduced from 0.5s
                selected = driver.find_element(By.XPATH, xpath).text
                if value_str in selected:
                    logging.info(f"Selected dropdown option '{value_str}' for '{form_header}'")
                    return True
                else:
                    logging.warning(f"Failed to select dropdown option '{value_str}' for '{form_header}'")
                    return False
            logging.warning(f"No matching dropdown option for '{form_header}' value '{value_str}'")
            return False

    except Exception as e:
        logging.error(f"Error filling form field '{form_header}' with type '{field_type}': {e}")
        return False

def fill_google_form(driver, row, headers, header_mapping, config):
    """Fill and submit a Google Form for one row of data."""
    temp_files = []
    fields_filled = True

    try:
        driver.get(config.get("GOOGLE_FORM_URL"))
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, "//form")))
        logging.info("Google Form loaded successfully")

        # Handle email checkbox once at the start
        email_checkbox_handled = False
        if not email_checkbox_handled:
            try:
                checkbox_xpath = '//div[.//span[text()="Email"]]/following::div[@role="checkbox"][1]'
                checkbox = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, checkbox_xpath)))
                scroll_into_view(driver, checkbox)
                is_checked = checkbox.get_attribute("aria-checked") == "true"
                if not is_checked:
                    checkbox.click()
                    logging.info("Checked 'Email' collection checkbox")
                else:
                    logging.info("Email checkbox already checked, skipping")
                email_checkbox_handled = True
            except TimeoutException:
                logging.info("No email checkbox found — skipping")
                email_checkbox_handled = True
            except Exception as e:
                logging.error(f"Error handling email checkbox: {e}")
                email_checkbox_handled = True

        # Process each header sequentially
        for excel_header, value in zip(headers, row):
            if excel_header not in header_mapping or not value:
                logging.info(f"Skipping empty or unmapped field: {excel_header}")
                continue
            form_header = header_mapping[excel_header]
            logging.info(f"Processing field: {form_header}")

            # Get the field type from FIELD_TYPE_MAPPING
            field_type = FIELD_TYPE_MAPPING.get(form_header)
            if not field_type:
                logging.warning(f"No field type defined for '{form_header}'")
                fields_filled = False
                continue
            # Handle file upload fields
            if field_type == "file":
                if isinstance(value, str) and "drive.google.com" in value:
                    # Log download start time
                    start_time = time.time()
                    temp_file_path = download_google_drive_image(value)
                    download_duration = time.time() - start_time
                    logging.info(f"Download took {download_duration:.2f} seconds for URL: {value}")

                    if temp_file_path:
                        temp_files.append(temp_file_path)
                        uploaded = False
                        file_name = os.path.basename(temp_file_path)
                        logging.debug(f"Attempting to upload file: {file_name} for field: {form_header}")

                        @retry(stop_max_attempt_number=3, wait_fixed=1000)
                        def attempt_upload():
                            nonlocal uploaded
                            try:
                                import unicodedata
                                form_header_normalized = unicodedata.normalize('NFKC', form_header).strip()
                                upload_btn_xpath = (
                                    f"//span[@class='M7eMe' and contains(normalize-space(.), '{form_header_normalized}')]/ancestor::div[@role='listitem']//"
                                    f"div[@role='button' and contains(@aria-label, 'Add File')]"
                                )
                                upload_button = WebDriverWait(driver, 30).until(
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

                                file_input = WebDriverWait(driver, 15).until(
                                    EC.presence_of_element_located((By.XPATH, "//input[@type='file']"))
                                )
                                file_input.send_keys(temp_file_path)
                                driver.switch_to.default_content()
                                time.sleep(1)  # Wait for form to update

                                # Simplified XPath for uploaded file verification
                                file_list_xpath = (
                                    f"//span[@class='M7eMe' and contains(normalize-space(.), '{form_header_normalized}')]/ancestor::div[@role='listitem']"
                                    f"//div[@role='listitem']//div[contains(text(), '{file_name}')]"
                                )
                                file_element = WebDriverWait(driver, 45).until(
                                    EC.presence_of_element_located((By.XPATH, file_list_xpath))
                                )
                                displayed_file_name = file_element.text.strip()

                                if file_name.lower() == displayed_file_name.lower():
                                    logging.info(f"File name matched: expected '{file_name}', got '{displayed_file_name}'")
                                    uploaded = True
                                else:
                                    logging.warning(f"File name mismatch: expected '{file_name}', got '{displayed_file_name}'")
                                    uploaded = False
                                time.sleep(0.5)
                                return uploaded

                            except TimeoutException as te:
                                logging.error(f"Timeout during file upload attempt for '{form_header}': {te}")
                                with open("page_source.html", "w", encoding="utf-8") as f:
                                    f.write(driver.page_source)
                                driver.save_screenshot("timeout_screenshot.png")
                                logging.info("Saved page source and screenshot for debugging")
                                driver.switch_to.default_content()
                                raise
                            except Exception as e:
                                logging.error(f"Error during file upload attempt for '{form_header}': {e}")
                                driver.switch_to.default_content()
                                raise

                        try:
                            uploaded = attempt_upload()
                            if uploaded:
                                logging.info(f"Successfully uploaded file '{file_name}' for '{form_header}'")
                            else:
                                logging.error(f"Failed to upload file '{file_name}' for '{form_header}'")
                                fields_filled = False
                        except Exception as e:
                            logging.error(f"Failed to upload file '{file_name}' for '{form_header}' after retries: {e}")
                            fields_filled = False
                    else:
                        logging.warning(f"Failed to download image from Google Drive for '{form_header}': {value}")
                        fields_filled = False
                else:
                    logging.warning(f"Invalid Google Drive URL for image field '{form_header}': {value}")
                    fields_filled = False
                time.sleep(0.3)
                continue
            # Handle special case for "Number of cable * Core"
            form_header_cleaned = form_header.replace("\n", "").strip()
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
                        input_elements[0].send_keys(value)
                        time.sleep(0.5)
                        # Verify the field value
                        if input_elements[0].get_attribute("value") == str(value):
                            logging.info(f"Filled 'Number of cable * Core' with value: {value}")
                        else:
                            logging.warning(f"Failed to fill 'Number of cable * Core' with value: {value}")
                            fields_filled = False
                    else:
                        logging.warning("No input element found for 'Number of cable * Core'")
                        fields_filled = False
                except Exception as e:
                    logging.error(f"Error filling 'Number of cable * Core': {e}")
                    fields_filled = False
                continue
            # Fill the field using the specified field type
            if not fill_form_field(driver, form_header, value, field_type):
                logging.warning(f"Failed to fill field '{form_header}' with value '{value}'")
                fields_filled = False
            time.sleep(0.2)  # Reduced from 0.5s

        # Submit the form only if all fields were filled successfully
        if fields_filled:
            try:
                submit_btn = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//span[text()='Submit']/ancestor::div[@role='button']"))
                )
                scroll_into_view(driver, submit_btn)
                driver.execute_script("arguments[0].click();", submit_btn)
                WebDriverWait(driver, 15).until(EC.url_contains("formResponse"))
                logging.info("Form submitted successfully")
                return True
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
