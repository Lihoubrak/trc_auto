import logging
import os
import time
from datetime import datetime
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from retrying import retry
from config import GOOGLE_FORM_URL
from image_utils import download_google_drive_image
import re

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
            # Verify the field value
            if date_inputs[0].get_attribute("value") == f"{year}-{month}-{day}":
                logging.info(f"Filled date field '{form_header}' with '{month}{day}{year}'")
                return True
            else:
                logging.warning(f"Date field '{form_header}' not filled correctly")
                return False

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
            # Verify the field value
            if inputs[field].get_attribute("value") != value:
                logging.warning(f"Date subfield '{field}' for '{form_header}' not filled correctly")
                return False
        logging.info(f"Filled date field '{form_header}' with '{month}/{day}/{year}'")
        return True

    except NoSuchElementException:
        logging.warning(f"Date field '{form_header}' not found")
        return False
    except Exception as e:
        logging.error(f"Error filling date field '{form_header}': {e}")
        return False

@retry(stop_max_attempt_number=3, wait_fixed=1000)
def fill_form_field(driver, form_header, value, field_type="text"):
    """Fill a specific form field based on its type."""
    try:
        value_str = str(value).strip()
        if not value_str:
            return False

        if field_type == "date":
            return fill_date_field(driver, form_header, value)

        form_header_cleaned = re.sub(r'\s+', ' ', form_header.replace("\xa0", " ")).strip()
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
            logging.info(f"Checkbox values for '{form_header}': {checkbox_values}")

            if not checkbox_values:
                logging.warning(f"No checkbox values provided for '{form_header}'")
                return False

            success = True
            for val in checkbox_values:
                checkbox_xpath = (
                    f"//span[@class='M7eMe' and contains(normalize-space(.), '{form_header_cleaned[:50]}')]"
                    f"/ancestor::div[@role='listitem']//div[@role='checkbox' and @data-answer-value='{val}']"
                )
                try:
                    checkbox = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH, checkbox_xpath))
                    )
                    scroll_into_view(driver, checkbox)
                    if checkbox.get_attribute("aria-checked") != "true":
                        driver.execute_script("arguments[0].click();", checkbox)
                        time.sleep(0.5)
                        # Verify checkbox state
                        if checkbox.get_attribute("aria-checked") != "true":
                            logging.warning(f"Failed to check checkbox '{val}' for '{form_header}'")
                            success = False
                        else:
                            logging.info(f"Checked checkbox '{val}' for '{form_header}'")
                    else:
                        logging.info(f"Checkbox '{val}' for '{form_header}' already checked")
                except TimeoutException:
                    logging.warning(f"Checkbox '{val}' not found or not clickable for '{form_header}'")
                    success = False
                except Exception as e:
                    logging.error(f"Error checking checkbox '{val}' for '{form_header}': {e}")
                    success = False
            return success

        elements = driver.find_elements(By.XPATH, xpath)
        if not elements:
            logging.warning(f"No elements found for form header '{form_header}' and type '{field_type}'")
            return False
    
        scroll_into_view(driver, elements[0])
        if field_type in ["text", "textarea"]:
            elements[0].clear()
            elements[0].send_keys(value_str)
            # Verify the field value
            if elements[0].get_attribute("value") == value_str:
                logging.info(f"Filled {field_type} field '{form_header}' with '{value_str}'")
                return True
            else:
                logging.warning(f"Failed to fill {field_type} field '{form_header}' with '{value_str}'")
                return False
        elif field_type == "dropdown":
            elements[0].click()
            time.sleep(0.5)
            option_xpath = f"//div[@role='option' and contains(., '{value_str[:50]}')]"
            options = driver.find_elements(By.XPATH, option_xpath)
            if options:
                scroll_into_view(driver, options[0])
                options[0].click()
                time.sleep(0.5)
                # Verify dropdown selection
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

def fill_google_form(driver, row, headers, header_mapping):
    """Fill and submit a Google Form for one row of data."""
    temp_files = []
    fields_filled = True

    try:
        driver.get(GOOGLE_FORM_URL)
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, "//form")))
        logging.info("Google Form loaded successfully")
        # Handle email checkbox once at the start
        # email_checkbox_handled = False
        # if not email_checkbox_handled:
        #     try:
        #         checkbox_xpath = '//div[.//span[text()="Email"]]/following::div[@role="checkbox"][1]'
        #         checkbox = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, checkbox_xpath)))
        #         scroll_into_view(driver, checkbox)
                
        #         # Check if the checkbox is already checked
        #         is_checked = checkbox.get_attribute("aria-checked") == "true"
        #         if not is_checked:
        #             checkbox.click()
        #             logging.info("Checked 'Email' collection checkbox")
        #         else:
        #             logging.info("Email checkbox already checked, skipping")
        #         email_checkbox_handled = True
        #     except TimeoutException:
        #         logging.info("No email checkbox found — skipping")
        #         email_checkbox_handled = True
        #     except Exception as e:
        #         logging.error(f"Error handling email checkbox: {e}")
        #         email_checkbox_handled = True

        # Process each header sequentially
        for excel_header, value in zip(headers, row):
            if excel_header not in header_mapping or not value:
                logging.info(f"Skipping empty or unmapped field: {excel_header}")
                continue
            form_header = header_mapping[excel_header]
            logging.info(f"Processing field: {form_header}")

            # Check if the field is an image upload field
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
                                upload_btn_xpath = (
                                    f"//span[@class='M7eMe' and contains(normalize-space(.), '{form_header}')]/ancestor::div[@role='listitem']//"
                                    f"div[@role='button' and contains(@aria-label, 'Add File')]"
                                )
                                upload_button = WebDriverWait(driver, 15).until(
                                    EC.element_to_be_clickable((By.XPATH, upload_btn_xpath))
                                )
                                scroll_into_view(driver, upload_button)
                                logging.info(f"Located 'Add File' button for '{form_header}'")

                                driver.execute_script("arguments[0].click();", upload_button)
                                logging.info("Clicked 'Add File' button")

                                picker_dialog_xpath = "//div[contains(@class, 'fFW7wc XKSfm-Sx9Kwc picker-dialog') and not(contains(@style, 'display: none'))]"
                                WebDriverWait(driver, 20).until(
                                    EC.presence_of_element_located((By.XPATH, picker_dialog_xpath))
                                )
                                logging.info("Picker dialog div is visible")

                                iframe_xpath = f"{picker_dialog_xpath}//iframe[contains(@src, 'docs.google.com/picker')]"
                                WebDriverWait(driver, 20).until(
                                    EC.frame_to_be_available_and_switch_to_it((By.XPATH, iframe_xpath))
                                )
                                logging.info("Switched to new file picker iframe")

                                file_input = WebDriverWait(driver, 15).until(
                                    EC.presence_of_element_located((By.XPATH, "//input[@type='file']"))
                                )
                                file_input.send_keys(temp_file_path)
                                logging.info(f"Sent file path to input: {temp_file_path}")

                                driver.switch_to.default_content()

                                WebDriverWait(driver, 30).until(
                                    EC.invisibility_of_element_located((By.XPATH, picker_dialog_xpath))
                                )
                                logging.info("Picker dialog and iframe closed")

                                file_list_xpath = (
                                    f"//span[@class='M7eMe' and contains(normalize-space(.), '{form_header}')]/ancestor::div[@role='listitem']"
                                    f"//div[@role='listitem']//div[contains(@class, 'ZXoVYe ybj8pf') and contains(text(), '{file_name}')]"
                                )
                                WebDriverWait(driver, 20).until(
                                    EC.presence_of_element_located((By.XPATH, file_list_xpath))
                                )
                                logging.info(f"File '{file_name}' confirmed in form's file list")
                                time.sleep(1)  # Added: Ensure file is fully processed
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
                time.sleep(0.5)  # Added: Pause after image field processing
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

            # Try filling the field with all possible field types
            filled = False
            for field_type in ["date", "text", "textarea", "dropdown", "checkbox"]:
                if fill_form_field(driver, form_header, value, field_type):
                    filled = True
                    break
                time.sleep(0.5)  # Brief pause between field type attempts

            if not filled:
                logging.warning(f"Failed to fill field '{form_header}' with value '{value}'")
                fields_filled = False

        # Submit the form only if all fields were filled successfully
        # Uncomment and adjust as needed
        # if fields_filled:
        #     try:
        #         submit_btn = WebDriverWait(driver, 10).until(
        #             EC.element_to_be_clickable((By.XPATH, "//span[text()='Submit']/ancestor::div[@role='button']"))
        #         )
        #         scroll_into_view(driver, submit_btn)
        #         driver.execute_script("arguments[0].click();", submit_btn)
        #         WebDriverWait(driver, 15).until(EC.url_contains("formResponse"))
        #         logging.info("Form submitted successfully")
        #         return True
        #     except Exception as e:
        #         logging.error(f"Form submission failed: {e}")
        #         fields_filled = False

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