from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import os
import logging
import time
# === CONFIG ===
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
form_url = "https://docs.google.com/forms/d/e/1FAIpQLSeMSCrn5RZNOrR86huoCQ-w2YKgVCscr1UTbU7uVO7yYqRccg/viewform"
file_path = os.path.abspath("image.png")

# Validate file
if not os.path.exists(file_path):
    raise FileNotFoundError(f"File not found: {file_path}")
if os.path.getsize(file_path) > 10 * 1024 * 1024:
    raise ValueError(f"File size exceeds 10 MB: {file_path}")
if not file_path.lower().endswith(('.bmp', '.gif', '.heic', '.heif', '.jpeg', '.jpg', '.png', '.tiff', '.ico', '.webp')):
    raise ValueError(f"Unsupported file type: {file_path}")

# === Setup Chrome Options with User Profile ===
options = webdriver.ChromeOptions()
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

# === Initialize WebDriver ===
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
wait = WebDriverWait(driver, 20)

try:
    # Open Google Form
    logging.info("Opening Google Form")
    driver.get(form_url)

    # Wait for and click "Add File" button
    logging.info("Clicking Add File button")
    add_file_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "div[role='button'][aria-label='Add File']")))
    add_file_btn.click()

    # Wait for iframe and switch to it with retry mechanism
    logging.info("Attempting to switch to Google Picker iframe")
    max_attempts = 3
    for attempt in range(max_attempts):
        try:
            wait.until(EC.frame_to_be_available_and_switch_to_it((By.XPATH, "//iframe[contains(@src, 'docs.google.com/picker')]")))
            logging.info("Successfully switched to iframe")
            break
        except TimeoutException:
            logging.warning(f"Iframe switch attempt {attempt + 1}/{max_attempts} failed")
            if attempt == max_attempts - 1:
                raise TimeoutException("Failed to switch to iframe after multiple attempts")
            time.sleep(2)  # Wait before retrying
    
    # Wait for the uploaded file to appear and click "Insert"
    logging.info("Waiting for Insert button after upload")
    insert_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(@class, 'VfPpkd-ksKsZd-mWPk3d')]//button[@jsname='PX1Pzd']//span[@jsname='V67aGc' and text()='Browse']/ancestor::button")))
    insert_btn.click()

    # Switch back to main content
    logging.info("Switching back to main content")
    driver.switch_to.default_content()

    # Click Submit
    logging.info("Clicking Submit button")
    submit_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Submit']/ancestor::div[@role='button' and not(@aria-disabled='true')]")))
    submit_btn.click()

    logging.info("✅ Form submitted successfully.")

except TimeoutException as e:
    logging.error(f"❌ Timed out waiting for element: {e}")
except NoSuchElementException as e:
    logging.error(f"❌ Element not found: {e}")
except Exception as e:
    logging.error(f"❌ Unexpected error: {e}")
finally:
    logging.info("Closing browser")
    driver.quit()