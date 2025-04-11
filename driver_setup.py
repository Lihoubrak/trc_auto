# driver_setup.py
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from config import logging

try:
    from webdriver_manager.chrome import ChromeDriverManager
    USE_WEBDRIVER_MANAGER = True
except ImportError:
    USE_WEBDRIVER_MANAGER = False

def setup_driver():
    """Set up Chrome WebDriver with options to mimic a real user."""
    options = Options()
    user_data_dir = r"C:\Users\KHC\AppData\Local\Google\Chrome\User Data"
    profile_dir = "Profile 1"
    options.add_argument(f"--user-data-dir={user_data_dir}")
    options.add_argument(f"--profile-directory={profile_dir}")
    options.add_argument("--start-maximized")
    options.add_argument("--disable-notifications")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/135.0.0.0 Safari/537.36")

    try:
        if USE_WEBDRIVER_MANAGER:
            driver = webdriver.Chrome(service=webdriver.chrome.service.Service(ChromeDriverManager().install()), options=options)
        else:
            driver = webdriver.Chrome(service=webdriver.chrome.service.Service('chromedriver.exe'), options=options)
        driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
            "source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
        })
        logging.info("WebDriver initialized successfully.")
        return driver
    except Exception as e:
        logging.error(f"Failed to set up WebDriver: {e}")
        raise