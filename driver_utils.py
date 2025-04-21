import logging
import psutil
from selenium import webdriver
from selenium.webdriver.chrome.service import Service

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

def initialize_driver(config):
    """Initialize and configure Chrome WebDriver."""
    options = webdriver.ChromeOptions()
    options.add_argument(f"--user-data-dir={config['USER_DATA_DIR']}")
    options.add_argument(f"--profile-directory={config['PROFILE_DIR']}")
    options.add_argument("--start-maximized")
    options.add_argument("--lang=km-KH")
    options.add_argument("--disable-notifications")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/135.0.0.0 Safari/537.36"
    )

    try:
        service = (
            Service(ChromeDriverManager().install())
            if USE_WEBDRIVER_MANAGER
            else Service(CHROMEDRIVER_PATH)
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