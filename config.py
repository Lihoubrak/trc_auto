# config.py
import logging
import tempfile

# Set up logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

# Constants
GOOGLE_FORM_URL = "https://docs.google.com/forms/d/e/1FAIpQLSeqWvnn4KIru5BYd6aNVTCvaej6KvPWdbK0tN3piOgU8u8ftg/viewform?usp=dialog"
EXCEL_FILE = "MAINTENANCE CABLE REQUEST TO VTC.xlsx"
SIMILARITY_THRESHOLD = 80  # For header and value matching
TEMP_DIR = tempfile.mkdtemp()  # Temporary folder for file downloads