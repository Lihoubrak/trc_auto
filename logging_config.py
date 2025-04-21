import logging
import sys
import codecs

def configure_logging():
    """Configure logging with a consistent format and UTF-8 encoding."""
    # Create a logger
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)

    # Clear any existing handlers to avoid duplicates
    logger.handlers = []

    # Create a console handler with UTF-8 encoding
    console_handler = logging.StreamHandler(stream=sys.stdout)
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))
    # Ensure the stream uses UTF-8 encoding
    console_handler.stream = codecs.getwriter('utf-8')(sys.stdout.detach())

    # Optionally, add a file handler with UTF-8 encoding
    file_handler = logging.FileHandler('app.log', encoding='utf-8', mode='a')
    file_handler.setLevel(logging.INFO)
    file_handler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))

    # Add handlers to the logger
    logger.addHandler(console_handler)
    logger.addHandler(file_handler)