# file_downloader.py
import requests
from config import logging, TEMP_DIR
import os
import time

def download_google_drive_file(url, output_path=None):
    """Download a file from a Google Drive URL."""
    try:
        file_id = None
        if "id=" in url:
            file_id = url.split("id=")[1].split("&")[0]
        elif "/file/d/" in url:
            file_id = url.split("/file/d/")[1].split("/")[0]
        if not file_id:
            logging.warning(f"Invalid Google Drive URL: {url}")
            return None

        if not output_path:
            output_path = os.path.join(TEMP_DIR, f"file_{int(time.time() * 1000)}.jpg")

        download_url = f"https://drive.google.com/uc?export=download&id={file_id}"
        response = requests.get(download_url, stream=True, timeout=30)
        if response.status_code != 200:
            logging.warning(f"Failed to download file: Status {response.status_code}")
            return None

        if int(response.headers.get("Content-Length", 0)) > 10 * 1024 * 1024:
            logging.warning(f"File too large: {url}")
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