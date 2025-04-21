import logging
import os
import re
import tempfile
import requests
from urllib.parse import urlparse

def download_google_drive_image(google_drive_link, temp_dir="images"):
    """Download an image from a Google Drive link."""
    try:
        valid_patterns = [
            r'https?://drive\.google\.com/file/d/([-\w]{25,})/',
            r'https?://drive\.google\.com/uc\?id=([-\w]{25,})',
            r'https?://drive\.google\.com/open\?id=([-\w]{25,})'
        ]
        file_id = None
        for pattern in valid_patterns:
            match = re.search(pattern, google_drive_link)
            if match:
                file_id = match.group(1)
                break
        
        if not file_id:
            logging.error(f"Invalid or unsupported Google Drive link: {google_drive_link}")
            return None

        if not re.match(r'^[a-zA-Z0-9_-]{25,40}$', file_id):
            logging.error(f"Suspicious file ID format: {file_id}")
            return None

        download_url = f"https://drive.google.com/uc?export=download&id={file_id}"
        parsed_url = urlparse(download_url)
        if parsed_url.netloc != 'drive.google.com':
            logging.error(f"Invalid domain in download URL: {parsed_url.netloc}")
            return None

        headers = {'User-Agent': 'Mozilla/5.0 (compatible; SafeDownloader/1.0)'}
        response = requests.get(download_url, stream=True, headers=headers, timeout=10)
        
        if response.status_code != 200:
            logging.error(f"Failed to download file from {download_url}: Status {response.status_code}")
            return None
        
        content_type = response.headers.get('Content-Type', '')
        if not content_type.startswith('image/'):
            logging.error(f"Unexpected content type: {content_type}")
            return None

        os.makedirs(temp_dir, exist_ok=True)
        temp_file = tempfile.NamedTemporaryFile(
            delete=False, 
            suffix='.png',
            dir=temp_dir
        )
        
        max_size = 10 * 1024 * 1024
        downloaded_size = 0
        for chunk in response.iter_content(chunk_size=8192):
            if chunk:
                downloaded_size += len(chunk)
                if downloaded_size > max_size:
                    temp_file.close()
                    os.unlink(temp_file.name)
                    logging.error(f"File exceeds maximum size limit: {max_size} bytes")
                    return None
                temp_file.write(chunk)
        temp_file.close()
        
        logging.info(f"Downloaded image to {temp_file.name}")
        return temp_file.name

    except requests.exceptions.Timeout:
        logging.error(f"Download timed out for link: {google_drive_link}")
        return None
    except requests.exceptions.RequestException as e:
        logging.error(f"Network error downloading Google Drive image: {e}")
        return None
    except Exception as e:
        logging.error(f"Error downloading Google Drive image: {e}")
        return None