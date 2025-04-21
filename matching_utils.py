import logging
import re
from fuzzywuzzy import fuzz


def normalize(text):
    """Normalize the text while preserving line breaks and other special characters."""
    # This removes only excessive whitespace between words but keeps special characters intact.
    return re.sub(r'\s+', ' ', text).strip()

def match_headers(excel_headers, form_headers):
    """Match Excel headers to form headers using fuzzy matching."""
    mapping = {}
    unmatched = []
    SIMILARITY_THRESHOLD = 80
    for excel_header in excel_headers:
        excel_clean = normalize(excel_header)
        best_match, best_score = None, 0

        for form_header in form_headers:
            form_clean = normalize(form_header)
            score = fuzz.ratio(excel_clean, form_clean)

            if score > best_score and score >= SIMILARITY_THRESHOLD:
                best_match, best_score = form_header, score

        if best_match:
            mapping[excel_header] = best_match
            logging.info(f"Matched '{excel_header}' to '{best_match}' (score: {best_score})")
        else:
            unmatched.append(excel_header)
            logging.warning(f"No match for Excel header: '{excel_header}'")

    if unmatched:
        logging.warning(f"Unmatched headers: {unmatched}")
    return mapping, unmatched
