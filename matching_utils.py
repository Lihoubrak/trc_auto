import logging
from fuzzywuzzy import fuzz
from config import SIMILARITY_THRESHOLD

def match_headers(excel_headers, form_headers):
    """Match Excel headers to form headers using fuzzy matching."""
    mapping = {}
    unmatched = []
    for excel_header in excel_headers:
        best_match, best_score = None, 0
        for form_header in form_headers:
            score = fuzz.token_sort_ratio(excel_header, form_header)
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