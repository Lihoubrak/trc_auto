from fuzzywuzzy import fuzz
from config import logging, SIMILARITY_THRESHOLD
from openpyxl import Workbook


from fuzzywuzzy import fuzz
from config import logging, SIMILARITY_THRESHOLD
from openpyxl import Workbook


def match_headers(excel_headers, form_headers, output_path=None):
    """Match Excel headers to form headers using fuzzywuzzy and optionally save results to Excel."""
    matched_headers, mapping, unmatched = [], {}, []
    for excel_header in excel_headers:
        best_match, best_score = None, 0
        for form_header in form_headers:
            score = fuzz.token_sort_ratio(excel_header.lower(), form_header.lower())
            logging.debug(f"Header match: '{excel_header}' vs '{form_header}' = {score}")
            if score > best_score and score >= SIMILARITY_THRESHOLD:
                best_match, best_score = form_header, score
        if best_match and best_match not in mapping.values():
            matched_headers.append(excel_header)
            mapping[excel_header] = best_match
            logging.info(f"Matched '{excel_header}' to '{best_match}' (score: {best_score})")
        else:
            unmatched.append(excel_header)
            logging.warning(f"No match for Excel header '{excel_header}'")

    # Save to Excel if output_path is provided
    if output_path:
        wb = Workbook()

        # Matched Headers sheet
        ws_matched = wb.active
        ws_matched.title = "Matched Headers"
        ws_matched.append(["Excel Header", "Matched Form Header"])
        for excel_header in matched_headers:
            form_header = mapping.get(excel_header, "")
            ws_matched.append([excel_header, form_header])

        # Unmatched Headers sheet
        ws_unmatched = wb.create_sheet("Unmatched Headers")
        ws_unmatched.append(["Unmatched Excel Header"])
        for header in unmatched:
            ws_unmatched.append([header])

        wb.save(output_path)
        logging.info(f"Header match results saved to {output_path}")

    return matched_headers, unmatched, mapping



def fuzzy_match_value(value_str, options):
    """Match a value to a list of options using fuzzywuzzy."""
    best_option, best_score = None, 0
    for opt, opt_text in options:
        score = fuzz.token_sort_ratio(value_str.lower(), opt_text.lower())
        logging.debug(f"Value match: '{value_str}' vs '{opt_text}' = {score}")
        if score > best_score and score >= SIMILARITY_THRESHOLD:
            best_option, best_score = opt, score
    return best_option, best_score

