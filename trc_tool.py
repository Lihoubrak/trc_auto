import openpyxl
import requests

# Load your Excel file (change the path to your file)
excel_file = 'MAINTENANCE CABLE REQUEST TO VTC.xlsx'
workbook = openpyxl.load_workbook(excel_file)
sheet = workbook.active

# Google Form URL (replace with your Google Form's action URL)
form_url = "https://docs.google.com/forms/d/e/1FAIpQLSeqWvnn4KIru5BYd6aNVTCvaej6KvPWdbK0tN3piOgU8u8ftg/formResponse"

# Iterate through each row of data in the Excel file (skip the header row)
for row in sheet.iter_rows(min_row=2, values_only=True):  # Skipping the header row
    # Map the form field IDs to the Excel data
    payload = {
        'entry.1234567890': row[0],  # Replace with actual form field ID and Excel data column
        'entry.2345678901': row[1],  # Replace with actual form field ID and Excel data column
        'entry.3456789012': row[2],  # Replace with actual form field ID and Excel data column
        # Add more fields if necessary
    }
    # Send the data to Google Form
    # response = requests.post(form_url, data=payload)

    # Check the response to make sure the submission was successful
    # if response.status_code == 200:
    #     print(f"Successfully submitted data for: {row}")
    # else:
    #     print(f"Failed to submit data for: {row}")
