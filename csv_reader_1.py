import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import pandas as pd

# Setup Google Sheets API client
scopes = ["https://www.googleapis.com/auth/spreadsheets"]
creds = Credentials.from_service_account_file("credentials.json", scopes=scopes)
client = gspread.authorize(creds)

# Open the Google Sheet
sheet_id = "1wyePQ834656UjFPFt9WpSAoOhXQ32YJ3Z90yvoGckSw"
workbook = client.open_by_key(sheet_id)
sheet = workbook.sheet1

# Expected header format
expected_header = ['Timestamp', 'Name', 'Email', 'Address', 'Number', 'Alum']

# Get the current header row
current_header = sheet.row_values(1)

# Check if headers match and update if needed
if current_header != expected_header:
    print("Headers don't match. Clearing the sheet and updating headers.")
    sheet.batch_clear(['A1:Z1', 'G:Z'])  # Clear the first row (headers)
    sheet.update('A1:F1', [expected_header])  # Set the expected header in the first row
else:
    print("Headers match. No action needed.")

# Step to clear columns with empty headers
for col_index, header in enumerate(current_header, start=1):
    if not header.strip():  # Check if the header is empty
        # Get the column letter (e.g., A, B, C)
        column_letter = gspread.utils.rowcol_to_a1(1, col_index).split('1')[0]
        # Get number of rows in the sheet
        column_values = sheet.col_values(col_index)
        num_rows = len(column_values)
        # Clear the entire column
        sheet.update(f"{column_letter}1:{column_letter}{num_rows}", [[""] for _ in range(num_rows)])

print("Completed clearing columns with empty headers.")

# Function to parse timestamp into datetime
def parse_timestamp(record):
    """Parses the 'Timestamp' field of a record into a datetime object."""
    try:
        return datetime.strptime(record['Timestamp'], "%Y-%m-%d %H:%M:%S")
    except KeyError:
        raise ValueError("The record does not contain a 'Timestamp' field.")
    except ValueError:
        raise ValueError(f"Invalid timestamp format in record: {record['Timestamp']}")

# Function to sort the data based on the 'Timestamp'
def sort_timestamp():
    # Fetch all data as a list of dictionaries
    data = sheet.get_all_records()

    # Ensure that all records have a valid 'Timestamp' field and format
    for record in data:
        if 'Timestamp' not in record:
            print(f"Missing 'Timestamp' in record: {record}")
            return
        try:
            datetime.strptime(record['Timestamp'], "%Y-%m-%d %H:%M:%S")
        except ValueError:
            print(f"Invalid timestamp format in record: {record['Timestamp']}")
            return

    # Sort the data using the parse_timestamp function
    sorted_data = sorted(data, key=parse_timestamp)
    
    # Optionally, you can print or process the sorted data
    for row in sorted_data:
        print(row)

    # If you want to update the sheet with sorted data:
    # Clear the existing data (keeping the headers)
    sheet.batch_clear(["A2:F100000"])  # You can adjust the range as needed
    # Write the sorted data back to the sheet
    sheet.append_rows(sorted_data, value_input_option='RAW')

# Run the sorting function
sort_timestamp()

