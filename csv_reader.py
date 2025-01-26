import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta

scopes = ["https://www.googleapis.com/auth/spreadsheets"]
creds = Credentials.from_service_account_file("credentials.json", scopes=scopes)
client = gspread.authorize(creds)

sheet_id = "1wyePQ834656UjFPFt9WpSAoOhXQ32YJ3Z90yvoGckSw"
workbook = client.open_by_key(sheet_id)

sheet = workbook.sheet1

# Fetch all the data from the sheet and convert it to a pandas DataFrame
data = sheet.get_all_values()
df = pd.DataFrame(data[1:], columns=data[0])  # Assumes first row is headers

# Convert 'Timestamp' column to datetime
df['Timestamp'] = pd.to_datetime(df['Timestamp'], errors='coerce')
df['Timestamp'] = df['Timestamp'].where(pd.notna(df['Timestamp']), None)

# Get today's date and subtract 13 years
today = datetime.today()
years_ago_13 = today - relativedelta(years=13)

# Add the 'Alum' column based on the condition
df['Alum'] = df['Timestamp'].apply(lambda x: 'Yes' if x < years_ago_13 else 'No')

# Get the index of the "Alum" column in the DataFrame
alum_column_index = df.columns.get_loc('Alum') + 1  # +1 because gspread uses 1-based indexing

# Prepare the list of alum values to update
alum_values = df['Alum'].tolist()

# Update only the "Alum" column in the Google Sheet
for i, alum in enumerate(alum_values, start=2):  # start=2 to skip the header row
    sheet.update_cell(i, alum_column_index, alum)

print("Alum column updated successfully!")
