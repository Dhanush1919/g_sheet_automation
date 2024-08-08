import gspread
from google.oauth2.service_account import Credentials

# Path to your service account credentials JSON file
SERVICE_ACCOUNT_FILE = '/home/nineleaps/Documents/Python_Gsheet_automation/credentials.json'

# Scopes required to access Google Sheets
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

# Authenticate using the service account
creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
client = gspread.authorize(creds)

# ID of the spreadsheet containing the subsheets
spreadsheet_id = '1E7tZ6MuhiA4bFfGZLCcCemdoslqLTicFlvTmRQVDowM'  # Replace with your spreadsheet ID

# Names of the subsheets to merge
sheet_names = [
    'Salesdata1',  
    'Salesdata2',  
    'Salesdata3'  
]

# Name of the destination sheet (within the same spreadsheet)
destination_sheet_name = 'MergedSheet'  # Replace with your desired destination sheet name

# Access the spreadsheet
spreadsheet = client.open_by_key(spreadsheet_id)

# Check if the destination sheet already exists; if not, create it
try:
    destination_sheet = spreadsheet.worksheet(destination_sheet_name)
    destination_sheet.clear()  # Clear existing content if any
except gspread.exceptions.WorksheetNotFound:
    destination_sheet = spreadsheet.add_worksheet(title=destination_sheet_name, rows="1000", cols="26")

# Track the row where we will start pasting data
start_row = 1

# Loop through each subsheet
for sheet_name in sheet_names:
    source_sheet = spreadsheet.worksheet(sheet_name)
    source_data = source_sheet.get_all_values()
    
    # Check if the destination sheet needs more rows
    required_rows = start_row + len(source_data) - 1
    current_max_rows = destination_sheet.row_count
    
    if required_rows > current_max_rows:
        destination_sheet.resize(rows=required_rows)
    
    # Determine the range in the destination sheet to paste the data
    end_row = start_row + len(source_data) - 1
    cell_range = f'A{start_row}'  # Adjust this based on your data width
    
    # Update the data in the destination sheet
    destination_sheet.update(cell_range, source_data)
    
    # Update the start_row for the next subsheet's data
    start_row = end_row + 1

print('Data successfully merged into the destination sheet!')