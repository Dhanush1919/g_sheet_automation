import gspread
import pandas as pd
from google.oauth2.service_account import Credentials

# Path to your service account credentials JSON file
SERVICE_ACCOUNT_FILE = '/home/nineleaps/Documents/Python_Gsheet_automation/credentials.json'

# Scopes required to access Google Sheets
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

# Authenticate using the service account
creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
client = gspread.authorize(creds)

# ID of the spreadsheet containing the merged data
spreadsheet_id = '1E7tZ6MuhiA4bFfGZLCcCemdoslqLTicFlvTmRQVDowM'
merged_sheet_name = 'MergedSheet'

# Access the merged sheet
spreadsheet = client.open_by_key(spreadsheet_id)
sheet = spreadsheet.worksheet(merged_sheet_name)

# Load the data into a pandas DataFrame
data = sheet.get_all_values()
headers = data.pop(0)  # Assuming the first row contains column headers
df = pd.DataFrame(data, columns=headers)

# Convert relevant columns to numeric types (e.g., Quantity, UnitPrice, TotalPrice)
df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce')
df['UnitPrice'] = pd.to_numeric(df['UnitPrice'], errors='coerce')
df['TotalPrice'] = pd.to_numeric(df['TotalPrice'], errors='coerce')

# Basic Analysis Examples:

# 1. Summary Statistics
summary_stats = df.describe()
print("Summary Statistics:\n", summary_stats)
print("\n")

# 2. Count of unique values in a specific column 
unique_counts = df['ProductName'].value_counts()
print("Unique ProductName Counts:\n", unique_counts)
print("\n")

# 3. Grouping and Aggregation
grouped_data = df.groupby('ProductName')['Quantity'].sum()
print("Grouped Data by ProductName:\n", grouped_data)
print("\n")

# 4. Filtering Data
filtered_data = df[df['UnitPrice'] > 100]
print("Filtered Data where UnitPrice > 100:\n", filtered_data)
print("\n")

# 5. Correlation Matrix
numeric_df = df[['Quantity', 'UnitPrice', 'TotalPrice']] 
correlation_matrix = numeric_df.corr()
print("Correlation Matrix:\n", correlation_matrix)
print("\n")

# 6. Custom Analysis (e.g., calculating a new metric)
df['Revenue'] = df['UnitPrice'] * df['Quantity']
print("Data with Revenue Calculated:\n", df[['ProductName', 'UnitPrice', 'Quantity', 'Revenue']])
print("\n")

# 7. Export the analysis back to Google Sheets (optional)
summary_sheet = spreadsheet.add_worksheet(title="SummaryStats_2", rows="100", cols="20")
summary_sheet.update([summary_stats.columns.values.tolist()] + summary_stats.values.tolist())

print("Analysis complete and results exported to Google Sheets!")


