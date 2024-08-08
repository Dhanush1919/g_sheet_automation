import gspread
import pandas as pd
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

# Scopes required to access Google Docs
SCOPES = ['https://www.googleapis.com/auth/documents', 'https://www.googleapis.com/auth/drive']

# Path to your service account credentials JSON file
SERVICE_ACCOUNT_FILE = '/home/nineleaps/Documents/Python_Gsheet_automation/credentials.json'

# Scopes required to access Google Sheets and Google Docs
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/documents'
]

# Authenticate using the service account
creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
client = gspread.authorize(creds)

# Google Docs API client
docs_service = build('docs', 'v1', credentials=creds)

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

# Analysis results
results = []

# 1. TOTAL NUMBER OF RECORDS
total_entries = df.shape[0]
results.append(f"Total Number of Records: {total_entries}\n")

# 2. TOTAL NUMBER OF CUSTOMERS
total_customer_counts = df['CustomerID'].nunique()
results.append(f"Total Number of Customers: {total_customer_counts}\n")

# 3. TOTAL NUMBER OF PRODUCTS
total_product_counts = df['ProductName'].nunique()
results.append(f"Total Number of Products: {total_product_counts}\n")

# 4. EACH PRODUCT WISE TOTAL ENTRIES
product_wise_entries = df['ProductName'].value_counts()
results.append("Each Product Wise Counts:\n")
results.append(product_wise_entries.to_string() + "\n\n")

# 5. PAYMENT OPTIONS AVAILABLE
available_payment_options = df['PaymentMethod'].nunique()
results.append(f"Available Payment Options: {available_payment_options}\n")

# 6. EACH PAYMENT OPTIONS UTILISED COUNT
each_payment_options_used = df['PaymentMethod'].value_counts()
results.append("Each Payment Option Utilized Count:\n")
results.append(each_payment_options_used.to_string() + "\n\n")

# 7. SUM OF TOTAL SALES
total_sales = df['TotalPrice'].sum()
results.append(f"Sum of Total Sales: {total_sales}\n")

# 8. TOTAL QUANTITIES SOLD
total_quantities_sold = df['Quantity'].sum()
results.append(f"Total Quantities Sold: {total_quantities_sold}\n")

# 9. DAY WISE TOTAL SALES
df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
df = df.dropna(subset=['Date'])
df['Date'] = df['Date'].dt.date
day_wise_total_price = df.groupby('Date')['TotalPrice'].sum().reset_index()
day_wise_total_price.columns = ['Date', 'TotalPrice']
results.append("Day Wise Total Price:\n")
results.append(day_wise_total_price.to_string(index=False) + "\n\n")

# 10. DAY WITH THE HIGHEST AND LEAST SALE
day_with_highest_sale = day_wise_total_price.loc[day_wise_total_price['TotalPrice'].idxmax()]
day_with_least_sale = day_wise_total_price.loc[day_wise_total_price['TotalPrice'].idxmin()]
results.append(f"Day with the Highest Sale:\n{day_with_highest_sale}\n\n")
results.append(f"Day with the Least Sale:\n{day_with_least_sale}\n")

# Create a new Google Doc and write the analysis results
doc = docs_service.documents().create(body={"title": "Data Analysis Summary"}).execute()
doc_id = doc['documentId']
print(f"Created document with ID: {doc_id}")

# Prepare the content to be inserted into the Google Doc
requests = [
    {
        'insertText': {
            'location': {
                'index': 1,
            },
            'text': ''.join(results)
        }
    }
]

# Execute the request to write the content to the Google Doc
docs_service.documents().batchUpdate(documentId=doc_id, body={'requests': requests}).execute()

print(f"Analysis complete and results exported to Google Doc: {doc['title']}")

print("Document updated successfully.")