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

# 1. TOTAL NUMBER OF RECORDS : 
total_entries = df.shape
print("Total Number of Records : ",total_entries[0])
print("\n")

# 2. TOTAL NUMBER OF CUSTOMERS : 
total_customer_counts = df['CustomerID'].nunique()
print("Total No. of Customer : ",total_customer_counts)
print("\n")

# 3. TOTAL NUMBER OF PRODUCTS :
total_product_counts = df['ProductName'].nunique()
print("Total No. of Products : ",total_product_counts)
print("\n")

# 4. EACH PRODUCT WISE TOTAL ENTRIES :
product_wise_entries = df['ProductName'].value_counts()
print("EACH PRODUCT WISE COUNTS : \n ", product_wise_entries)
print("\n")

# 5. PAYMENT OPTIONS AVAILABLE :
available_payment_options = df['PaymentMethod'].nunique()
print("Available Payment options : ",available_payment_options)
print("\n")

# 6. EACH PAYMENT OPTIONS UTILISED COUNT : 
each_payment_options_used = df['PaymentMethod'].value_counts()
print("Each Payment Options utilised count : \n", each_payment_options_used)
print("\n")

# 7. SUM OF TOTAL SALES :
total_sales = df['TotalPrice'].sum()
print("Sum of Total sales : ", total_sales)
print("\n")

# 8. TOTAL QUANTITIES SOLD : 
total_quantities_sold = df['Quantity'].sum()
print("Total Quantities sold :",total_quantities_sold)
print("\n")

# 9. DAY WISE TOTAL SALES :
df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
df = df.dropna(subset=['Date'])
df['Date'] = df['Date'].dt.date
day_wise_total_price = df.groupby('Date')['TotalPrice'].sum().reset_index()
day_wise_total_price.columns = ['Date', 'TotalPrice']
print("Day wise total Price : \n")
print(day_wise_total_price)

# 10. DAY WITH THE HIGHEST AND LEAST SALE :
day_with_highest_sale = day_wise_total_price.loc[day_wise_total_price['TotalPrice'].idxmax()]

day_with_least_sale = day_wise_total_price.loc[day_wise_total_price['TotalPrice'].idxmin()]

print("Day with the highest sale with sale amount :\n", day_with_highest_sale)
print("\n")
print("Day with the least sale with sale amount :\n", day_with_least_sale)
