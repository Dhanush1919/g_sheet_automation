# G-sheet automation Sales Data

# Process
- Create a Google Cloud Project: Navigate to the Google Cloud Console and create a new project.
- Enable APIs: Enable the Google Sheets API and Google Drive API for the project.
- Create Service Account: Go to the IAM & Admin section, create a service account, and download the credentials JSON file. This file will allow your Python script to authenticate with Google Sheets.
- Install gspread for interacting with Google Sheets and oauth2client for authentication.
- Authenticating and Connecting to Google Sheets
- Accessing Data from Various Sheets and aggregating data into a master sheet
- Analyzing the data using python libraries
- Exporting the results into a Google Doc file and sharing access to the mentioned mail
