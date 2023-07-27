import openpyxl
import pandas as pd
import requests
import os
from io import BytesIO

print("========== STEP 5: Process Online Class OAS data ==========")

# Read from Google Sheets without API
spreadsheetId = "1XQxHZ6wmp2lWJwvWFy6SXolb60yxaE4Bob7DO6O3P-Y" # Spreadsheet ID to change every QCP period
sheetId = "764556405"
url = "https://docs.google.com/spreadsheets/d/" + spreadsheetId + "/export?format=xlsx&gid=" + sheetId
res = requests.get(url)
data = BytesIO(res.content)
xlsx = openpyxl.load_workbook(filename=data)

extension = ".xlsx"

df = pd.concat(pd.read_excel(data, sheet_name=None, header=None, names=['Class: Class Code', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S']), ignore_index=True)

# Common conditions for DataFrame queries
not_scanned_condition = 'I == False & H == H & G == G'
scanned_condition = 'I == True & H == H & G == G'
not_converted_condition = 'N == N & O != O'
converted_condition = 'N == N & O == O'

# Dataframe queries
PendingConversion = df.query(f'{scanned_condition} & {not_converted_condition}')
OutstandingScans = df.query(f'{not_scanned_condition}')
ScannedOAS = df.query(f'{scanned_condition}')
ConvertedOAS = df.query(f'{scanned_condition} & {converted_condition}')

# Write to XLSX file on local device in working directory
PendingConversion.to_excel("Pending_Conversion_Online_Class" + extension, index=False)
print("CREATE: Pending_Conversion_Online_Class" + extension)

OutstandingScans.to_excel("Outstanding_Scans_Online_Class" + extension, index=False)
print("CREATE: Outstanding_Scans_Online_Class" + extension)

ScannedOAS.to_excel("Scanned_OAS_Online_Class" + extension, index=False)
print("CREATE: Scanned_OAS_Online_Class" + extension)

ConvertedOAS.to_excel("Converted_OAS_Online_Class" + extension, index=False)
print("CREATE: Converted_OAS_Online_Class" + extension)