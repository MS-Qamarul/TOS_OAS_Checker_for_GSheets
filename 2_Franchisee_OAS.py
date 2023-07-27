import openpyxl
import pandas as pd
import requests
from io import BytesIO

print("========== STEP 2: Process Franchisee OAS data ==========")

# Read from Google Sheets without API
spreadsheetId = "1DKwHmmEJm6BWEJ5fsZ0KSOLHoQG6yDxJBmMAOpugxgY" # Spreadsheet ID to change every QCP period
url = "https://docs.google.com/spreadsheets/export?exportFormat=xlsx&id=" + spreadsheetId
res = requests.get(url)
data = BytesIO(res.content)
xlsx = openpyxl.load_workbook(filename=data)

extension = ".xlsx"

df = pd.concat(pd.read_excel(data, sheet_name=None), ignore_index=True)

# Common conditions for DataFrame queries
not_scanned_condition = 'Scanned == False & Teacher == Teacher'
scanned_condition = 'Scanned == True & Teacher == Teacher'
not_converted_condition = '`Date of IT conversion to Verificare` != `Date of IT conversion to Verificare`'
converted_condition = '`Date of IT conversion to Verificare` == `Date of IT conversion to Verificare`'

# Dataframe queries
PendingConversion = df.query(f'{scanned_condition} & {not_converted_condition}')
OutstandingScans = df.query(f'{not_scanned_condition}')
ScannedOAS = df.query(f'{scanned_condition}')
ConvertedOAS = df.query(f'{scanned_condition} & {converted_condition}')

# Write to XLSX file on local device in working directory
PendingConversion.to_excel("Pending_Conversion_Franchisee" + extension, index=False)
print("CREATE: Pending_Conversion_Franchisee" + extension)

OutstandingScans.to_excel("Outstanding_Scans_Franchisee" + extension, index=False)
print("CREATE: Outstanding_Scans_Franchisee" + extension)

ScannedOAS.to_excel("Scanned_OAS_Franchisee" + extension, index=False)
print("CREATE: Scanned_OAS_Franchisee" + extension)

ConvertedOAS.to_excel("Converted_OAS_Franchisee" + extension, index=False)
print("CREATE: Converted_OAS_Franchisee" + extension)