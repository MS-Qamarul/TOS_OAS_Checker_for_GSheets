import openpyxl
import pandas as pd
import requests
import os
from io import BytesIO

# Read from Google Sheets without API
spreadsheetId = "1XQxHZ6wmp2lWJwvWFy6SXolb60yxaE4Bob7DO6O3P-Y" # Spreadsheet ID to change every QCP period
sheetId = "764556405"
url = "https://docs.google.com/spreadsheets/d/" + spreadsheetId + "/export?format=xlsx&gid=" + sheetId
res = requests.get(url)
data = BytesIO(res.content)
xlsx = openpyxl.load_workbook(filename=data)

df = pd.concat(pd.read_excel(data, sheet_name=None), ignore_index=False)
print("Successfully read data from sheet...")

# Dataframe queries
PendingConversion = df.query('`OAS/ MCQ Scan in Folder (Please check only after u have checked that there is OAS or MCQ answers on the paper in the scan file in the class QCP folder, to prevent false submission records)` == True & `QCP2/ Prelim OAS Week` == `QCP2/ Prelim OAS Week` & `Student Name` == `Student Name` & `PIC for OAS converted to Googlesheet` == `PIC for OAS converted to Googlesheet` & `Date OAS converted to googlesheet` != `Date OAS converted to googlesheet`')
OutstandingScans = df.query('`OAS/ MCQ Scan in Folder (Please check only after u have checked that there is OAS or MCQ answers on the paper in the scan file in the class QCP folder, to prevent false submission records)` == False & `QCP2/ Prelim OAS Week` == `QCP2/ Prelim OAS Week` & `Student Name` == `Student Name` & `PIC for OAS converted to Googlesheet` == `PIC for OAS converted to Googlesheet`')
ScannedOAS = df.query('`OAS/ MCQ Scan in Folder (Please check only after u have checked that there is OAS or MCQ answers on the paper in the scan file in the class QCP folder, to prevent false submission records)` == True & `QCP2/ Prelim OAS Week` == `QCP2/ Prelim OAS Week` & `Student Name` == `Student Name` & `PIC for OAS converted to Googlesheet` == `PIC for OAS converted to Googlesheet`')
ConvertedOAS = df.query('`OAS/ MCQ Scan in Folder (Please check only after u have checked that there is OAS or MCQ answers on the paper in the scan file in the class QCP folder, to prevent false submission records)` == True & `QCP2/ Prelim OAS Week` == `QCP2/ Prelim OAS Week` & `Student Name` == `Student Name` & `PIC for OAS converted to Googlesheet` == `PIC for OAS converted to Googlesheet` & `Date OAS converted to googlesheet` == `Date OAS converted to googlesheet`')

# Check if a previous version of XLSX files exists
try:
    # Pending Conversion
    os.path.exists("./PO Pending Conversion.xlsx")
    os.remove("./PO Pending Conversion.xlsx")
    print("A previous version of PO Pending Conversion.xlsx detected! Deleting file...")
    # Outstanding Scans
    os.path.exists("./PO Outstanding Scans.xlsx")
    os.remove("./PO Outstanding Scans.xlsx")
    print("A previous version of PO Outstanding Scans.xlsx detected! Deleting file...")
    # Scanned OAS
    os.path.exists("./PO Scanned OAS.xlsx")
    os.remove("./PO Scanned OAS.xlsx")
    print("A previous version of PO Scanned OAS.xlsx detected! Deleting file...")
except:
    print("Proceed to export data...")

# Write to XLSX file on local device in working directory
PendingConversion.to_excel("PO Pending Conversion.xlsx")
print("PO Pending Conversion.xlsx created!")
OutstandingScans.to_excel("PO Outstanding Scans.xlsx")
print("PO Outstanding Scans.xlsx created!")
ScannedOAS.to_excel("PO Scanned OAS.xlsx")
print("PO Scanned OAS.xlsx created!")
ConvertedOAS.to_excel("PO Converted OAS.xlsx")
print("PO Converted OAS.xlsx created!")
print("Success!")