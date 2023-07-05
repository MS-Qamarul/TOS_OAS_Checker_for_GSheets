import openpyxl
import pandas as pd
import requests
import os
from io import BytesIO

# Read from Google Sheets without API
spreadsheetId = "1Eob55-5hqyjRB49iJQe-1xuZtWTcu9TKm3ivkpSd_Vw" # Spreadsheet ID to change every QCP period
url = "https://docs.google.com/spreadsheets/export?exportFormat=xlsx&id=" + spreadsheetId
res = requests.get(url)
data = BytesIO(res.content)
xlsx = openpyxl.load_workbook(filename=data)

df = pd.concat(pd.read_excel(data, sheet_name=None), ignore_index=True)
print("Successfully read data from sheet...")

# Dataframe queries
Scanned =  df.query('Scanned == True & Teacher == Teacher')
PendingConversion = df.query('Scanned == True & Teacher == Teacher & `Date of IT conversion to Verificare` != `Date of IT conversion to Verificare`')
OutstandingScans = df.query('Scanned == False & Teacher == Teacher')
ConvertedOAS = df.query('Scanned == True & Teacher == Teacher & `Date of IT conversion to Verificare` == `Date of IT conversion to Verificare`')

# Check if a previous version of XLSX files exists
try:
    # Scanned
    os.path.exists("./ABS Scanned Franchisee.xlsx")
    os.remove("./ABS Scanned Franchisee.xlsx")
    print("A previous version of ABS Scanned Franchisee.xlsx detected! Deleting file...")
    # Pending Conversion
    os.path.exists("./ABS Pending Conversion Franchisee.xlsx")
    os.remove("./ABS Pending Conversion Franchisee.xlsx")
    print("A previous version of ABS Pending Conversion Franchisee.xlsx detected! Deleting file...")
    # Outstanding Scans
    os.path.exists("./ABS Outstanding Scans Franchisee.xlsx")
    os.remove("./ABS Outstanding Scans Franchisee.xlsx")
    print("A previous version of ABS Outstanding Scans Franchisee.xlsx detected! Deleting file...")
    # Converted OAS
    os.path.exists("./ABS Converted OAS Franchisee.xlsx")
    os.remove("./ABS Converted OAS Franchisee.xlsx")
    print("A previous version of ABS Converted OAS Franchisee.xlsx detected! Deleting file...")
except:
    print("Proceed to export data...")

# Write to XLSX file on local device in working directory
Scanned.to_excel("ABS Scanned Franchisee.xlsx")
print("ABS Scanned Franchisee.xlsx created!")
PendingConversion.to_excel("ABS Pending Conversion Franchisee.xlsx")
print("ABS Pending Conversion Franchisee.xlsx created!")
OutstandingScans.to_excel("ABS Outstanding Scans Franchisee.xlsx")
print("ABS Outstanding Scans Franchisee.xlsx created!")
ConvertedOAS.to_excel("ABS Converted OAS Franchisee.xlsx")
print("ABS Converted OAS Franchisee.xlsx created!")