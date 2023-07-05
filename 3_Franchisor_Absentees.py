import openpyxl
import pandas as pd
import requests
import os
from io import BytesIO

# Read from Google Sheets without API
spreadsheetId = "1pzQvT4itxCfYhCjCNZMZYFiKXbvwsTTfWRJGaEg7Gjc" # Spreadsheet ID to change every QCP period
url = "https://docs.google.com/spreadsheets/export?exportFormat=xlsx&id=" + spreadsheetId
res = requests.get(url)
data = BytesIO(res.content)
xlsx = openpyxl.load_workbook(filename=data)

df = pd.concat(pd.read_excel(data, sheet_name=None), ignore_index=True)# pandas count distinct values in column
print("Successfully read data from sheet...")

# Dataframe queries
Scanned =  df.query('Scanned == True & Teacher == Teacher')
PendingConversion = df.query('Scanned == True & Teacher == Teacher & `Date of IT conversion to Verificare` != `Date of IT conversion to Verificare`')
OutstandingScans = df.query('Scanned == False & Teacher == Teacher')
ConvertedOAS = df.query('Scanned == True & Teacher == Teacher & `Date of IT conversion to Verificare` == `Date of IT conversion to Verificare`')

# Check if a previous version of XLSX files exists
try:
    # Scanned
    os.path.exists("./ABS Scanned Franchisor.xlsx")
    os.remove("./ABS Scanned Franchisor.xlsx")
    print("A previous version of ABS Scanned Franchisor.xlsx detected! Deleting file...")
    # Pending Conversion
    os.path.exists("./ABS Pending Conversion Franchisor.xlsx")
    os.remove("./ABS Pending Conversion Franchisor.xlsx")
    print("A previous version of ABS Pending Conversion Franchisor.xlsx detected! Deleting file...")
    # Outstanding Scans
    os.path.exists("./ABS Outstanding Scans Franchisor.xlsx")
    os.remove("./ABS Outstanding Scans Franchisor.xlsx")
    print("A previous version of ABS Outstanding Scans Franchisor.xlsx detected! Deleting file...")
    # Converted OAS
    os.path.exists("./ABS Converted OAS Franchisor.xlsx")
    os.remove("./ABS Converted OAS Franchisor.xlsx")
    print("A previous version of ABS Converted OAS Franchisor.xlsx detected! Deleting file...")
except:
    print("Proceed to export data...")

# Write to XLSX file on local device in working directory
Scanned.to_excel("ABS Scanned Franchisor.xlsx")
print("ABS Scanned Franchisor.xlsx created!")
PendingConversion.to_excel("ABS Pending Conversion Franchisor.xlsx")
print("ABS Pending Conversion Franchisor.xlsx created!")
OutstandingScans.to_excel("ABS Outstanding Scans Franchisor.xlsx")
print("ABS Outstanding Scans Franchisor.xlsx created!")
ConvertedOAS.to_excel("ABS Converted OAS Franchisor.xlsx")
print("ABS Converted OAS Franchisor.xlsx created!")