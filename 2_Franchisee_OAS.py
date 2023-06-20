import openpyxl
import pandas as pd
import requests
import logging
import os
from io import BytesIO
from tkinter import messagebox

# Create log file
logging.basicConfig(filename="log.txt", level=logging.DEBUG, format="%(asctime)s %(message)s", filemode="a")
logging.info("Running 2_Franchisee_OAS.py...")

# Read from Google Sheets without API
# QCP2 1DKwHmmEJm6BWEJ5fsZ0KSOLHoQG6yDxJBmMAOpugxgY / SA1 1j_z-6aAELodrDe0S08hViss9h9Rp0xEuUefGsnjc2nM
spreadsheetId = "1DKwHmmEJm6BWEJ5fsZ0KSOLHoQG6yDxJBmMAOpugxgY" # Spreadsheet ID to change every QCP period, current example for AY2023-SA1
url = "https://docs.google.com/spreadsheets/export?exportFormat=xlsx&id=" + spreadsheetId
res = requests.get(url)
data = BytesIO(res.content)
xlsx = openpyxl.load_workbook(filename=data)

df = pd.concat(pd.read_excel(data, sheet_name=None), ignore_index=True)
print("Successfully read data from sheet...")

# Dataframe queries
PendingConversion = df.query('Scanned == True & Teacher == Teacher & `Date of IT conversion to Verificare` != `Date of IT conversion to Verificare`')
OutstandingScans = df.query('Scanned == False & Teacher == Teacher')
ScannedOAS = df.query('Scanned == True & Teacher == Teacher')
ConvertedOAS = df.query('Scanned == True & Teacher == Teacher & `Date of IT conversion to Verificare` == `Date of IT conversion to Verificare`')

# Check if a previous version of XLSX files exists
try:
    # Pending Conversion
    os.path.exists("./Pending Conversion Franchisee.xlsx")
    os.remove("./Pending Conversion Franchisee.xlsx")
    print("A previous version of Pending Conversion Franchisee.xlsx detected! Deleting file...")
    # Outstanding Scans
    os.path.exists("./Outstanding Scans Franchisee.xlsx")
    os.remove("./Outstanding Scans Franchisee.xlsx")
    print("A previous version of Outstanding Scans Franchisee.xlsx detected! Deleting file...")
    # Scanned OAS
    os.path.exists("./Scanned OAS Franchisee.xlsx")
    os.remove("./Scanned OAS Franchisee.xlsx")
    print("A previous version of Scanned OAS Franchisee.xlsx detected! Deleting file...")
    # Converted OAS
    os.path.exists("./Converted OAS Franchisee.xlsx")
    os.remove("./Converted OAS Franchisee.xlsx")
    print("A previous version of Converted OAS Franchisee.xlsx detected! Deleting file...")
except:
    logging.info("Proceed to export data...")
    print("Proceed to export data...")

# Write to XLSX file on local device in working directory
PendingConversion.to_excel("Pending Conversion Franchisee.xlsx")
print("Pending Conversion Franchisee.xlsx created!")
OutstandingScans.to_excel("Outstanding Scans Franchisee.xlsx")
print("Franchisee Outstanding Scnas.xlsx created!")
ScannedOAS.to_excel("Scanned OAS Franchisee.xlsx")
print("Scanned OAS Franchisee.xlsx created!")
ConvertedOAS.to_excel("Converted OAS Franchisee.xlsx")
print("Converted OAS Franchisee.xlsx created!")
logging.info("End of 2_Franchisee_OAS.py...")