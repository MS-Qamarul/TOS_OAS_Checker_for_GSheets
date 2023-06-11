import openpyxl
import pandas as pd
import requests
import logging
import os
from io import BytesIO
from tkinter import messagebox

# Create log file
logging.basicConfig(filename="log.txt", level=logging.DEBUG, format="%(asctime)s %(message)s", filemode="a")
logging.info("Running 1_Franchisor_OAS.py...")

# Read from Google Sheets without API
spreadsheetId = "1E2qRZq_BXSLLw8-CAFW5mbL233kx1P_JAPb2cC9ZvZE" # Spreadsheet ID to change every QCP period, current example for AY2023-SA1
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

# Check if a previous version of XLSX files exists
try:
    # Pending Conversion
    os.path.exists("./Pending Conversion Franchisor.xlsx")
    os.remove("./Pending Conversion Franchisor.xlsx")
    print("A previous version of Pending Conversion Franchisor.xlsx detected! Deleting file...")
    # Outstanding Scans
    os.path.exists("./Outstanding Scans Franchisor.xlsx")
    os.remove("./Outstanding Scans Franchisor.xlsx")
    print("A previous version of Outstanding Scans Franchisor.xlsx detected! Deleting file...")
    # Scanned OAS
    os.path.exists("./Scanned OAS Franchisor.xlsx")
    os.remove("./Scanned OAS Franchisor.xlsx")
    print("A previous version of Scanned OAS Franchisor.xlsx detected! Deleting file...")
except:
    logging.info("Proceed to export data...")
    print("Proceed to export data...")

# Write to XLSX file on local device in working directory
PendingConversion.to_excel("Pending Conversion Franchisor.xlsx")
print("Pending Conversion Franchisor.xlsx created!")
OutstandingScans.to_excel("Outstanding Scans Franchisor.xlsx")
print("Outstanding Scans Franchisor.xlsx created!")
ScannedOAS.to_excel("Scanned OAS Franchisor.xlsx")
print("Scanned OAS Franchisor.xlsx created!")
print("Success!")
logging.info("End of 1_Franchisor_OAS.py...")