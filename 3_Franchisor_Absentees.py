import openpyxl
import pandas as pd
import requests
import logging
import os
from io import BytesIO
from tkinter import messagebox

# Create log file
logging.basicConfig(filename="log.txt", level=logging.DEBUG, format="%(asctime)s %(message)s", filemode="a")
logging.info("Running 3_Franchisor_Absentees.py...")

# Read from Google Sheets without API
# QCP2 1pzQvT4itxCfYhCjCNZMZYFiKXbvwsTTfWRJGaEg7Gjc / SA1 194PK9WOqA0RUY0XAegI7KblWjLvmACoJE3JjUJn2Ckk
spreadsheetId = "1pzQvT4itxCfYhCjCNZMZYFiKXbvwsTTfWRJGaEg7Gjc" # Spreadsheet ID to change every QCP period, current example for AY2023-SA1
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
except:
    logging.info("Proceed to export data...")
    print("Proceed to export data...")

# Write to XLSX file on local device in working directory
Scanned.to_excel("ABS Scanned Franchisor.xlsx")
print("ABS Scanned Franchisor.xlsx created!")
PendingConversion.to_excel("ABS Pending Conversion Franchisor.xlsx")
print("ABS Pending Conversion Franchisor.xlsx created!")
OutstandingScans.to_excel("ABS Outstanding Scans Franchisor.xlsx")
print("ABS Outstanding Scans Franchisor.xlsx created!")
print("Success!")
logging.info("End of 3_Franchisor_Absentees.py...")