import pandas as pd
import os
import logging

# Create log file
logging.basicConfig(filename="log.txt", level=logging.DEBUG, format="%(asctime)s %(message)s", filemode="a")
logging.info("Running 5_Count_By_Week.py...")

# Week variables - To change accordingly to QCP period
WeekA = "Week 19"
WeekB = "Week 21"
WeekC = "Week 22"
WeekD = "Week 23"
# WeekE = "Week 35"
# WeekF = "Week 36"
# WeekG = "Week 37"
# WeekH = "Week 38"
# Add more if needed

# Query variables - To change accordingly to QCP period
QueryWeekA_df = 'Week == 19'
QueryWeekB_df = 'Week == 21'
QueryWeekC_df = 'Week == 22'
QueryWeekD_df = 'Week == 23'
# QueryWeekE_df = 'Week == 35'
# QueryWeekF_df = 'Week == 36'
# QueryWeekG_df = 'Week == 37'
# QueryWeekH_df = 'Week == 38'
# Add more if needed

# Check if a previous version of XLSX files exists
try:
    # Pending Conversion
    os.path.exists("./z_Count.xlsx")
    os.remove("./z_Count.xlsx")
    print("A previous version of z_Count detected! Deleting file...")
except:
    logging.info("Proceed to export data...")
    print("Proceed to export data...")

# Define the list of Excel files
scanned_excel_files = [
    'Scanned OAS Franchisee.xlsx',
    'Scanned OAS Franchisor.xlsx',
]

pending_excel_files = [
    'Pending Conversion Franchisee.xlsx',
    'Pending Conversion Franchisor.xlsx',
]

outstanding_excel_files = [
    'Outstanding Scans Franchisee.xlsx',
    'Outstanding Scans Franchisor.xlsx',
]

abs_scanned_excel_files = [
    'ABS Scanned Franchisee.xlsx',
    'ABS Scanned Franchisor.xlsx',
]

abs_pending_excel_files = [
    'ABS Pending Conversion Franchisee.xlsx',
    'ABS Pending Conversion Franchisor.xlsx',
]

abs_outstanding_excel_files = [
    'ABS Outstanding Scans Franchisee.xlsx',
    'ABS Outstanding Scans Franchisor.xlsx',
]

# Create a DataFrame to store the counts
file_counts = pd.DataFrame(columns=['File Name', 'Week', 'OAS Count', 'Type', 'Status'])

# Iterate over the Excel files
for file in scanned_excel_files:
    # Read the Excel file
    df = pd.read_excel(file)

    # Dataframe Queries
    QueryWeekA = df.query(QueryWeekA_df)
    QueryWeekB = df.query(QueryWeekB_df)
    QueryWeekC = df.query(QueryWeekC_df)
    QueryWeekD = df.query(QueryWeekD_df)

    # Get the OAS Count
    CountWeekA = len(QueryWeekA)
    CountWeekB = len(QueryWeekB)
    CountWeekC = len(QueryWeekC)
    CountWeekD = len(QueryWeekD)

    # Store the file name and OAS Count in the DataFrame
    file_counts = file_counts.append({'File Name': file, 'Week': WeekA, 'OAS Count': CountWeekA, 'Type': 'Class', 'Status': 'Scanned'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Week': WeekB, 'OAS Count': CountWeekB, 'Type': 'Class', 'Status': 'Scanned'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Week': WeekC, 'OAS Count': CountWeekC, 'Type': 'Class', 'Status': 'Scanned'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Week': WeekD, 'OAS Count': CountWeekD, 'Type': 'Class', 'Status': 'Scanned'}, ignore_index=True)

for file in pending_excel_files:
    # Read the Excel file
    df = pd.read_excel(file)

    # Dataframe Queries
    QueryWeekA = df.query(QueryWeekA_df)
    QueryWeekB = df.query(QueryWeekB_df)
    QueryWeekC = df.query(QueryWeekC_df)
    QueryWeekD = df.query(QueryWeekD_df)

    # Get the OAS Count
    CountWeekA = len(QueryWeekA)
    CountWeekB = len(QueryWeekB)
    CountWeekC = len(QueryWeekC)
    CountWeekD = len(QueryWeekD)

    # Store the file name and OAS Count in the DataFrame
    file_counts = file_counts.append({'File Name': file, 'Week': WeekA, 'OAS Count': CountWeekA, 'Type': 'Class', 'Status': 'Pending Conversion'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Week': WeekB, 'OAS Count': CountWeekB, 'Type': 'Class', 'Status': 'Pending Conversion'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Week': WeekC, 'OAS Count': CountWeekC, 'Type': 'Class', 'Status': 'Pending Conversion'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Week': WeekD, 'OAS Count': CountWeekD, 'Type': 'Class', 'Status': 'Pending Conversion'}, ignore_index=True)

for file in outstanding_excel_files:
    # Read the Excel file
    df = pd.read_excel(file)

    # Dataframe Queries
    QueryWeekA = df.query(QueryWeekA_df)
    QueryWeekB = df.query(QueryWeekB_df)
    QueryWeekC = df.query(QueryWeekC_df)
    QueryWeekD = df.query(QueryWeekD_df)

    # Get the OAS Count
    CountWeekA = len(QueryWeekA)
    CountWeekB = len(QueryWeekB)
    CountWeekC = len(QueryWeekC)
    CountWeekD = len(QueryWeekD)

    # Store the file name and OAS Count in the DataFrame
    file_counts = file_counts.append({'File Name': file, 'Week': WeekA, 'OAS Count': CountWeekA, 'Type': 'Class', 'Status': 'Outstanding Scan(s)'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Week': WeekB, 'OAS Count': CountWeekB, 'Type': 'Class', 'Status': 'Outstanding Scan(s)'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Week': WeekC, 'OAS Count': CountWeekC, 'Type': 'Class', 'Status': 'Outstanding Scan(s)'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Week': WeekD, 'OAS Count': CountWeekD, 'Type': 'Class', 'Status': 'Outstanding Scan(s)'}, ignore_index=True)

for file in abs_scanned_excel_files:
    # Read the Excel file
    df = pd.read_excel(file)

    # Get the OAS Count
    record_count = len(df)

    # Store the file name and OAS Count in the DataFrame
    file_counts = file_counts.append({'File Name': file, 'Week': 'NIL', 'OAS Count': record_count, 'Type': 'Student', 'Status': 'ABS Scanned'}, ignore_index=True)

for file in abs_pending_excel_files:
    # Read the Excel file
    df = pd.read_excel(file)

    # Get the OAS Count
    record_count = len(df)

    # Store the file name and OAS Count in the DataFrame
    file_counts = file_counts.append({'File Name': file, 'Week': 'NIL', 'OAS Count': record_count, 'Type': 'Student', 'Status': 'ABS Pending Conversion'}, ignore_index=True)

for file in abs_outstanding_excel_files:
    # Read the Excel file
    df = pd.read_excel(file)

    # Get the OAS Count
    record_count = len(df)

    # Store the file name and OAS Count in the DataFrame
    file_counts = file_counts.append({'File Name': file, 'Week': 'NIL', 'OAS Count': record_count, 'Type': 'Student', 'Status': 'ABS Outstanding Scan(s)'}, ignore_index=True)

# Write the file counts to a new Excel file
file_counts.to_excel('z_Count.xlsx', index=False)