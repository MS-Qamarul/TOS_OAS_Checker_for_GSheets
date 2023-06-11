import pandas as pd
import os
import logging

# Create log file
logging.basicConfig(filename="log.txt", level=logging.DEBUG, format="%(asctime)s %(message)s", filemode="a")
logging.info("Running 5_Count_By_Group.py...")

# Week variables - To change accordingly to QCP period <<<
GroupA = "S1-3 Chinese (Week 31-32)"
GroupB = "S1-3 Sciences (Week 33-34)"
GroupC = "P3-5 Courses (Week 34-35)"
GroupD = "P6, OL Courses (Week 36-38)"
# Add more if needed

# Query variables - To change accordingly to QCP period <<<
QueryOne = '`Class: Class Code`.str.contains("S1C") or `Class: Class Code`.str.contains("S2C") or `Class: Class Code`.str.contains("S3C")'
QueryTwo = '`Class: Class Code`.str.contains("S1S") or `Class: Class Code`.str.contains("S2S") or `Class: Class Code`.str.contains("S3S")'
QueryThree = '`Class: Class Code`.str.contains("P3") or `Class: Class Code`.str.contains("P4") or `Class: Class Code`.str.contains("P5")'
QueryFour = '`Class: Class Code`.str.contains("P6") or `Class: Class Code`.str.contains("OL")'
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
file_counts = pd.DataFrame(columns=['File Name', 'Group', 'Record Count', 'Type', 'Status'])

# Iterate over the Excel files
for file in scanned_excel_files:
    # Read the Excel file
    df = pd.read_excel(file)

    # Dataframe Queries
    QueryA = df.query(QueryOne, engine = 'python')
    QueryB = df.query(QueryTwo, engine = 'python')
    QueryC = df.query(QueryThree, engine = 'python')
    QueryD = df.query(QueryFour, engine = 'python')

    # Get the record count
    CountQueryA = len(QueryA)
    CountQueryB = len(QueryB)
    CountQueryC = len(QueryC)
    CountQueryD = len(QueryD)

    # Store the file name and record count in the DataFrame
    file_counts = file_counts.append({'File Name': file, 'Group': GroupA, 'Record Count': CountQueryA, 'Type': 'Class', 'Status': 'Scanned'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Group': GroupB, 'Record Count': CountQueryB, 'Type': 'Class', 'Status': 'Scanned'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Group': GroupC, 'Record Count': CountQueryC, 'Type': 'Class', 'Status': 'Scanned'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Group': GroupD, 'Record Count': CountQueryD, 'Type': 'Class', 'Status': 'Scanned'}, ignore_index=True)

for file in pending_excel_files:
    # Read the Excel file
    df = pd.read_excel(file)

    # Dataframe Queries
    QueryA = df.query(QueryOne, engine = 'python')
    QueryB = df.query(QueryTwo, engine = 'python')
    QueryC = df.query(QueryThree, engine = 'python')
    QueryD = df.query(QueryFour, engine = 'python')

    # Get the record count
    CountQueryA = len(QueryA)
    CountQueryB = len(QueryB)
    CountQueryC = len(QueryC)
    CountQueryD = len(QueryD)

    # Store the file name and record count in the DataFrame
    file_counts = file_counts.append({'File Name': file, 'Group': GroupA, 'Record Count': CountQueryA, 'Type': 'Class', 'Status': 'Pending Conversion'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Group': GroupB, 'Record Count': CountQueryB, 'Type': 'Class', 'Status': 'Pending Conversion'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Group': GroupC, 'Record Count': CountQueryC, 'Type': 'Class', 'Status': 'Pending Conversion'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Group': GroupD, 'Record Count': CountQueryD, 'Type': 'Class', 'Status': 'Pending Conversion'}, ignore_index=True)

for file in outstanding_excel_files:
    # Read the Excel file
    df = pd.read_excel(file)

    # Dataframe Queries
    QueryA = df.query(QueryOne, engine = 'python')
    QueryB = df.query(QueryTwo, engine = 'python')
    QueryC = df.query(QueryThree, engine = 'python')
    QueryD = df.query(QueryFour, engine = 'python')

    # Get the record count
    CountQueryA = len(QueryA)
    CountQueryB = len(QueryB)
    CountQueryC = len(QueryC)
    CountQueryD = len(QueryD)

    # Store the file name and record count in the DataFrame
    file_counts = file_counts.append({'File Name': file, 'Group': GroupA, 'Record Count': CountQueryA, 'Type': 'Class', 'Status': 'Outstanding Scan(s)'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Group': GroupB, 'Record Count': CountQueryB, 'Type': 'Class', 'Status': 'Outstanding Scan(s)'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Group': GroupC, 'Record Count': CountQueryC, 'Type': 'Class', 'Status': 'Outstanding Scan(s)'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Group': GroupD, 'Record Count': CountQueryD, 'Type': 'Class', 'Status': 'Outstanding Scan(s)'}, ignore_index=True)

for file in abs_scanned_excel_files:
    # Read the Excel file
    df = pd.read_excel(file)

    # Get the record count
    record_count = len(df)

    # Store the file name and record count in the DataFrame
    file_counts = file_counts.append({'File Name': file, 'Group': 'Absentees', 'Record Count': record_count, 'Type': 'Student', 'Status': 'ABS Scanned'}, ignore_index=True)

for file in abs_pending_excel_files:
    # Read the Excel file
    df = pd.read_excel(file)

    # Get the record count
    record_count = len(df)

    # Store the file name and record count in the DataFrame
    file_counts = file_counts.append({'File Name': file, 'Group': 'Absentees', 'Record Count': record_count, 'Type': 'Student', 'Status': 'ABS Pending Conversion'}, ignore_index=True)

for file in abs_outstanding_excel_files:
    # Read the Excel file
    df = pd.read_excel(file)

    # Get the record count
    record_count = len(df)

    # Store the file name and record count in the DataFrame
    file_counts = file_counts.append({'File Name': file, 'Group': 'Absentees', 'Record Count': record_count, 'Type': 'Student', 'Status': 'ABS Outstanding Scan(s)'}, ignore_index=True)

# Write the file counts to a new Excel file
file_counts.to_excel('z_Count.xlsx', index=False)