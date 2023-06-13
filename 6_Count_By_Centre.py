import pandas as pd
import os
import logging

# Create log file
logging.basicConfig(filename="log.txt", level=logging.DEBUG, format="%(asctime)s %(message)s", filemode="a")
logging.info("Running 6_Count_By_Centre.py...")

# List of MS Centres
GroupA = "Balmoral"
GroupB = "Bishan"
GroupC = "Bukit Batok Central"
GroupD = "Bukit Panjang"
GroupE = "Central Campus"
GroupF = "Choa Chu Kang"
GroupG = "East Campus"
GroupH = "Hougang"
GroupI = "Jurong West"
GroupJ = "Lequest"
GroupK = "Marine Parade"
GroupL = "Northshore"
GroupM = "Pioneer"
GroupN = "Punggol Central"
GroupO = "Punggol Plaza"
GroupP = "Rivervale"
GroupQ = "Seletar"
GroupR = "Sembawang"
GroupS = "Simei"
GroupT = "West Campus"
GroupU = "Woodlands Central"
GroupV = "Woodleigh"
GroupW = "Yishun"
# Add more if needed

# Query variables
QueryA = '`Class: Class Code`.str.contains("BM2")'
QueryB = '`Class: Class Code`.str.contains("BS2")'
QueryC = '`Class: Class Code`.str.contains("BBC2")'
QueryD = '`Class: Class Code`.str.contains("BP2")'
QueryE = '`Class: Class Code`.str.contains("TP2")'
QueryF = '`Class: Class Code`.str.contains("CCK2")'
QueryG = '`Class: Class Code`.str.contains("EC2")'
QueryH = '`Class: Class Code`.str.contains("HG2")'
QueryI = '`Class: Class Code`.str.contains("JW2")'
QueryJ = '`Class: Class Code`.str.contains("LQ2")'
QueryK = '`Class: Class Code`.str.contains("MP2")'
QueryL = '`Class: Class Code`.str.contains("NS2")'
QueryM = '`Class: Class Code`.str.contains("PN2")'
QueryN = '`Class: Class Code`.str.contains("PC2")'
QueryO = '`Class: Class Code`.str.contains("PG2")'
QueryP = '`Class: Class Code`.str.contains("RM2")'
QueryQ = '`Class: Class Code`.str.contains("SM2")'
QueryR = '`Class: Class Code`.str.contains("SW2")'
QueryS = '`Class: Class Code`.str.contains("SME2")'
QueryT = '`Class: Class Code`.str.contains("WC2")'
QueryU = '`Class: Class Code`.str.contains("WLC2")'
QueryV = '`Class: Class Code`.str.contains("WLH2")'
QueryW = '`Class: Class Code`.str.contains("YS2")'
# Add more if needed

# Check if a previous version of XLSX files exists
try:
    # Pending Conversion
    os.path.exists("./z_Count_By_Centre.xlsx")
    os.remove("./z_Count_By_Centre.xlsx")
    print("A previous version of z_Count_By_Centre detected! Deleting file...")
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

# Create a DataFrame to store the counts
file_counts = pd.DataFrame(columns=['File Name', 'Centre', 'OAS Count', 'Type', 'Status'])

# Iterate over the Excel files
for file in scanned_excel_files:
    # Read the Excel file
    df = pd.read_excel(file)

    # Dataframe Queries
    Query1 = df.query(QueryA, engine = 'python')
    Query2 = df.query(QueryB, engine = 'python')
    Query3 = df.query(QueryC, engine = 'python')
    Query4 = df.query(QueryD, engine = 'python')
    Query5 = df.query(QueryE, engine = 'python')
    Query6 = df.query(QueryF, engine = 'python')
    Query7 = df.query(QueryG, engine = 'python')
    Query8 = df.query(QueryH, engine = 'python')
    Query9 = df.query(QueryI, engine = 'python')
    Query10 = df.query(QueryJ, engine = 'python')
    Query11 = df.query(QueryK, engine = 'python')
    Query12 = df.query(QueryL, engine = 'python')
    Query13 = df.query(QueryM, engine = 'python')
    Query14 = df.query(QueryN, engine = 'python')
    Query15 = df.query(QueryO, engine = 'python')
    Query16 = df.query(QueryP, engine = 'python')
    Query17 = df.query(QueryQ, engine = 'python')
    Query18 = df.query(QueryR, engine = 'python')
    Query19 = df.query(QueryS, engine = 'python')
    Query20 = df.query(QueryT, engine = 'python')
    Query21 = df.query(QueryU, engine = 'python')
    Query22 = df.query(QueryV, engine = 'python')
    Query23 = df.query(QueryW, engine = 'python')

    # Get the OAS Count
    CountQueryA = len(Query1)
    CountQueryB = len(Query2)
    CountQueryC = len(Query3)
    CountQueryD = len(Query4)
    CountQueryE = len(Query5)
    CountQueryF = len(Query6)
    CountQueryG = len(Query7)
    CountQueryH = len(Query8)
    CountQueryI = len(Query9)
    CountQueryJ = len(Query10)
    CountQueryK = len(Query11)
    CountQueryL = len(Query12)
    CountQueryM = len(Query13)
    CountQueryN = len(Query14)
    CountQueryO = len(Query15)
    CountQueryP = len(Query16)
    CountQueryQ = len(Query17)
    CountQueryR = len(Query18)
    CountQueryS = len(Query19)
    CountQueryT = len(Query20)
    CountQueryU = len(Query21)
    CountQueryV = len(Query22)
    CountQueryW = len(Query23)

    # Store the file name and OAS Count in the DataFrame
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupA, 'OAS Count': CountQueryA, 'Type': 'Class', 'Status': 'Scanned'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupB, 'OAS Count': CountQueryB, 'Type': 'Class', 'Status': 'Scanned'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupC, 'OAS Count': CountQueryC, 'Type': 'Class', 'Status': 'Scanned'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupD, 'OAS Count': CountQueryD, 'Type': 'Class', 'Status': 'Scanned'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupE, 'OAS Count': CountQueryE, 'Type': 'Class', 'Status': 'Scanned'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupF, 'OAS Count': CountQueryF, 'Type': 'Class', 'Status': 'Scanned'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupG, 'OAS Count': CountQueryG, 'Type': 'Class', 'Status': 'Scanned'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupH, 'OAS Count': CountQueryH, 'Type': 'Class', 'Status': 'Scanned'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupI, 'OAS Count': CountQueryI, 'Type': 'Class', 'Status': 'Scanned'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupJ, 'OAS Count': CountQueryJ, 'Type': 'Class', 'Status': 'Scanned'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupK, 'OAS Count': CountQueryK, 'Type': 'Class', 'Status': 'Scanned'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupL, 'OAS Count': CountQueryL, 'Type': 'Class', 'Status': 'Scanned'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupM, 'OAS Count': CountQueryM, 'Type': 'Class', 'Status': 'Scanned'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupN, 'OAS Count': CountQueryN, 'Type': 'Class', 'Status': 'Scanned'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupO, 'OAS Count': CountQueryO, 'Type': 'Class', 'Status': 'Scanned'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupP, 'OAS Count': CountQueryP, 'Type': 'Class', 'Status': 'Scanned'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupQ, 'OAS Count': CountQueryQ, 'Type': 'Class', 'Status': 'Scanned'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupR, 'OAS Count': CountQueryR, 'Type': 'Class', 'Status': 'Scanned'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupS, 'OAS Count': CountQueryS, 'Type': 'Class', 'Status': 'Scanned'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupT, 'OAS Count': CountQueryT, 'Type': 'Class', 'Status': 'Scanned'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupU, 'OAS Count': CountQueryU, 'Type': 'Class', 'Status': 'Scanned'}, ignore_index=True)
    # file_counts = file_counts.append({'File Name': file, 'Centre': GroupV, 'OAS Count': CountQueryV, 'Type': 'Class', 'Status': 'Scanned'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupW, 'OAS Count': CountQueryW, 'Type': 'Class', 'Status': 'Scanned'}, ignore_index=True)

for file in pending_excel_files:
    # Read the Excel file
    df = pd.read_excel(file)

    # Dataframe Queries
    Query1 = df.query(QueryA, engine = 'python')
    Query2 = df.query(QueryB, engine = 'python')
    Query3 = df.query(QueryC, engine = 'python')
    Query4 = df.query(QueryD, engine = 'python')
    Query5 = df.query(QueryE, engine = 'python')
    Query6 = df.query(QueryF, engine = 'python')
    Query7 = df.query(QueryG, engine = 'python')
    Query8 = df.query(QueryH, engine = 'python')
    Query9 = df.query(QueryI, engine = 'python')
    Query10 = df.query(QueryJ, engine = 'python')
    Query11 = df.query(QueryK, engine = 'python')
    Query12 = df.query(QueryL, engine = 'python')
    Query13 = df.query(QueryM, engine = 'python')
    Query14 = df.query(QueryN, engine = 'python')
    Query15 = df.query(QueryO, engine = 'python')
    Query16 = df.query(QueryP, engine = 'python')
    Query17 = df.query(QueryQ, engine = 'python')
    Query18 = df.query(QueryR, engine = 'python')
    Query19 = df.query(QueryS, engine = 'python')
    Query20 = df.query(QueryT, engine = 'python')
    Query21 = df.query(QueryU, engine = 'python')
    Query22 = df.query(QueryV, engine = 'python')
    Query23 = df.query(QueryW, engine = 'python')

    # Get the OAS Count
    CountQueryA = len(Query1)
    CountQueryB = len(Query2)
    CountQueryC = len(Query3)
    CountQueryD = len(Query4)
    CountQueryE = len(Query5)
    CountQueryF = len(Query6)
    CountQueryG = len(Query7)
    CountQueryH = len(Query8)
    CountQueryI = len(Query9)
    CountQueryJ = len(Query10)
    CountQueryK = len(Query11)
    CountQueryL = len(Query12)
    CountQueryM = len(Query13)
    CountQueryN = len(Query14)
    CountQueryO = len(Query15)
    CountQueryP = len(Query16)
    CountQueryQ = len(Query17)
    CountQueryR = len(Query18)
    CountQueryS = len(Query19)
    CountQueryT = len(Query20)
    CountQueryU = len(Query21)
    CountQueryV = len(Query22)
    CountQueryW = len(Query23)

    # Store the file name and OAS Count in the DataFrame
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupA, 'OAS Count': CountQueryA, 'Type': 'Class', 'Status': 'Pending Conversion'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupB, 'OAS Count': CountQueryB, 'Type': 'Class', 'Status': 'Pending Conversion'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupC, 'OAS Count': CountQueryC, 'Type': 'Class', 'Status': 'Pending Conversion'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupD, 'OAS Count': CountQueryD, 'Type': 'Class', 'Status': 'Pending Conversion'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupE, 'OAS Count': CountQueryE, 'Type': 'Class', 'Status': 'Pending Conversion'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupF, 'OAS Count': CountQueryF, 'Type': 'Class', 'Status': 'Pending Conversion'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupG, 'OAS Count': CountQueryG, 'Type': 'Class', 'Status': 'Pending Conversion'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupH, 'OAS Count': CountQueryH, 'Type': 'Class', 'Status': 'Pending Conversion'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupI, 'OAS Count': CountQueryI, 'Type': 'Class', 'Status': 'Pending Conversion'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupJ, 'OAS Count': CountQueryJ, 'Type': 'Class', 'Status': 'Pending Conversion'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupK, 'OAS Count': CountQueryK, 'Type': 'Class', 'Status': 'Pending Conversion'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupL, 'OAS Count': CountQueryL, 'Type': 'Class', 'Status': 'Pending Conversion'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupM, 'OAS Count': CountQueryM, 'Type': 'Class', 'Status': 'Pending Conversion'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupN, 'OAS Count': CountQueryN, 'Type': 'Class', 'Status': 'Pending Conversion'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupO, 'OAS Count': CountQueryO, 'Type': 'Class', 'Status': 'Pending Conversion'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupP, 'OAS Count': CountQueryP, 'Type': 'Class', 'Status': 'Pending Conversion'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupQ, 'OAS Count': CountQueryQ, 'Type': 'Class', 'Status': 'Pending Conversion'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupR, 'OAS Count': CountQueryR, 'Type': 'Class', 'Status': 'Pending Conversion'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupS, 'OAS Count': CountQueryS, 'Type': 'Class', 'Status': 'Pending Conversion'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupT, 'OAS Count': CountQueryT, 'Type': 'Class', 'Status': 'Pending Conversion'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupU, 'OAS Count': CountQueryU, 'Type': 'Class', 'Status': 'Pending Conversion'}, ignore_index=True)
    # file_counts = file_counts.append({'File Name': file, 'Centre': GroupV, 'OAS Count': CountQueryV, 'Type': 'Class', 'Status': 'Pending Conversion'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupW, 'OAS Count': CountQueryW, 'Type': 'Class', 'Status': 'Pending Conversion'}, ignore_index=True)

for file in outstanding_excel_files:
    # Read the Excel file
    df = pd.read_excel(file)

    # Dataframe Queries
    Query1 = df.query(QueryA, engine = 'python')
    Query2 = df.query(QueryB, engine = 'python')
    Query3 = df.query(QueryC, engine = 'python')
    Query4 = df.query(QueryD, engine = 'python')
    Query5 = df.query(QueryE, engine = 'python')
    Query6 = df.query(QueryF, engine = 'python')
    Query7 = df.query(QueryG, engine = 'python')
    Query8 = df.query(QueryH, engine = 'python')
    Query9 = df.query(QueryI, engine = 'python')
    Query10 = df.query(QueryJ, engine = 'python')
    Query11 = df.query(QueryK, engine = 'python')
    Query12 = df.query(QueryL, engine = 'python')
    Query13 = df.query(QueryM, engine = 'python')
    Query14 = df.query(QueryN, engine = 'python')
    Query15 = df.query(QueryO, engine = 'python')
    Query16 = df.query(QueryP, engine = 'python')
    Query17 = df.query(QueryQ, engine = 'python')
    Query18 = df.query(QueryR, engine = 'python')
    Query19 = df.query(QueryS, engine = 'python')
    Query20 = df.query(QueryT, engine = 'python')
    Query21 = df.query(QueryU, engine = 'python')
    Query22 = df.query(QueryV, engine = 'python')
    Query23 = df.query(QueryW, engine = 'python')

    # Get the OAS Count
    CountQueryA = len(Query1)
    CountQueryB = len(Query2)
    CountQueryC = len(Query3)
    CountQueryD = len(Query4)
    CountQueryE = len(Query5)
    CountQueryF = len(Query6)
    CountQueryG = len(Query7)
    CountQueryH = len(Query8)
    CountQueryI = len(Query9)
    CountQueryJ = len(Query10)
    CountQueryK = len(Query11)
    CountQueryL = len(Query12)
    CountQueryM = len(Query13)
    CountQueryN = len(Query14)
    CountQueryO = len(Query15)
    CountQueryP = len(Query16)
    CountQueryQ = len(Query17)
    CountQueryR = len(Query18)
    CountQueryS = len(Query19)
    CountQueryT = len(Query20)
    CountQueryU = len(Query21)
    CountQueryV = len(Query22)
    CountQueryW = len(Query23)

    # Store the file name and OAS Count in the DataFrame
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupA, 'OAS Count': CountQueryA, 'Type': 'Class', 'Status': 'Outstanding Scan(s)'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupB, 'OAS Count': CountQueryB, 'Type': 'Class', 'Status': 'Outstanding Scan(s)'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupC, 'OAS Count': CountQueryC, 'Type': 'Class', 'Status': 'Outstanding Scan(s)'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupD, 'OAS Count': CountQueryD, 'Type': 'Class', 'Status': 'Outstanding Scan(s)'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupE, 'OAS Count': CountQueryE, 'Type': 'Class', 'Status': 'Outstanding Scan(s)'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupF, 'OAS Count': CountQueryF, 'Type': 'Class', 'Status': 'Outstanding Scan(s)'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupG, 'OAS Count': CountQueryG, 'Type': 'Class', 'Status': 'Outstanding Scan(s)'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupH, 'OAS Count': CountQueryH, 'Type': 'Class', 'Status': 'Outstanding Scan(s)'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupI, 'OAS Count': CountQueryI, 'Type': 'Class', 'Status': 'Outstanding Scan(s)'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupJ, 'OAS Count': CountQueryJ, 'Type': 'Class', 'Status': 'Outstanding Scan(s)'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupK, 'OAS Count': CountQueryK, 'Type': 'Class', 'Status': 'Outstanding Scan(s)'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupL, 'OAS Count': CountQueryL, 'Type': 'Class', 'Status': 'Outstanding Scan(s)'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupM, 'OAS Count': CountQueryM, 'Type': 'Class', 'Status': 'Outstanding Scan(s)'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupN, 'OAS Count': CountQueryN, 'Type': 'Class', 'Status': 'Outstanding Scan(s)'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupO, 'OAS Count': CountQueryO, 'Type': 'Class', 'Status': 'Outstanding Scan(s)'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupP, 'OAS Count': CountQueryP, 'Type': 'Class', 'Status': 'Outstanding Scan(s)'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupQ, 'OAS Count': CountQueryQ, 'Type': 'Class', 'Status': 'Outstanding Scan(s)'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupR, 'OAS Count': CountQueryR, 'Type': 'Class', 'Status': 'Outstanding Scan(s)'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupS, 'OAS Count': CountQueryS, 'Type': 'Class', 'Status': 'Outstanding Scan(s)'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupT, 'OAS Count': CountQueryT, 'Type': 'Class', 'Status': 'Outstanding Scan(s)'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupU, 'OAS Count': CountQueryU, 'Type': 'Class', 'Status': 'Outstanding Scan(s)'}, ignore_index=True)
    # file_counts = file_counts.append({'File Name': file, 'Centre': GroupV, 'OAS Count': CountQueryV, 'Type': 'Class', 'Status': 'Outstanding Scan(s)'}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupW, 'OAS Count': CountQueryW, 'Type': 'Class', 'Status': 'Outstanding Scan(s)'}, ignore_index=True)

# Write the file counts to a new Excel file
file_counts.to_excel('z_Count_By_Centre.xlsx', index=False)