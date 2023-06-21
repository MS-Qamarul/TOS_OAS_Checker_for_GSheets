import pandas as pd
import os
import logging

# Create log file
logging.basicConfig(filename="log.txt", level=logging.DEBUG, format="%(asctime)s %(message)s", filemode="a")
logging.info("Running 6_Final.py...")

# Scan Status
ScanStatusTrue = "Scanned"
ScanStatusFalse = "Outstanding"

# Conversion Status
ConStatsTrue = "Converted"
ConStatsFalse = "Pending Conversion"

# Week variables - To change accordingly to QCP period <<<
Group1 = "P6, OL Courses (Week 36-38)"
Group2 = "S1-3 Chinese (Week 31-32)"
Group3 = "P3-5 Courses (Week 34-35)"
Group4 = "S1-3 Sciences (Week 33-34)"

# Types
Type0 = "Absentee"
Type1 = "Class"
Type2 = "Students"

# List of MS Centres
GroupABS = "ABS"
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
GroupV = "PO"
GroupW = "Yishun"
# Add more if needed

# Query variables
QueryA = '`Class: Class Code`.str.contains("BM2023-P6") or `Class: Class Code`.str.contains("BM2023-OL")'
QueryB = '`Class: Class Code`.str.contains("BS2023-P6") or `Class: Class Code`.str.contains("BS2023-OL")'
QueryC = '`Class: Class Code`.str.contains("BBC2023-P6") or `Class: Class Code`.str.contains("BBC2023-OL")'
QueryD = '`Class: Class Code`.str.contains("BP2023-P6") or `Class: Class Code`.str.contains("BP2023-OL")'
QueryE = '`Class: Class Code`.str.contains("TP2023-P6") or `Class: Class Code`.str.contains("TP2023-OL")'
QueryF = '`Class: Class Code`.str.contains("CCK2023-P6") or `Class: Class Code`.str.contains("CCK2023-OL")'
QueryG = '`Class: Class Code`.str.contains("EC2023-P6") or `Class: Class Code`.str.contains("EC2023-OL")'
QueryH = '`Class: Class Code`.str.contains("HG2023-P6") or `Class: Class Code`.str.contains("HG2023-OL")'
QueryI = '`Class: Class Code`.str.contains("JW2023-P6") or `Class: Class Code`.str.contains("JW2023-OL")'
QueryJ = '`Class: Class Code`.str.contains("LQ2023-P6") or `Class: Class Code`.str.contains("LQ2023-OL")'
QueryK = '`Class: Class Code`.str.contains("MP2023-P6") or `Class: Class Code`.str.contains("MP2023-OL")'
QueryL = '`Class: Class Code`.str.contains("NS2023-P6") or `Class: Class Code`.str.contains("NS2023-OL")'
QueryM = '`Class: Class Code`.str.contains("PN2023-P6") or `Class: Class Code`.str.contains("PN2023-OL")'
QueryN = '`Class: Class Code`.str.contains("PC2023-P6") or `Class: Class Code`.str.contains("PC2023-OL")'
QueryO = '`Class: Class Code`.str.contains("PG2023-P6") or `Class: Class Code`.str.contains("PG2023-OL")'
QueryP = '`Class: Class Code`.str.contains("RM2023-P6") or `Class: Class Code`.str.contains("RM2023-OL")'
QueryQ = '`Class: Class Code`.str.contains("SM2023-P6") or `Class: Class Code`.str.contains("SM2023-OL")'
QueryR = '`Class: Class Code`.str.contains("SW2023-P6") or `Class: Class Code`.str.contains("SW2023-OL")'
QueryS = '`Class: Class Code`.str.contains("SME2023-P6") or `Class: Class Code`.str.contains("SME2023-OL")'
QueryT = '`Class: Class Code`.str.contains("WC2023-P6") or `Class: Class Code`.str.contains("WC2023-OL")'
QueryU = '`Class: Class Code`.str.contains("WLC2023-P6") or `Class: Class Code`.str.contains("WLC2023-OL")'
QueryV = '`Class: Class Code`.str.contains("ED2023-P6") or `Class: Class Code`.str.contains("ED2023-OL")'
QueryW = '`Class: Class Code`.str.contains("YS2023-P6") or `Class: Class Code`.str.contains("YS2023-OL")'
# Add more if needed

# Query variables
QueryAA = '`Class: Class Code`.str.contains("BM2023-S1C") or `Class: Class Code`.str.contains("BM2023-S2C") or `Class: Class Code`.str.contains("BM2023-S3C")'
QueryBB = '`Class: Class Code`.str.contains("BS2023-S1C") or `Class: Class Code`.str.contains("BS2023-S2C") or `Class: Class Code`.str.contains("BS2023-S3C")'
QueryCC = '`Class: Class Code`.str.contains("BBC2023-S1C") or `Class: Class Code`.str.contains("BBC2023-S2C") or `Class: Class Code`.str.contains("BBC2023-S3C")'
QueryDD = '`Class: Class Code`.str.contains("BP2023-S1C") or `Class: Class Code`.str.contains("BP2023-S2C") or `Class: Class Code`.str.contains("BP2023-S3C")'
QueryEE = '`Class: Class Code`.str.contains("TP2023-S1C") or `Class: Class Code`.str.contains("TP2023-S2C") or `Class: Class Code`.str.contains("TP2023-S3C")'
QueryFF = '`Class: Class Code`.str.contains("CCK2023-S1C") or `Class: Class Code`.str.contains("CCK2023-S2C") or `Class: Class Code`.str.contains("CCK2023-S3C")'
QueryGG = '`Class: Class Code`.str.contains("EC2023-S1C") or `Class: Class Code`.str.contains("EC2023-S2C") or `Class: Class Code`.str.contains("EC2023-S3C")'
QueryHH = '`Class: Class Code`.str.contains("HG2023-S1C") or `Class: Class Code`.str.contains("HG2023-S2C") or `Class: Class Code`.str.contains("HG2023-S3C")'
QueryII = '`Class: Class Code`.str.contains("JW2023-S1C") or `Class: Class Code`.str.contains("JW2023-S2C") or `Class: Class Code`.str.contains("JW2023-S3C")'
QueryJJ = '`Class: Class Code`.str.contains("LQ2023-S1C") or `Class: Class Code`.str.contains("LQ2023-S2C") or `Class: Class Code`.str.contains("LQ2023-S3C")'
QueryKK = '`Class: Class Code`.str.contains("MP2023-S1C") or `Class: Class Code`.str.contains("MP2023-S2C") or `Class: Class Code`.str.contains("MP2023-S3C")'
QueryLL = '`Class: Class Code`.str.contains("NS2023-S1C") or `Class: Class Code`.str.contains("NS2023-S2C") or `Class: Class Code`.str.contains("NS2023-S3C")'
QueryMM = '`Class: Class Code`.str.contains("PN2023-S1C") or `Class: Class Code`.str.contains("PN2023-S2C") or `Class: Class Code`.str.contains("PN2023-S3C")'
QueryNN = '`Class: Class Code`.str.contains("PC2023-S1C") or `Class: Class Code`.str.contains("PC2023-S2C") or `Class: Class Code`.str.contains("PC2023-S3C")'
QueryOO = '`Class: Class Code`.str.contains("PG2023-S1C") or `Class: Class Code`.str.contains("PG2023-S2C") or `Class: Class Code`.str.contains("PG2023-S3C")'
QueryPP = '`Class: Class Code`.str.contains("RM2023-S1C") or `Class: Class Code`.str.contains("RM2023-S2C") or `Class: Class Code`.str.contains("RM2023-S3C")'
QueryQQ = '`Class: Class Code`.str.contains("SM2023-S1C") or `Class: Class Code`.str.contains("SM2023-S2C") or `Class: Class Code`.str.contains("SM2023-S3C")'
QueryRR = '`Class: Class Code`.str.contains("SW2023-S1C") or `Class: Class Code`.str.contains("SW2023-S2C") or `Class: Class Code`.str.contains("SW2023-S3C")'
QuerySS = '`Class: Class Code`.str.contains("SME2023-S1C") or `Class: Class Code`.str.contains("SME2023-S2C") or `Class: Class Code`.str.contains("SME2023-S3C")'
QueryTT = '`Class: Class Code`.str.contains("WC2023-S1C") or `Class: Class Code`.str.contains("WC2023-S2C") or `Class: Class Code`.str.contains("WC2023-S3C")'
QueryUU = '`Class: Class Code`.str.contains("WLC2023-S1C") or `Class: Class Code`.str.contains("WLC2023-S2C") or `Class: Class Code`.str.contains("WLC2023-S3C")'
QueryVV = '`Class: Class Code`.str.contains("ED2023-S1C") or `Class: Class Code`.str.contains("ED2023-S2C") or `Class: Class Code`.str.contains("ED2023-S3C")'
QueryWW = '`Class: Class Code`.str.contains("YS2023-S1C") or `Class: Class Code`.str.contains("YS2023-S2C") or `Class: Class Code`.str.contains("YS2023-S3C")'
# Add more if needed

# Query variables
QueryAAA = '`Class: Class Code`.str.contains("BM2023-P3") or `Class: Class Code`.str.contains("BM2023-P4") or `Class: Class Code`.str.contains("BM2023-P5")'
QueryBBB = '`Class: Class Code`.str.contains("BS2023-P3") or `Class: Class Code`.str.contains("BS2023-P4") or `Class: Class Code`.str.contains("BS2023-P5")'
QueryCCC = '`Class: Class Code`.str.contains("BBC2023-P3") or `Class: Class Code`.str.contains("BBC2023-P4") or `Class: Class Code`.str.contains("BBC2023-P5")'
QueryDDD = '`Class: Class Code`.str.contains("BP2023-P3") or `Class: Class Code`.str.contains("BP2023-P4") or `Class: Class Code`.str.contains("BP2023-P5")'
QueryEEE = '`Class: Class Code`.str.contains("TP2023-P3") or `Class: Class Code`.str.contains("TP2023-P4") or `Class: Class Code`.str.contains("TP2023-P5")'
QueryFFF = '`Class: Class Code`.str.contains("CCK2023-P3") or `Class: Class Code`.str.contains("CCK2023-P4") or `Class: Class Code`.str.contains("CCK2023-P5")'
QueryGGG = '`Class: Class Code`.str.contains("EC2023-P3") or `Class: Class Code`.str.contains("EC2023-P4") or `Class: Class Code`.str.contains("EC2023-P5")'
QueryHHH = '`Class: Class Code`.str.contains("HG2023-P3") or `Class: Class Code`.str.contains("HG2023-P4") or `Class: Class Code`.str.contains("HG2023-P5")'
QueryIII = '`Class: Class Code`.str.contains("JW2023-P3") or `Class: Class Code`.str.contains("JW2023-P4") or `Class: Class Code`.str.contains("JW2023-P5")'
QueryJJJ = '`Class: Class Code`.str.contains("LQ2023-P3") or `Class: Class Code`.str.contains("LQ2023-P4") or `Class: Class Code`.str.contains("LQ2023-P5")'
QueryKKK = '`Class: Class Code`.str.contains("MP2023-P3") or `Class: Class Code`.str.contains("MP2023-P4") or `Class: Class Code`.str.contains("MP2023-P5")'
QueryLLL = '`Class: Class Code`.str.contains("NS2023-P3") or `Class: Class Code`.str.contains("NS2023-P4") or `Class: Class Code`.str.contains("NS2023-P5")'
QueryMMM = '`Class: Class Code`.str.contains("PN2023-P3") or `Class: Class Code`.str.contains("PN2023-P4") or `Class: Class Code`.str.contains("PN2023-P5")'
QueryNNN = '`Class: Class Code`.str.contains("PC2023-P3") or `Class: Class Code`.str.contains("PC2023-P4") or `Class: Class Code`.str.contains("PC2023-P5")'
QueryOOO = '`Class: Class Code`.str.contains("PG2023-P3") or `Class: Class Code`.str.contains("PG2023-P4") or `Class: Class Code`.str.contains("PG2023-P5")'
QueryPPP = '`Class: Class Code`.str.contains("RM2023-P3") or `Class: Class Code`.str.contains("RM2023-P4") or `Class: Class Code`.str.contains("RM2023-P5")'
QueryQQQ = '`Class: Class Code`.str.contains("SM2023-P3") or `Class: Class Code`.str.contains("SM2023-P4") or `Class: Class Code`.str.contains("SM2023-P5")'
QueryRRR = '`Class: Class Code`.str.contains("SW2023-P3") or `Class: Class Code`.str.contains("SW2023-P4") or `Class: Class Code`.str.contains("SW2023-P5")'
QuerySSS = '`Class: Class Code`.str.contains("SME2023-P3") or `Class: Class Code`.str.contains("SME2023-P4") or `Class: Class Code`.str.contains("SME2023-P5")'
QueryTTT = '`Class: Class Code`.str.contains("WC2023-P3") or `Class: Class Code`.str.contains("WC2023-P4") or `Class: Class Code`.str.contains("WC2023-P5")'
QueryUUU = '`Class: Class Code`.str.contains("WLC2023-P3") or `Class: Class Code`.str.contains("WLC2023-P4") or `Class: Class Code`.str.contains("WLC2023-P5")'
QueryVVV = '`Class: Class Code`.str.contains("ED2023-P3") or `Class: Class Code`.str.contains("ED2023-P4") or `Class: Class Code`.str.contains("ED2023-P5")'
QueryWWW = '`Class: Class Code`.str.contains("YS2023-P3") or `Class: Class Code`.str.contains("YS2023-P4") or `Class: Class Code`.str.contains("YS2023-P5")'
# Add more if needed

# Query variables
QueryAAAA = '`Class: Class Code`.str.contains("BM2023-S1S") or `Class: Class Code`.str.contains("BM2023-S2S") or `Class: Class Code`.str.contains("BM2023-S3S")'
QueryBBBB = '`Class: Class Code`.str.contains("BS2023-S1S") or `Class: Class Code`.str.contains("BS2023-S2S") or `Class: Class Code`.str.contains("BS2023-S3S")'
QueryCCCC = '`Class: Class Code`.str.contains("BBC2023-S1S") or `Class: Class Code`.str.contains("BBC2023-S2S") or `Class: Class Code`.str.contains("BBC2023-S3S")'
QueryDDDD = '`Class: Class Code`.str.contains("BP2023-S1S") or `Class: Class Code`.str.contains("BP2023-S2S") or `Class: Class Code`.str.contains("BP2023-S3S")'
QueryEEEE = '`Class: Class Code`.str.contains("TP2023-S1S") or `Class: Class Code`.str.contains("TP2023-S2S") or `Class: Class Code`.str.contains("TP2023-S3S")'
QueryFFFF = '`Class: Class Code`.str.contains("CCK2023-S1S") or `Class: Class Code`.str.contains("CCK2023-S2S") or `Class: Class Code`.str.contains("CCK2023-S3S")'
QueryGGGG = '`Class: Class Code`.str.contains("EC2023-S1S") or `Class: Class Code`.str.contains("EC2023-S2S") or `Class: Class Code`.str.contains("EC2023-S3S")'
QueryHHHH = '`Class: Class Code`.str.contains("HG2023-S1S") or `Class: Class Code`.str.contains("HG2023-S2S") or `Class: Class Code`.str.contains("HG2023-S3S")'
QueryIIII = '`Class: Class Code`.str.contains("JW2023-S1S") or `Class: Class Code`.str.contains("JW2023-S2S") or `Class: Class Code`.str.contains("JW2023-S3S")'
QueryJJJJ = '`Class: Class Code`.str.contains("LQ2023-S1S") or `Class: Class Code`.str.contains("LQ2023-S2S") or `Class: Class Code`.str.contains("LQ2023-S3S")'
QueryKKKK = '`Class: Class Code`.str.contains("MP2023-S1S") or `Class: Class Code`.str.contains("MP2023-S2S") or `Class: Class Code`.str.contains("MP2023-S3S")'
QueryLLLL = '`Class: Class Code`.str.contains("NS2023-S1S") or `Class: Class Code`.str.contains("NS2023-S2S") or `Class: Class Code`.str.contains("NS2023-S3S")'
QueryMMMM = '`Class: Class Code`.str.contains("PN2023-S1S") or `Class: Class Code`.str.contains("PN2023-S2S") or `Class: Class Code`.str.contains("PN2023-S3S")'
QueryNNNN = '`Class: Class Code`.str.contains("PC2023-S1S") or `Class: Class Code`.str.contains("PC2023-S2S") or `Class: Class Code`.str.contains("PC2023-S3S")'
QueryOOOO = '`Class: Class Code`.str.contains("PG2023-S1S") or `Class: Class Code`.str.contains("PG2023-S2S") or `Class: Class Code`.str.contains("PG2023-S3S")'
QueryPPPP = '`Class: Class Code`.str.contains("RM2023-S1S") or `Class: Class Code`.str.contains("RM2023-S2S") or `Class: Class Code`.str.contains("RM2023-S3S")'
QueryQQQQ = '`Class: Class Code`.str.contains("SM2023-S1S") or `Class: Class Code`.str.contains("SM2023-S2S") or `Class: Class Code`.str.contains("SM2023-S3S")'
QueryRRRR = '`Class: Class Code`.str.contains("SW2023-S1S") or `Class: Class Code`.str.contains("SW2023-S2S") or `Class: Class Code`.str.contains("SW2023-S3S")'
QuerySSSS = '`Class: Class Code`.str.contains("SME2023-S1S") or `Class: Class Code`.str.contains("SME2023-S2S") or `Class: Class Code`.str.contains("SME2023-S3S")'
QueryTTTT = '`Class: Class Code`.str.contains("WC2023-S1S") or `Class: Class Code`.str.contains("WC2023-S2S") or `Class: Class Code`.str.contains("WC2023-S3S")'
QueryUUUU = '`Class: Class Code`.str.contains("WLC2023-S1S") or `Class: Class Code`.str.contains("WLC2023-S2S") or `Class: Class Code`.str.contains("WLC2023-S3S")'
QueryVVVV = '`Class: Class Code`.str.contains("ED2023-S1S") or `Class: Class Code`.str.contains("ED2023-S2S") or `Class: Class Code`.str.contains("ED2023-S3S")'
QueryWWWW = '`Class: Class Code`.str.contains("YS2023-S1S") or `Class: Class Code`.str.contains("YS2023-S2S") or `Class: Class Code`.str.contains("YS2023-S3S")'
# Add more if needed

# Check if a previous version of XLSX files exists
try:
    # Pending Conversion
    os.path.exists("./z_Final.xlsx")
    os.remove("./z_Final.xlsx")
    print("A previous version of z_Final detected! Deleting file...")
except:
    logging.info("Proceed to export data...")
    print("Proceed to export data...")

# Define the list of Excel files
scanned_excel_files = [
    'Scanned OAS Franchisee.xlsx',
    'Scanned OAS Franchisor.xlsx',
    'PO Scanned OAS.xlsx',
]

pending_excel_files = [
    'Pending Conversion Franchisee.xlsx',
    'Pending Conversion Franchisor.xlsx',
    'PO Pending Conversion.xlsx',
]

outstanding_excel_files = [
    'Outstanding Scans Franchisee.xlsx',
    'Outstanding Scans Franchisor.xlsx',
    'PO Outstanding Scans.xlsx',
]

converted_excel_files = [
    'Converted OAS Franchisee.xlsx',
    'Converted OAS Franchisor.xlsx',
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

abs_converted_excel_files = [
    'ABS Converted OAS Franchisee.xlsx',
    'ABS Converted OAS Franchisor.xlsx',
]

# Create a DataFrame to store the counts
file_counts = pd.DataFrame(columns=['File Name', 'Centre', 'Group', 'OAS Count', 'Type', 'Scan Status', 'Conversion Status'])

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

    # Dataframe Queries
    Query1A = df.query(QueryAA, engine = 'python')
    Query2A = df.query(QueryBB, engine = 'python')
    Query3A = df.query(QueryCC, engine = 'python')
    Query4A = df.query(QueryDD, engine = 'python')
    Query5A = df.query(QueryEE, engine = 'python')
    Query6A = df.query(QueryFF, engine = 'python')
    Query7A = df.query(QueryGG, engine = 'python')
    Query8A = df.query(QueryHH, engine = 'python')
    Query9A = df.query(QueryII, engine = 'python')
    Query10A = df.query(QueryJJ, engine = 'python')
    Query11A = df.query(QueryKK, engine = 'python')
    Query12A = df.query(QueryLL, engine = 'python')
    Query13A = df.query(QueryMM, engine = 'python')
    Query14A = df.query(QueryNN, engine = 'python')
    Query15A = df.query(QueryOO, engine = 'python')
    Query16A = df.query(QueryPP, engine = 'python')
    Query17A = df.query(QueryQQ, engine = 'python')
    Query18A = df.query(QueryRR, engine = 'python')
    Query19A = df.query(QuerySS, engine = 'python')
    Query20A = df.query(QueryTT, engine = 'python')
    Query21A = df.query(QueryUU, engine = 'python')
    Query22A = df.query(QueryVV, engine = 'python')
    Query23A = df.query(QueryWW, engine = 'python')

    # Dataframe Queries
    Query1B = df.query(QueryAAA, engine = 'python')
    Query2B = df.query(QueryBBB, engine = 'python')
    Query3B = df.query(QueryCCC, engine = 'python')
    Query4B = df.query(QueryDDD, engine = 'python')
    Query5B = df.query(QueryEEE, engine = 'python')
    Query6B = df.query(QueryFFF, engine = 'python')
    Query7B = df.query(QueryGGG, engine = 'python')
    Query8B = df.query(QueryHHH, engine = 'python')
    Query9B = df.query(QueryIII, engine = 'python')
    Query10B = df.query(QueryJJJ, engine = 'python')
    Query11B = df.query(QueryKKK, engine = 'python')
    Query12B = df.query(QueryLLL, engine = 'python')
    Query13B = df.query(QueryMMM, engine = 'python')
    Query14B = df.query(QueryNNN, engine = 'python')
    Query15B = df.query(QueryOOO, engine = 'python')
    Query16B = df.query(QueryPPP, engine = 'python')
    Query17B = df.query(QueryQQQ, engine = 'python')
    Query18B = df.query(QueryRRR, engine = 'python')
    Query19B = df.query(QuerySSS, engine = 'python')
    Query20B = df.query(QueryTTT, engine = 'python')
    Query21B = df.query(QueryUUU, engine = 'python')
    Query22B = df.query(QueryVVV, engine = 'python')
    Query23B = df.query(QueryWWW, engine = 'python')

    # Dataframe Queries
    Query1C = df.query(QueryAAAA, engine = 'python')
    Query2C = df.query(QueryBBBB, engine = 'python')
    Query3C = df.query(QueryCCCC, engine = 'python')
    Query4C = df.query(QueryDDDD, engine = 'python')
    Query5C = df.query(QueryEEEE, engine = 'python')
    Query6C = df.query(QueryFFFF, engine = 'python')
    Query7C = df.query(QueryGGGG, engine = 'python')
    Query8C = df.query(QueryHHHH, engine = 'python')
    Query9C = df.query(QueryIIII, engine = 'python')
    Query10C = df.query(QueryJJJJ, engine = 'python')
    Query11C = df.query(QueryKKKK, engine = 'python')
    Query12C = df.query(QueryLLLL, engine = 'python')
    Query13C = df.query(QueryMMMM, engine = 'python')
    Query14C = df.query(QueryNNNN, engine = 'python')
    Query15C = df.query(QueryOOOO, engine = 'python')
    Query16C = df.query(QueryPPPP, engine = 'python')
    Query17C = df.query(QueryQQQQ, engine = 'python')
    Query18C = df.query(QueryRRRR, engine = 'python')
    Query19C = df.query(QuerySSSS, engine = 'python')
    Query20C = df.query(QueryTTTT, engine = 'python')
    Query21C = df.query(QueryUUUU, engine = 'python')
    Query22C = df.query(QueryVVVV, engine = 'python')
    Query23C = df.query(QueryWWWW, engine = 'python')

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

    # Get the OAS Count
    CountQueryAA = len(Query1A)
    CountQueryBB = len(Query2A)
    CountQueryCC = len(Query3A)
    CountQueryDD = len(Query4A)
    CountQueryEE = len(Query5A)
    CountQueryFF = len(Query6A)
    CountQueryGG = len(Query7A)
    CountQueryHH = len(Query8A)
    CountQueryII = len(Query9A)
    CountQueryJJ = len(Query10A)
    CountQueryKK = len(Query11A)
    CountQueryLL = len(Query12A)
    CountQueryMM = len(Query13A)
    CountQueryNN = len(Query14A)
    CountQueryOO = len(Query15A)
    CountQueryPP = len(Query16A)
    CountQueryQQ = len(Query17A)
    CountQueryRR = len(Query18A)
    CountQuerySS = len(Query19A)
    CountQueryTT = len(Query20A)
    CountQueryUU = len(Query21A)
    CountQueryVV = len(Query22A)
    CountQueryWW = len(Query23A)

    # Get the OAS Count
    CountQueryAAA = len(Query1B)
    CountQueryBBB = len(Query2B)
    CountQueryCCC = len(Query3B)
    CountQueryDDD = len(Query4B)
    CountQueryEEE = len(Query5B)
    CountQueryFFF = len(Query6B)
    CountQueryGGG = len(Query7B)
    CountQueryHHH = len(Query8B)
    CountQueryIII = len(Query9B)
    CountQueryJJJ = len(Query10B)
    CountQueryKKK = len(Query11B)
    CountQueryLLL = len(Query12B)
    CountQueryMMM = len(Query13B)
    CountQueryNNN = len(Query14B)
    CountQueryOOO = len(Query15B)
    CountQueryPPP = len(Query16B)
    CountQueryQQQ = len(Query17B)
    CountQueryRRR = len(Query18B)
    CountQuerySSS = len(Query19B)
    CountQueryTTT = len(Query20B)
    CountQueryUUU = len(Query21B)
    CountQueryVVV = len(Query22B)
    CountQueryWWW = len(Query23B)

    # Get the OAS Count
    CountQueryAAAA = len(Query1C)
    CountQueryBBBB = len(Query2C)
    CountQueryCCCC = len(Query3C)
    CountQueryDDDD = len(Query4C)
    CountQueryEEEE = len(Query5C)
    CountQueryFFFF = len(Query6C)
    CountQueryGGGG = len(Query7C)
    CountQueryHHHH = len(Query8C)
    CountQueryIIII = len(Query9C)
    CountQueryJJJJ = len(Query10C)
    CountQueryKKKK = len(Query11C)
    CountQueryLLLL = len(Query12C)
    CountQueryMMMM = len(Query13C)
    CountQueryNNNN = len(Query14C)
    CountQueryOOOO = len(Query15C)
    CountQueryPPPP = len(Query16C)
    CountQueryQQQQ = len(Query17C)
    CountQueryRRRR = len(Query18C)
    CountQuerySSSS = len(Query19C)
    CountQueryTTTT = len(Query20C)
    CountQueryUUUU = len(Query21C)
    CountQueryVVVV = len(Query22C)
    CountQueryWWWW = len(Query23C)

    # Store the file name and OAS Count in the DataFrame
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupA, 'Group': Group1, 'OAS Count': CountQueryA, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupB, 'Group': Group1, 'OAS Count': CountQueryB, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupC, 'Group': Group1, 'OAS Count': CountQueryC, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupD, 'Group': Group1, 'OAS Count': CountQueryD, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupE, 'Group': Group1, 'OAS Count': CountQueryE, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupF, 'Group': Group1, 'OAS Count': CountQueryF, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupG, 'Group': Group1, 'OAS Count': CountQueryG, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupH, 'Group': Group1, 'OAS Count': CountQueryH, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupI, 'Group': Group1, 'OAS Count': CountQueryI, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupJ, 'Group': Group1, 'OAS Count': CountQueryJ, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupK, 'Group': Group1, 'OAS Count': CountQueryK, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupL, 'Group': Group1, 'OAS Count': CountQueryL, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupM, 'Group': Group1, 'OAS Count': CountQueryM, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupN, 'Group': Group1, 'OAS Count': CountQueryN, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupO, 'Group': Group1, 'OAS Count': CountQueryO, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupP, 'Group': Group1, 'OAS Count': CountQueryP, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupQ, 'Group': Group1, 'OAS Count': CountQueryQ, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupR, 'Group': Group1, 'OAS Count': CountQueryR, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupS, 'Group': Group1, 'OAS Count': CountQueryS, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupT, 'Group': Group1, 'OAS Count': CountQueryT, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupU, 'Group': Group1, 'OAS Count': CountQueryU, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupV, 'Group': Group1, 'OAS Count': CountQueryV, 'Type': Type2, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupW, 'Group': Group1, 'OAS Count': CountQueryW, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)

    # Store the file name and OAS Count in the DataFrame
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupA, 'Group': Group2, 'OAS Count': CountQueryAA, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupB, 'Group': Group2, 'OAS Count': CountQueryBB, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupC, 'Group': Group2, 'OAS Count': CountQueryCC, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupD, 'Group': Group2, 'OAS Count': CountQueryDD, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupE, 'Group': Group2, 'OAS Count': CountQueryEE, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupF, 'Group': Group2, 'OAS Count': CountQueryFF, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupG, 'Group': Group2, 'OAS Count': CountQueryGG, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupH, 'Group': Group2, 'OAS Count': CountQueryHH, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupI, 'Group': Group2, 'OAS Count': CountQueryII, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupJ, 'Group': Group2, 'OAS Count': CountQueryJJ, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupK, 'Group': Group2, 'OAS Count': CountQueryKK, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupL, 'Group': Group2, 'OAS Count': CountQueryLL, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupM, 'Group': Group2, 'OAS Count': CountQueryMM, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupN, 'Group': Group2, 'OAS Count': CountQueryNN, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupO, 'Group': Group2, 'OAS Count': CountQueryOO, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupP, 'Group': Group2, 'OAS Count': CountQueryPP, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupQ, 'Group': Group2, 'OAS Count': CountQueryQQ, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupR, 'Group': Group2, 'OAS Count': CountQueryRR, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupS, 'Group': Group2, 'OAS Count': CountQuerySS, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupT, 'Group': Group2, 'OAS Count': CountQueryTT, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupU, 'Group': Group2, 'OAS Count': CountQueryUU, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupV, 'Group': Group2, 'OAS Count': CountQueryVV, 'Type': Type2, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupW, 'Group': Group2, 'OAS Count': CountQueryWW, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)

    # Store the file name and OAS Count in the DataFrame
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupA, 'Group': Group3, 'OAS Count': CountQueryAAA, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupB, 'Group': Group3, 'OAS Count': CountQueryBBB, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupC, 'Group': Group3, 'OAS Count': CountQueryCCC, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupD, 'Group': Group3, 'OAS Count': CountQueryDDD, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupE, 'Group': Group3, 'OAS Count': CountQueryEEE, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupF, 'Group': Group3, 'OAS Count': CountQueryFFF, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupG, 'Group': Group3, 'OAS Count': CountQueryGGG, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupH, 'Group': Group3, 'OAS Count': CountQueryHHH, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupI, 'Group': Group3, 'OAS Count': CountQueryIII, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupJ, 'Group': Group3, 'OAS Count': CountQueryJJJ, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupK, 'Group': Group3, 'OAS Count': CountQueryKKK, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupL, 'Group': Group3, 'OAS Count': CountQueryLLL, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupM, 'Group': Group3, 'OAS Count': CountQueryMMM, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupN, 'Group': Group3, 'OAS Count': CountQueryNNN, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupO, 'Group': Group3, 'OAS Count': CountQueryOOO, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupP, 'Group': Group3, 'OAS Count': CountQueryPPP, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupQ, 'Group': Group3, 'OAS Count': CountQueryQQQ, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupR, 'Group': Group3, 'OAS Count': CountQueryRRR, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupS, 'Group': Group3, 'OAS Count': CountQuerySSS, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupT, 'Group': Group3, 'OAS Count': CountQueryTTT, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupU, 'Group': Group3, 'OAS Count': CountQueryUUU, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupV, 'Group': Group3, 'OAS Count': CountQueryVVV, 'Type': Type2, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupW, 'Group': Group3, 'OAS Count': CountQueryWWW, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)

    # Store the file name and OAS Count in the DataFrame
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupA, 'Group': Group4, 'OAS Count': CountQueryAAAA, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupB, 'Group': Group4, 'OAS Count': CountQueryBBBB, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupC, 'Group': Group4, 'OAS Count': CountQueryCCCC, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupD, 'Group': Group4, 'OAS Count': CountQueryDDDD, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupE, 'Group': Group4, 'OAS Count': CountQueryEEEE, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupF, 'Group': Group4, 'OAS Count': CountQueryFFFF, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupG, 'Group': Group4, 'OAS Count': CountQueryGGGG, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupH, 'Group': Group4, 'OAS Count': CountQueryHHHH, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupI, 'Group': Group4, 'OAS Count': CountQueryIIII, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupJ, 'Group': Group4, 'OAS Count': CountQueryJJJJ, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupK, 'Group': Group4, 'OAS Count': CountQueryKKKK, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupL, 'Group': Group4, 'OAS Count': CountQueryLLLL, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupM, 'Group': Group4, 'OAS Count': CountQueryMMMM, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupN, 'Group': Group4, 'OAS Count': CountQueryNNNN, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupO, 'Group': Group4, 'OAS Count': CountQueryOOOO, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupP, 'Group': Group4, 'OAS Count': CountQueryPPPP, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupQ, 'Group': Group4, 'OAS Count': CountQueryQQQQ, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupR, 'Group': Group4, 'OAS Count': CountQueryRRRR, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupS, 'Group': Group4, 'OAS Count': CountQuerySSSS, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupT, 'Group': Group4, 'OAS Count': CountQueryTTTT, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupU, 'Group': Group4, 'OAS Count': CountQueryUUUU, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupV, 'Group': Group4, 'OAS Count': CountQueryVVVV, 'Type': Type2, 'Scan Status': ScanStatusTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupW, 'Group': Group4, 'OAS Count': CountQueryWWWW, 'Type': Type1, 'Scan Status': ScanStatusTrue}, ignore_index=True)

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

    # Dataframe Queries
    Query1A = df.query(QueryAA, engine = 'python')
    Query2A = df.query(QueryBB, engine = 'python')
    Query3A = df.query(QueryCC, engine = 'python')
    Query4A = df.query(QueryDD, engine = 'python')
    Query5A = df.query(QueryEE, engine = 'python')
    Query6A = df.query(QueryFF, engine = 'python')
    Query7A = df.query(QueryGG, engine = 'python')
    Query8A = df.query(QueryHH, engine = 'python')
    Query9A = df.query(QueryII, engine = 'python')
    Query10A = df.query(QueryJJ, engine = 'python')
    Query11A = df.query(QueryKK, engine = 'python')
    Query12A = df.query(QueryLL, engine = 'python')
    Query13A = df.query(QueryMM, engine = 'python')
    Query14A = df.query(QueryNN, engine = 'python')
    Query15A = df.query(QueryOO, engine = 'python')
    Query16A = df.query(QueryPP, engine = 'python')
    Query17A = df.query(QueryQQ, engine = 'python')
    Query18A = df.query(QueryRR, engine = 'python')
    Query19A = df.query(QuerySS, engine = 'python')
    Query20A = df.query(QueryTT, engine = 'python')
    Query21A = df.query(QueryUU, engine = 'python')
    Query22A = df.query(QueryVV, engine = 'python')
    Query23A = df.query(QueryWW, engine = 'python')

    # Dataframe Queries
    Query1B = df.query(QueryAAA, engine = 'python')
    Query2B = df.query(QueryBBB, engine = 'python')
    Query3B = df.query(QueryCCC, engine = 'python')
    Query4B = df.query(QueryDDD, engine = 'python')
    Query5B = df.query(QueryEEE, engine = 'python')
    Query6B = df.query(QueryFFF, engine = 'python')
    Query7B = df.query(QueryGGG, engine = 'python')
    Query8B = df.query(QueryHHH, engine = 'python')
    Query9B = df.query(QueryIII, engine = 'python')
    Query10B = df.query(QueryJJJ, engine = 'python')
    Query11B = df.query(QueryKKK, engine = 'python')
    Query12B = df.query(QueryLLL, engine = 'python')
    Query13B = df.query(QueryMMM, engine = 'python')
    Query14B = df.query(QueryNNN, engine = 'python')
    Query15B = df.query(QueryOOO, engine = 'python')
    Query16B = df.query(QueryPPP, engine = 'python')
    Query17B = df.query(QueryQQQ, engine = 'python')
    Query18B = df.query(QueryRRR, engine = 'python')
    Query19B = df.query(QuerySSS, engine = 'python')
    Query20B = df.query(QueryTTT, engine = 'python')
    Query21B = df.query(QueryUUU, engine = 'python')
    Query22B = df.query(QueryVVV, engine = 'python')
    Query23B = df.query(QueryWWW, engine = 'python')

    # Dataframe Queries
    Query1C = df.query(QueryAAAA, engine = 'python')
    Query2C = df.query(QueryBBBB, engine = 'python')
    Query3C = df.query(QueryCCCC, engine = 'python')
    Query4C = df.query(QueryDDDD, engine = 'python')
    Query5C = df.query(QueryEEEE, engine = 'python')
    Query6C = df.query(QueryFFFF, engine = 'python')
    Query7C = df.query(QueryGGGG, engine = 'python')
    Query8C = df.query(QueryHHHH, engine = 'python')
    Query9C = df.query(QueryIIII, engine = 'python')
    Query10C = df.query(QueryJJJJ, engine = 'python')
    Query11C = df.query(QueryKKKK, engine = 'python')
    Query12C = df.query(QueryLLLL, engine = 'python')
    Query13C = df.query(QueryMMMM, engine = 'python')
    Query14C = df.query(QueryNNNN, engine = 'python')
    Query15C = df.query(QueryOOOO, engine = 'python')
    Query16C = df.query(QueryPPPP, engine = 'python')
    Query17C = df.query(QueryQQQQ, engine = 'python')
    Query18C = df.query(QueryRRRR, engine = 'python')
    Query19C = df.query(QuerySSSS, engine = 'python')
    Query20C = df.query(QueryTTTT, engine = 'python')
    Query21C = df.query(QueryUUUU, engine = 'python')
    Query22C = df.query(QueryVVVV, engine = 'python')
    Query23C = df.query(QueryWWWW, engine = 'python')

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

    # Get the OAS Count
    CountQueryAA = len(Query1A)
    CountQueryBB = len(Query2A)
    CountQueryCC = len(Query3A)
    CountQueryDD = len(Query4A)
    CountQueryEE = len(Query5A)
    CountQueryFF = len(Query6A)
    CountQueryGG = len(Query7A)
    CountQueryHH = len(Query8A)
    CountQueryII = len(Query9A)
    CountQueryJJ = len(Query10A)
    CountQueryKK = len(Query11A)
    CountQueryLL = len(Query12A)
    CountQueryMM = len(Query13A)
    CountQueryNN = len(Query14A)
    CountQueryOO = len(Query15A)
    CountQueryPP = len(Query16A)
    CountQueryQQ = len(Query17A)
    CountQueryRR = len(Query18A)
    CountQuerySS = len(Query19A)
    CountQueryTT = len(Query20A)
    CountQueryUU = len(Query21A)
    CountQueryVV = len(Query22A)
    CountQueryWW = len(Query23A)

    # Get the OAS Count
    CountQueryAAA = len(Query1B)
    CountQueryBBB = len(Query2B)
    CountQueryCCC = len(Query3B)
    CountQueryDDD = len(Query4B)
    CountQueryEEE = len(Query5B)
    CountQueryFFF = len(Query6B)
    CountQueryGGG = len(Query7B)
    CountQueryHHH = len(Query8B)
    CountQueryIII = len(Query9B)
    CountQueryJJJ = len(Query10B)
    CountQueryKKK = len(Query11B)
    CountQueryLLL = len(Query12B)
    CountQueryMMM = len(Query13B)
    CountQueryNNN = len(Query14B)
    CountQueryOOO = len(Query15B)
    CountQueryPPP = len(Query16B)
    CountQueryQQQ = len(Query17B)
    CountQueryRRR = len(Query18B)
    CountQuerySSS = len(Query19B)
    CountQueryTTT = len(Query20B)
    CountQueryUUU = len(Query21B)
    CountQueryVVV = len(Query22B)
    CountQueryWWW = len(Query23B)

    # Get the OAS Count
    CountQueryAAAA = len(Query1C)
    CountQueryBBBB = len(Query2C)
    CountQueryCCCC = len(Query3C)
    CountQueryDDDD = len(Query4C)
    CountQueryEEEE = len(Query5C)
    CountQueryFFFF = len(Query6C)
    CountQueryGGGG = len(Query7C)
    CountQueryHHHH = len(Query8C)
    CountQueryIIII = len(Query9C)
    CountQueryJJJJ = len(Query10C)
    CountQueryKKKK = len(Query11C)
    CountQueryLLLL = len(Query12C)
    CountQueryMMMM = len(Query13C)
    CountQueryNNNN = len(Query14C)
    CountQueryOOOO = len(Query15C)
    CountQueryPPPP = len(Query16C)
    CountQueryQQQQ = len(Query17C)
    CountQueryRRRR = len(Query18C)
    CountQuerySSSS = len(Query19C)
    CountQueryTTTT = len(Query20C)
    CountQueryUUUU = len(Query21C)
    CountQueryVVVV = len(Query22C)
    CountQueryWWWW = len(Query23C)

    # Store the file name and OAS Count in the DataFrame
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupA, 'Group': Group1, 'OAS Count': CountQueryA, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupB, 'Group': Group1, 'OAS Count': CountQueryB, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupC, 'Group': Group1, 'OAS Count': CountQueryC, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupD, 'Group': Group1, 'OAS Count': CountQueryD, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupE, 'Group': Group1, 'OAS Count': CountQueryE, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupF, 'Group': Group1, 'OAS Count': CountQueryF, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupG, 'Group': Group1, 'OAS Count': CountQueryG, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupH, 'Group': Group1, 'OAS Count': CountQueryH, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupI, 'Group': Group1, 'OAS Count': CountQueryI, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupJ, 'Group': Group1, 'OAS Count': CountQueryJ, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupK, 'Group': Group1, 'OAS Count': CountQueryK, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupL, 'Group': Group1, 'OAS Count': CountQueryL, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupM, 'Group': Group1, 'OAS Count': CountQueryM, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupN, 'Group': Group1, 'OAS Count': CountQueryN, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupO, 'Group': Group1, 'OAS Count': CountQueryO, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupP, 'Group': Group1, 'OAS Count': CountQueryP, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupQ, 'Group': Group1, 'OAS Count': CountQueryQ, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupR, 'Group': Group1, 'OAS Count': CountQueryR, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupS, 'Group': Group1, 'OAS Count': CountQueryS, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupT, 'Group': Group1, 'OAS Count': CountQueryT, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupU, 'Group': Group1, 'OAS Count': CountQueryU, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupV, 'Group': Group1, 'OAS Count': CountQueryV, 'Type': Type2, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupW, 'Group': Group1, 'OAS Count': CountQueryW, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)

    # Store the file name and OAS Count in the DataFrame
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupA, 'Group': Group2, 'OAS Count': CountQueryAA, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupB, 'Group': Group2, 'OAS Count': CountQueryBB, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupC, 'Group': Group2, 'OAS Count': CountQueryCC, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupD, 'Group': Group2, 'OAS Count': CountQueryDD, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupE, 'Group': Group2, 'OAS Count': CountQueryEE, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupF, 'Group': Group2, 'OAS Count': CountQueryFF, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupG, 'Group': Group2, 'OAS Count': CountQueryGG, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupH, 'Group': Group2, 'OAS Count': CountQueryHH, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupI, 'Group': Group2, 'OAS Count': CountQueryII, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupJ, 'Group': Group2, 'OAS Count': CountQueryJJ, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupK, 'Group': Group2, 'OAS Count': CountQueryKK, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupL, 'Group': Group2, 'OAS Count': CountQueryLL, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupM, 'Group': Group2, 'OAS Count': CountQueryMM, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupN, 'Group': Group2, 'OAS Count': CountQueryNN, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupO, 'Group': Group2, 'OAS Count': CountQueryOO, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupP, 'Group': Group2, 'OAS Count': CountQueryPP, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupQ, 'Group': Group2, 'OAS Count': CountQueryQQ, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupR, 'Group': Group2, 'OAS Count': CountQueryRR, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupS, 'Group': Group2, 'OAS Count': CountQuerySS, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupT, 'Group': Group2, 'OAS Count': CountQueryTT, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupU, 'Group': Group2, 'OAS Count': CountQueryUU, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupV, 'Group': Group2, 'OAS Count': CountQueryVV, 'Type': Type2, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupW, 'Group': Group2, 'OAS Count': CountQueryWW, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)

    # Store the file name and OAS Count in the DataFrame
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupA, 'Group': Group3, 'OAS Count': CountQueryAAA, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupB, 'Group': Group3, 'OAS Count': CountQueryBBB, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupC, 'Group': Group3, 'OAS Count': CountQueryCCC, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupD, 'Group': Group3, 'OAS Count': CountQueryDDD, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupE, 'Group': Group3, 'OAS Count': CountQueryEEE, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupF, 'Group': Group3, 'OAS Count': CountQueryFFF, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupG, 'Group': Group3, 'OAS Count': CountQueryGGG, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupH, 'Group': Group3, 'OAS Count': CountQueryHHH, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupI, 'Group': Group3, 'OAS Count': CountQueryIII, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupJ, 'Group': Group3, 'OAS Count': CountQueryJJJ, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupK, 'Group': Group3, 'OAS Count': CountQueryKKK, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupL, 'Group': Group3, 'OAS Count': CountQueryLLL, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupM, 'Group': Group3, 'OAS Count': CountQueryMMM, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupN, 'Group': Group3, 'OAS Count': CountQueryNNN, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupO, 'Group': Group3, 'OAS Count': CountQueryOOO, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupP, 'Group': Group3, 'OAS Count': CountQueryPPP, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupQ, 'Group': Group3, 'OAS Count': CountQueryQQQ, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupR, 'Group': Group3, 'OAS Count': CountQueryRRR, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupS, 'Group': Group3, 'OAS Count': CountQuerySSS, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupT, 'Group': Group3, 'OAS Count': CountQueryTTT, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupU, 'Group': Group3, 'OAS Count': CountQueryUUU, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupV, 'Group': Group3, 'OAS Count': CountQueryVVV, 'Type': Type2, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupW, 'Group': Group3, 'OAS Count': CountQueryWWW, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)

    # Store the file name and OAS Count in the DataFrame
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupA, 'Group': Group4, 'OAS Count': CountQueryAAAA, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupB, 'Group': Group4, 'OAS Count': CountQueryBBBB, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupC, 'Group': Group4, 'OAS Count': CountQueryCCCC, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupD, 'Group': Group4, 'OAS Count': CountQueryDDDD, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupE, 'Group': Group4, 'OAS Count': CountQueryEEEE, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupF, 'Group': Group4, 'OAS Count': CountQueryFFFF, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupG, 'Group': Group4, 'OAS Count': CountQueryGGGG, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupH, 'Group': Group4, 'OAS Count': CountQueryHHHH, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupI, 'Group': Group4, 'OAS Count': CountQueryIIII, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupJ, 'Group': Group4, 'OAS Count': CountQueryJJJJ, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupK, 'Group': Group4, 'OAS Count': CountQueryKKKK, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupL, 'Group': Group4, 'OAS Count': CountQueryLLLL, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupM, 'Group': Group4, 'OAS Count': CountQueryMMMM, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupN, 'Group': Group4, 'OAS Count': CountQueryNNNN, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupO, 'Group': Group4, 'OAS Count': CountQueryOOOO, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupP, 'Group': Group4, 'OAS Count': CountQueryPPPP, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupQ, 'Group': Group4, 'OAS Count': CountQueryQQQQ, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupR, 'Group': Group4, 'OAS Count': CountQueryRRRR, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupS, 'Group': Group4, 'OAS Count': CountQuerySSSS, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupT, 'Group': Group4, 'OAS Count': CountQueryTTTT, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupU, 'Group': Group4, 'OAS Count': CountQueryUUUU, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupV, 'Group': Group4, 'OAS Count': CountQueryVVVV, 'Type': Type2, 'Conversion Status': ConStatsFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupW, 'Group': Group4, 'OAS Count': CountQueryWWWW, 'Type': Type1, 'Conversion Status': ConStatsFalse}, ignore_index=True)

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

    # Dataframe Queries
    Query1A = df.query(QueryAA, engine = 'python')
    Query2A = df.query(QueryBB, engine = 'python')
    Query3A = df.query(QueryCC, engine = 'python')
    Query4A = df.query(QueryDD, engine = 'python')
    Query5A = df.query(QueryEE, engine = 'python')
    Query6A = df.query(QueryFF, engine = 'python')
    Query7A = df.query(QueryGG, engine = 'python')
    Query8A = df.query(QueryHH, engine = 'python')
    Query9A = df.query(QueryII, engine = 'python')
    Query10A = df.query(QueryJJ, engine = 'python')
    Query11A = df.query(QueryKK, engine = 'python')
    Query12A = df.query(QueryLL, engine = 'python')
    Query13A = df.query(QueryMM, engine = 'python')
    Query14A = df.query(QueryNN, engine = 'python')
    Query15A = df.query(QueryOO, engine = 'python')
    Query16A = df.query(QueryPP, engine = 'python')
    Query17A = df.query(QueryQQ, engine = 'python')
    Query18A = df.query(QueryRR, engine = 'python')
    Query19A = df.query(QuerySS, engine = 'python')
    Query20A = df.query(QueryTT, engine = 'python')
    Query21A = df.query(QueryUU, engine = 'python')
    Query22A = df.query(QueryVV, engine = 'python')
    Query23A = df.query(QueryWW, engine = 'python')

    # Dataframe Queries
    Query1B = df.query(QueryAAA, engine = 'python')
    Query2B = df.query(QueryBBB, engine = 'python')
    Query3B = df.query(QueryCCC, engine = 'python')
    Query4B = df.query(QueryDDD, engine = 'python')
    Query5B = df.query(QueryEEE, engine = 'python')
    Query6B = df.query(QueryFFF, engine = 'python')
    Query7B = df.query(QueryGGG, engine = 'python')
    Query8B = df.query(QueryHHH, engine = 'python')
    Query9B = df.query(QueryIII, engine = 'python')
    Query10B = df.query(QueryJJJ, engine = 'python')
    Query11B = df.query(QueryKKK, engine = 'python')
    Query12B = df.query(QueryLLL, engine = 'python')
    Query13B = df.query(QueryMMM, engine = 'python')
    Query14B = df.query(QueryNNN, engine = 'python')
    Query15B = df.query(QueryOOO, engine = 'python')
    Query16B = df.query(QueryPPP, engine = 'python')
    Query17B = df.query(QueryQQQ, engine = 'python')
    Query18B = df.query(QueryRRR, engine = 'python')
    Query19B = df.query(QuerySSS, engine = 'python')
    Query20B = df.query(QueryTTT, engine = 'python')
    Query21B = df.query(QueryUUU, engine = 'python')
    Query22B = df.query(QueryVVV, engine = 'python')
    Query23B = df.query(QueryWWW, engine = 'python')

    # Dataframe Queries
    Query1C = df.query(QueryAAAA, engine = 'python')
    Query2C = df.query(QueryBBBB, engine = 'python')
    Query3C = df.query(QueryCCCC, engine = 'python')
    Query4C = df.query(QueryDDDD, engine = 'python')
    Query5C = df.query(QueryEEEE, engine = 'python')
    Query6C = df.query(QueryFFFF, engine = 'python')
    Query7C = df.query(QueryGGGG, engine = 'python')
    Query8C = df.query(QueryHHHH, engine = 'python')
    Query9C = df.query(QueryIIII, engine = 'python')
    Query10C = df.query(QueryJJJJ, engine = 'python')
    Query11C = df.query(QueryKKKK, engine = 'python')
    Query12C = df.query(QueryLLLL, engine = 'python')
    Query13C = df.query(QueryMMMM, engine = 'python')
    Query14C = df.query(QueryNNNN, engine = 'python')
    Query15C = df.query(QueryOOOO, engine = 'python')
    Query16C = df.query(QueryPPPP, engine = 'python')
    Query17C = df.query(QueryQQQQ, engine = 'python')
    Query18C = df.query(QueryRRRR, engine = 'python')
    Query19C = df.query(QuerySSSS, engine = 'python')
    Query20C = df.query(QueryTTTT, engine = 'python')
    Query21C = df.query(QueryUUUU, engine = 'python')
    Query22C = df.query(QueryVVVV, engine = 'python')
    Query23C = df.query(QueryWWWW, engine = 'python')

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

    # Get the OAS Count
    CountQueryAA = len(Query1A)
    CountQueryBB = len(Query2A)
    CountQueryCC = len(Query3A)
    CountQueryDD = len(Query4A)
    CountQueryEE = len(Query5A)
    CountQueryFF = len(Query6A)
    CountQueryGG = len(Query7A)
    CountQueryHH = len(Query8A)
    CountQueryII = len(Query9A)
    CountQueryJJ = len(Query10A)
    CountQueryKK = len(Query11A)
    CountQueryLL = len(Query12A)
    CountQueryMM = len(Query13A)
    CountQueryNN = len(Query14A)
    CountQueryOO = len(Query15A)
    CountQueryPP = len(Query16A)
    CountQueryQQ = len(Query17A)
    CountQueryRR = len(Query18A)
    CountQuerySS = len(Query19A)
    CountQueryTT = len(Query20A)
    CountQueryUU = len(Query21A)
    CountQueryVV = len(Query22A)
    CountQueryWW = len(Query23A)

    # Get the OAS Count
    CountQueryAAA = len(Query1B)
    CountQueryBBB = len(Query2B)
    CountQueryCCC = len(Query3B)
    CountQueryDDD = len(Query4B)
    CountQueryEEE = len(Query5B)
    CountQueryFFF = len(Query6B)
    CountQueryGGG = len(Query7B)
    CountQueryHHH = len(Query8B)
    CountQueryIII = len(Query9B)
    CountQueryJJJ = len(Query10B)
    CountQueryKKK = len(Query11B)
    CountQueryLLL = len(Query12B)
    CountQueryMMM = len(Query13B)
    CountQueryNNN = len(Query14B)
    CountQueryOOO = len(Query15B)
    CountQueryPPP = len(Query16B)
    CountQueryQQQ = len(Query17B)
    CountQueryRRR = len(Query18B)
    CountQuerySSS = len(Query19B)
    CountQueryTTT = len(Query20B)
    CountQueryUUU = len(Query21B)
    CountQueryVVV = len(Query22B)
    CountQueryWWW = len(Query23B)

    # Get the OAS Count
    CountQueryAAAA = len(Query1C)
    CountQueryBBBB = len(Query2C)
    CountQueryCCCC = len(Query3C)
    CountQueryDDDD = len(Query4C)
    CountQueryEEEE = len(Query5C)
    CountQueryFFFF = len(Query6C)
    CountQueryGGGG = len(Query7C)
    CountQueryHHHH = len(Query8C)
    CountQueryIIII = len(Query9C)
    CountQueryJJJJ = len(Query10C)
    CountQueryKKKK = len(Query11C)
    CountQueryLLLL = len(Query12C)
    CountQueryMMMM = len(Query13C)
    CountQueryNNNN = len(Query14C)
    CountQueryOOOO = len(Query15C)
    CountQueryPPPP = len(Query16C)
    CountQueryQQQQ = len(Query17C)
    CountQueryRRRR = len(Query18C)
    CountQuerySSSS = len(Query19C)
    CountQueryTTTT = len(Query20C)
    CountQueryUUUU = len(Query21C)
    CountQueryVVVV = len(Query22C)
    CountQueryWWWW = len(Query23C)

    # Store the file name and OAS Count in the DataFrame
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupA, 'Group': Group1, 'OAS Count': CountQueryA, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupB, 'Group': Group1, 'OAS Count': CountQueryB, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupC, 'Group': Group1, 'OAS Count': CountQueryC, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupD, 'Group': Group1, 'OAS Count': CountQueryD, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupE, 'Group': Group1, 'OAS Count': CountQueryE, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupF, 'Group': Group1, 'OAS Count': CountQueryF, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupG, 'Group': Group1, 'OAS Count': CountQueryG, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupH, 'Group': Group1, 'OAS Count': CountQueryH, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupI, 'Group': Group1, 'OAS Count': CountQueryI, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupJ, 'Group': Group1, 'OAS Count': CountQueryJ, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupK, 'Group': Group1, 'OAS Count': CountQueryK, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupL, 'Group': Group1, 'OAS Count': CountQueryL, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupM, 'Group': Group1, 'OAS Count': CountQueryM, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupN, 'Group': Group1, 'OAS Count': CountQueryN, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupO, 'Group': Group1, 'OAS Count': CountQueryO, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupP, 'Group': Group1, 'OAS Count': CountQueryP, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupQ, 'Group': Group1, 'OAS Count': CountQueryQ, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupR, 'Group': Group1, 'OAS Count': CountQueryR, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupS, 'Group': Group1, 'OAS Count': CountQueryS, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupT, 'Group': Group1, 'OAS Count': CountQueryT, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupU, 'Group': Group1, 'OAS Count': CountQueryU, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupV, 'Group': Group1, 'OAS Count': CountQueryV, 'Type': Type2, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupW, 'Group': Group1, 'OAS Count': CountQueryW, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)

    # Store the file name and OAS Count in the DataFrame
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupA, 'Group': Group2, 'OAS Count': CountQueryAA, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupB, 'Group': Group2, 'OAS Count': CountQueryBB, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupC, 'Group': Group2, 'OAS Count': CountQueryCC, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupD, 'Group': Group2, 'OAS Count': CountQueryDD, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupE, 'Group': Group2, 'OAS Count': CountQueryEE, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupF, 'Group': Group2, 'OAS Count': CountQueryFF, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupG, 'Group': Group2, 'OAS Count': CountQueryGG, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupH, 'Group': Group2, 'OAS Count': CountQueryHH, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupI, 'Group': Group2, 'OAS Count': CountQueryII, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupJ, 'Group': Group2, 'OAS Count': CountQueryJJ, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupK, 'Group': Group2, 'OAS Count': CountQueryKK, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupL, 'Group': Group2, 'OAS Count': CountQueryLL, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupM, 'Group': Group2, 'OAS Count': CountQueryMM, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupN, 'Group': Group2, 'OAS Count': CountQueryNN, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupO, 'Group': Group2, 'OAS Count': CountQueryOO, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupP, 'Group': Group2, 'OAS Count': CountQueryPP, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupQ, 'Group': Group2, 'OAS Count': CountQueryQQ, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupR, 'Group': Group2, 'OAS Count': CountQueryRR, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupS, 'Group': Group2, 'OAS Count': CountQuerySS, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupT, 'Group': Group2, 'OAS Count': CountQueryTT, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupU, 'Group': Group2, 'OAS Count': CountQueryUU, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupV, 'Group': Group2, 'OAS Count': CountQueryVV, 'Type': Type2, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupW, 'Group': Group2, 'OAS Count': CountQueryWW, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)

    # Store the file name and OAS Count in the DataFrame
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupA, 'Group': Group3, 'OAS Count': CountQueryAAA, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupB, 'Group': Group3, 'OAS Count': CountQueryBBB, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupC, 'Group': Group3, 'OAS Count': CountQueryCCC, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupD, 'Group': Group3, 'OAS Count': CountQueryDDD, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupE, 'Group': Group3, 'OAS Count': CountQueryEEE, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupF, 'Group': Group3, 'OAS Count': CountQueryFFF, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupG, 'Group': Group3, 'OAS Count': CountQueryGGG, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupH, 'Group': Group3, 'OAS Count': CountQueryHHH, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupI, 'Group': Group3, 'OAS Count': CountQueryIII, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupJ, 'Group': Group3, 'OAS Count': CountQueryJJJ, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupK, 'Group': Group3, 'OAS Count': CountQueryKKK, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupL, 'Group': Group3, 'OAS Count': CountQueryLLL, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupM, 'Group': Group3, 'OAS Count': CountQueryMMM, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupN, 'Group': Group3, 'OAS Count': CountQueryNNN, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupO, 'Group': Group3, 'OAS Count': CountQueryOOO, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupP, 'Group': Group3, 'OAS Count': CountQueryPPP, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupQ, 'Group': Group3, 'OAS Count': CountQueryQQQ, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupR, 'Group': Group3, 'OAS Count': CountQueryRRR, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupS, 'Group': Group3, 'OAS Count': CountQuerySSS, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupT, 'Group': Group3, 'OAS Count': CountQueryTTT, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupU, 'Group': Group3, 'OAS Count': CountQueryUUU, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupV, 'Group': Group3, 'OAS Count': CountQueryVVV, 'Type': Type2, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupW, 'Group': Group3, 'OAS Count': CountQueryWWW, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)

    # Store the file name and OAS Count in the DataFrame
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupA, 'Group': Group4, 'OAS Count': CountQueryAAAA, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupB, 'Group': Group4, 'OAS Count': CountQueryBBBB, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupC, 'Group': Group4, 'OAS Count': CountQueryCCCC, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupD, 'Group': Group4, 'OAS Count': CountQueryDDDD, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupE, 'Group': Group4, 'OAS Count': CountQueryEEEE, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupF, 'Group': Group4, 'OAS Count': CountQueryFFFF, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupG, 'Group': Group4, 'OAS Count': CountQueryGGGG, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupH, 'Group': Group4, 'OAS Count': CountQueryHHHH, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupI, 'Group': Group4, 'OAS Count': CountQueryIIII, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupJ, 'Group': Group4, 'OAS Count': CountQueryJJJJ, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupK, 'Group': Group4, 'OAS Count': CountQueryKKKK, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupL, 'Group': Group4, 'OAS Count': CountQueryLLLL, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupM, 'Group': Group4, 'OAS Count': CountQueryMMMM, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupN, 'Group': Group4, 'OAS Count': CountQueryNNNN, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupO, 'Group': Group4, 'OAS Count': CountQueryOOOO, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupP, 'Group': Group4, 'OAS Count': CountQueryPPPP, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupQ, 'Group': Group4, 'OAS Count': CountQueryQQQQ, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupR, 'Group': Group4, 'OAS Count': CountQueryRRRR, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupS, 'Group': Group4, 'OAS Count': CountQuerySSSS, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupT, 'Group': Group4, 'OAS Count': CountQueryTTTT, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupU, 'Group': Group4, 'OAS Count': CountQueryUUUU, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupV, 'Group': Group4, 'OAS Count': CountQueryVVVV, 'Type': Type2, 'Scan Status': ScanStatusFalse}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupW, 'Group': Group4, 'OAS Count': CountQueryWWWW, 'Type': Type1, 'Scan Status': ScanStatusFalse}, ignore_index=True)

for file in converted_excel_files:
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

    # Dataframe Queries
    Query1A = df.query(QueryAA, engine = 'python')
    Query2A = df.query(QueryBB, engine = 'python')
    Query3A = df.query(QueryCC, engine = 'python')
    Query4A = df.query(QueryDD, engine = 'python')
    Query5A = df.query(QueryEE, engine = 'python')
    Query6A = df.query(QueryFF, engine = 'python')
    Query7A = df.query(QueryGG, engine = 'python')
    Query8A = df.query(QueryHH, engine = 'python')
    Query9A = df.query(QueryII, engine = 'python')
    Query10A = df.query(QueryJJ, engine = 'python')
    Query11A = df.query(QueryKK, engine = 'python')
    Query12A = df.query(QueryLL, engine = 'python')
    Query13A = df.query(QueryMM, engine = 'python')
    Query14A = df.query(QueryNN, engine = 'python')
    Query15A = df.query(QueryOO, engine = 'python')
    Query16A = df.query(QueryPP, engine = 'python')
    Query17A = df.query(QueryQQ, engine = 'python')
    Query18A = df.query(QueryRR, engine = 'python')
    Query19A = df.query(QuerySS, engine = 'python')
    Query20A = df.query(QueryTT, engine = 'python')
    Query21A = df.query(QueryUU, engine = 'python')
    Query22A = df.query(QueryVV, engine = 'python')
    Query23A = df.query(QueryWW, engine = 'python')

    # Dataframe Queries
    Query1B = df.query(QueryAAA, engine = 'python')
    Query2B = df.query(QueryBBB, engine = 'python')
    Query3B = df.query(QueryCCC, engine = 'python')
    Query4B = df.query(QueryDDD, engine = 'python')
    Query5B = df.query(QueryEEE, engine = 'python')
    Query6B = df.query(QueryFFF, engine = 'python')
    Query7B = df.query(QueryGGG, engine = 'python')
    Query8B = df.query(QueryHHH, engine = 'python')
    Query9B = df.query(QueryIII, engine = 'python')
    Query10B = df.query(QueryJJJ, engine = 'python')
    Query11B = df.query(QueryKKK, engine = 'python')
    Query12B = df.query(QueryLLL, engine = 'python')
    Query13B = df.query(QueryMMM, engine = 'python')
    Query14B = df.query(QueryNNN, engine = 'python')
    Query15B = df.query(QueryOOO, engine = 'python')
    Query16B = df.query(QueryPPP, engine = 'python')
    Query17B = df.query(QueryQQQ, engine = 'python')
    Query18B = df.query(QueryRRR, engine = 'python')
    Query19B = df.query(QuerySSS, engine = 'python')
    Query20B = df.query(QueryTTT, engine = 'python')
    Query21B = df.query(QueryUUU, engine = 'python')
    Query22B = df.query(QueryVVV, engine = 'python')
    Query23B = df.query(QueryWWW, engine = 'python')

    # Dataframe Queries
    Query1C = df.query(QueryAAAA, engine = 'python')
    Query2C = df.query(QueryBBBB, engine = 'python')
    Query3C = df.query(QueryCCCC, engine = 'python')
    Query4C = df.query(QueryDDDD, engine = 'python')
    Query5C = df.query(QueryEEEE, engine = 'python')
    Query6C = df.query(QueryFFFF, engine = 'python')
    Query7C = df.query(QueryGGGG, engine = 'python')
    Query8C = df.query(QueryHHHH, engine = 'python')
    Query9C = df.query(QueryIIII, engine = 'python')
    Query10C = df.query(QueryJJJJ, engine = 'python')
    Query11C = df.query(QueryKKKK, engine = 'python')
    Query12C = df.query(QueryLLLL, engine = 'python')
    Query13C = df.query(QueryMMMM, engine = 'python')
    Query14C = df.query(QueryNNNN, engine = 'python')
    Query15C = df.query(QueryOOOO, engine = 'python')
    Query16C = df.query(QueryPPPP, engine = 'python')
    Query17C = df.query(QueryQQQQ, engine = 'python')
    Query18C = df.query(QueryRRRR, engine = 'python')
    Query19C = df.query(QuerySSSS, engine = 'python')
    Query20C = df.query(QueryTTTT, engine = 'python')
    Query21C = df.query(QueryUUUU, engine = 'python')
    Query22C = df.query(QueryVVVV, engine = 'python')
    Query23C = df.query(QueryWWWW, engine = 'python')

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

    # Get the OAS Count
    CountQueryAA = len(Query1A)
    CountQueryBB = len(Query2A)
    CountQueryCC = len(Query3A)
    CountQueryDD = len(Query4A)
    CountQueryEE = len(Query5A)
    CountQueryFF = len(Query6A)
    CountQueryGG = len(Query7A)
    CountQueryHH = len(Query8A)
    CountQueryII = len(Query9A)
    CountQueryJJ = len(Query10A)
    CountQueryKK = len(Query11A)
    CountQueryLL = len(Query12A)
    CountQueryMM = len(Query13A)
    CountQueryNN = len(Query14A)
    CountQueryOO = len(Query15A)
    CountQueryPP = len(Query16A)
    CountQueryQQ = len(Query17A)
    CountQueryRR = len(Query18A)
    CountQuerySS = len(Query19A)
    CountQueryTT = len(Query20A)
    CountQueryUU = len(Query21A)
    CountQueryVV = len(Query22A)
    CountQueryWW = len(Query23A)

    # Get the OAS Count
    CountQueryAAA = len(Query1B)
    CountQueryBBB = len(Query2B)
    CountQueryCCC = len(Query3B)
    CountQueryDDD = len(Query4B)
    CountQueryEEE = len(Query5B)
    CountQueryFFF = len(Query6B)
    CountQueryGGG = len(Query7B)
    CountQueryHHH = len(Query8B)
    CountQueryIII = len(Query9B)
    CountQueryJJJ = len(Query10B)
    CountQueryKKK = len(Query11B)
    CountQueryLLL = len(Query12B)
    CountQueryMMM = len(Query13B)
    CountQueryNNN = len(Query14B)
    CountQueryOOO = len(Query15B)
    CountQueryPPP = len(Query16B)
    CountQueryQQQ = len(Query17B)
    CountQueryRRR = len(Query18B)
    CountQuerySSS = len(Query19B)
    CountQueryTTT = len(Query20B)
    CountQueryUUU = len(Query21B)
    CountQueryVVV = len(Query22B)
    CountQueryWWW = len(Query23B)

    # Get the OAS Count
    CountQueryAAAA = len(Query1C)
    CountQueryBBBB = len(Query2C)
    CountQueryCCCC = len(Query3C)
    CountQueryDDDD = len(Query4C)
    CountQueryEEEE = len(Query5C)
    CountQueryFFFF = len(Query6C)
    CountQueryGGGG = len(Query7C)
    CountQueryHHHH = len(Query8C)
    CountQueryIIII = len(Query9C)
    CountQueryJJJJ = len(Query10C)
    CountQueryKKKK = len(Query11C)
    CountQueryLLLL = len(Query12C)
    CountQueryMMMM = len(Query13C)
    CountQueryNNNN = len(Query14C)
    CountQueryOOOO = len(Query15C)
    CountQueryPPPP = len(Query16C)
    CountQueryQQQQ = len(Query17C)
    CountQueryRRRR = len(Query18C)
    CountQuerySSSS = len(Query19C)
    CountQueryTTTT = len(Query20C)
    CountQueryUUUU = len(Query21C)
    CountQueryVVVV = len(Query22C)
    CountQueryWWWW = len(Query23C)

    # Store the file name and OAS Count in the DataFrame
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupA, 'Group': Group1, 'OAS Count': CountQueryA, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupB, 'Group': Group1, 'OAS Count': CountQueryB, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupC, 'Group': Group1, 'OAS Count': CountQueryC, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupD, 'Group': Group1, 'OAS Count': CountQueryD, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupE, 'Group': Group1, 'OAS Count': CountQueryE, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupF, 'Group': Group1, 'OAS Count': CountQueryF, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupG, 'Group': Group1, 'OAS Count': CountQueryG, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupH, 'Group': Group1, 'OAS Count': CountQueryH, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupI, 'Group': Group1, 'OAS Count': CountQueryI, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupJ, 'Group': Group1, 'OAS Count': CountQueryJ, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupK, 'Group': Group1, 'OAS Count': CountQueryK, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupL, 'Group': Group1, 'OAS Count': CountQueryL, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupM, 'Group': Group1, 'OAS Count': CountQueryM, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupN, 'Group': Group1, 'OAS Count': CountQueryN, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupO, 'Group': Group1, 'OAS Count': CountQueryO, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupP, 'Group': Group1, 'OAS Count': CountQueryP, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupQ, 'Group': Group1, 'OAS Count': CountQueryQ, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupR, 'Group': Group1, 'OAS Count': CountQueryR, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupS, 'Group': Group1, 'OAS Count': CountQueryS, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupT, 'Group': Group1, 'OAS Count': CountQueryT, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupU, 'Group': Group1, 'OAS Count': CountQueryU, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupV, 'Group': Group1, 'OAS Count': CountQueryV, 'Type': Type2, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupW, 'Group': Group1, 'OAS Count': CountQueryW, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)

    # Store the file name and OAS Count in the DataFrame
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupA, 'Group': Group2, 'OAS Count': CountQueryAA, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupB, 'Group': Group2, 'OAS Count': CountQueryBB, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupC, 'Group': Group2, 'OAS Count': CountQueryCC, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupD, 'Group': Group2, 'OAS Count': CountQueryDD, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupE, 'Group': Group2, 'OAS Count': CountQueryEE, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupF, 'Group': Group2, 'OAS Count': CountQueryFF, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupG, 'Group': Group2, 'OAS Count': CountQueryGG, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupH, 'Group': Group2, 'OAS Count': CountQueryHH, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupI, 'Group': Group2, 'OAS Count': CountQueryII, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupJ, 'Group': Group2, 'OAS Count': CountQueryJJ, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupK, 'Group': Group2, 'OAS Count': CountQueryKK, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupL, 'Group': Group2, 'OAS Count': CountQueryLL, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupM, 'Group': Group2, 'OAS Count': CountQueryMM, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupN, 'Group': Group2, 'OAS Count': CountQueryNN, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupO, 'Group': Group2, 'OAS Count': CountQueryOO, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupP, 'Group': Group2, 'OAS Count': CountQueryPP, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupQ, 'Group': Group2, 'OAS Count': CountQueryQQ, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupR, 'Group': Group2, 'OAS Count': CountQueryRR, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupS, 'Group': Group2, 'OAS Count': CountQuerySS, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupT, 'Group': Group2, 'OAS Count': CountQueryTT, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupU, 'Group': Group2, 'OAS Count': CountQueryUU, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupV, 'Group': Group2, 'OAS Count': CountQueryVV, 'Type': Type2, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupW, 'Group': Group2, 'OAS Count': CountQueryWW, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)

    # Store the file name and OAS Count in the DataFrame
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupA, 'Group': Group3, 'OAS Count': CountQueryAAA, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupB, 'Group': Group3, 'OAS Count': CountQueryBBB, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupC, 'Group': Group3, 'OAS Count': CountQueryCCC, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupD, 'Group': Group3, 'OAS Count': CountQueryDDD, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupE, 'Group': Group3, 'OAS Count': CountQueryEEE, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupF, 'Group': Group3, 'OAS Count': CountQueryFFF, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupG, 'Group': Group3, 'OAS Count': CountQueryGGG, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupH, 'Group': Group3, 'OAS Count': CountQueryHHH, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupI, 'Group': Group3, 'OAS Count': CountQueryIII, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupJ, 'Group': Group3, 'OAS Count': CountQueryJJJ, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupK, 'Group': Group3, 'OAS Count': CountQueryKKK, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupL, 'Group': Group3, 'OAS Count': CountQueryLLL, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupM, 'Group': Group3, 'OAS Count': CountQueryMMM, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupN, 'Group': Group3, 'OAS Count': CountQueryNNN, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupO, 'Group': Group3, 'OAS Count': CountQueryOOO, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupP, 'Group': Group3, 'OAS Count': CountQueryPPP, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupQ, 'Group': Group3, 'OAS Count': CountQueryQQQ, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupR, 'Group': Group3, 'OAS Count': CountQueryRRR, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupS, 'Group': Group3, 'OAS Count': CountQuerySSS, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupT, 'Group': Group3, 'OAS Count': CountQueryTTT, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupU, 'Group': Group3, 'OAS Count': CountQueryUUU, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupV, 'Group': Group3, 'OAS Count': CountQueryVVV, 'Type': Type2, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupW, 'Group': Group3, 'OAS Count': CountQueryWWW, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)

    # Store the file name and OAS Count in the DataFrame
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupA, 'Group': Group4, 'OAS Count': CountQueryAAAA, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupB, 'Group': Group4, 'OAS Count': CountQueryBBBB, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupC, 'Group': Group4, 'OAS Count': CountQueryCCCC, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupD, 'Group': Group4, 'OAS Count': CountQueryDDDD, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupE, 'Group': Group4, 'OAS Count': CountQueryEEEE, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupF, 'Group': Group4, 'OAS Count': CountQueryFFFF, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupG, 'Group': Group4, 'OAS Count': CountQueryGGGG, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupH, 'Group': Group4, 'OAS Count': CountQueryHHHH, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupI, 'Group': Group4, 'OAS Count': CountQueryIIII, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupJ, 'Group': Group4, 'OAS Count': CountQueryJJJJ, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupK, 'Group': Group4, 'OAS Count': CountQueryKKKK, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupL, 'Group': Group4, 'OAS Count': CountQueryLLLL, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupM, 'Group': Group4, 'OAS Count': CountQueryMMMM, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupN, 'Group': Group4, 'OAS Count': CountQueryNNNN, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupO, 'Group': Group4, 'OAS Count': CountQueryOOOO, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupP, 'Group': Group4, 'OAS Count': CountQueryPPPP, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupQ, 'Group': Group4, 'OAS Count': CountQueryQQQQ, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupR, 'Group': Group4, 'OAS Count': CountQueryRRRR, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupS, 'Group': Group4, 'OAS Count': CountQuerySSSS, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupT, 'Group': Group4, 'OAS Count': CountQueryTTTT, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupU, 'Group': Group4, 'OAS Count': CountQueryUUUU, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupV, 'Group': Group4, 'OAS Count': CountQueryVVVV, 'Type': Type2, 'Conversion Status': ConStatsTrue}, ignore_index=True)
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupW, 'Group': Group4, 'OAS Count': CountQueryWWWW, 'Type': Type1, 'Conversion Status': ConStatsTrue}, ignore_index=True)

for file in abs_scanned_excel_files:
    # Read the Excel file
    df = pd.read_excel(file)

    # Get the OAS Count
    record_count = len(df)

    # Store the file name and OAS Count in the DataFrame
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupABS, 'Group': GroupABS, 'OAS Count': record_count, 'Type': Type0, 'Scan Status': ScanStatusTrue}, ignore_index=True)

for file in abs_pending_excel_files:
    # Read the Excel file
    df = pd.read_excel(file)

    # Get the OAS Count
    record_count = len(df)

    # Store the file name and OAS Count in the DataFrame
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupABS,  'Group': GroupABS, 'OAS Count': record_count, 'Type': Type0, 'Conversion Status': ConStatsFalse}, ignore_index=True)

for file in abs_outstanding_excel_files:
    # Read the Excel file
    df = pd.read_excel(file)

    # Get the OAS Count
    record_count = len(df)

    # Store the file name and OAS Count in the DataFrame
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupABS,  'Group': GroupABS, 'OAS Count': record_count, 'Type': Type0, 'Scan Status': ScanStatusFalse}, ignore_index=True)

for file in abs_converted_excel_files:
    # Read the Excel file
    df = pd.read_excel(file)

    # Get the OAS Count
    record_count = len(df)

    # Store the file name and OAS Count in the DataFrame
    file_counts = file_counts.append({'File Name': file, 'Centre': GroupABS,  'Group': GroupABS, 'OAS Count': record_count, 'Type': Type0, 'Conversion Status': ConStatsTrue}, ignore_index=True)

# Write the file counts to a new Excel file
file_counts.to_excel('z_Final.xlsx', index=False)