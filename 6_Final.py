import pandas as pd
import os

# Check and delete previous Excel file
try:
    # Pending Conversion
    os.path.exists("./z_Final.xlsx")
    os.remove("./z_Final.xlsx")
    print("A previous version of z_Final detected! Deleting file...")
except:
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
    'PO Converted OAS.xlsx',
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

# Scan Status
ScanTrue = "Scanned"
ScanFalse = "Outstanding"

# Conversion Status
ConvertTrue = "Converted"
ConvertFalse = "Pending Conversion"

# Batches
Batch1 = "P6, OL Courses (Week 36-38)"
Batch2 = "S1-3 Chinese (Week 31-32)"
Batch3 = "P3-5 Courses (Week 34-35)"
Batch4 = "S1-3 Sciences (Week 33-34)"

# Types
DataType1 = "Class"
DataType2 = "Absentee"
DataType3 = "PO Students"

year = '2023'

batch_codes = ['BM', 'BS', 'BBC', 'BP', 'TP', 'CCK', 'EC', 'HG', 'JW', 'LQ',
               'MP', 'NS', 'PN', 'PC', 'PG', 'RM', 'SM', 'SW', 'SME', 'WC',
               'WLC', 'YS', 'BS8', 'WLH', 'OPG', 'ED']

query_batch1_code = ' or '.join([f'`Class: Class Code`.str.contains("{code+year}-P6") or `Class: Class Code`.str.contains("{code+year}-OL")' for code in batch_codes])
query_batch2_code = ' or '.join([f'`Class: Class Code`.str.contains("{code+year}-S1C") or `Class: Class Code`.str.contains("{code+year}-S2C") or `Class: Class Code`.str.contains("{code+year}-S3C")' for code in batch_codes])
query_batch3_code = ' or '.join([f'`Class: Class Code`.str.contains("{code+year}-P3") or `Class: Class Code`.str.contains("{code+year}-P4") or `Class: Class Code`.str.contains("{code+year}-P5")' for code in batch_codes])
query_batch4_code = ' or '.join([f'`Class: Class Code`.str.contains("{code+year}-S1S") or `Class: Class Code`.str.contains("{code+year}-S2S") or `Class: Class Code`.str.contains("{code+year}-S3S")' for code in batch_codes])

# Create a DataFrame to store the counts
file_counts = pd.DataFrame(columns=['File Name', 'Centre', 'Batch', 'Count', 'Type', 'Scan Status', 'Conversion Status'])

# Create a list to store the data for each row
scanned_data_rows = []
outstanding_data_rows = []
converted_data_rows = []
pending_data_rows = []
abs_scanned_data_rows = []
abs_outstanding_data_rows = []
abs_converted_data_rows = []
abs_pending_data_rows = []

# Iterate Excel files - Scanned
for file in scanned_excel_files:
    # Read the Excel file
    df = pd.read_excel(file)

    for code in batch_codes:
        query_batch1_code = f'`Class: Class Code`.str.contains("{code+year}-P6") or `Class: Class Code`.str.contains("{code+year}-OL")'
        query_batch1_code_df = df.query(query_batch1_code, engine='python')
        count_query_batch1_code = len(query_batch1_code_df)

        if code == "ED":
            scanned_data_rows.append({'File Name': file, 'Centre': code, 'Batch': Batch1, 'Count': count_query_batch1_code, 'Type': DataType3, 'Scan Status': ScanTrue})
        else:
            scanned_data_rows.append({'File Name': file, 'Centre': code, 'Batch': Batch1, 'Count': count_query_batch1_code, 'Type': DataType1, 'Scan Status': ScanTrue})

        query_batch2_code = f'`Class: Class Code`.str.contains("{code+year}-S1C") or `Class: Class Code`.str.contains("{code+year}-S2C") or `Class: Class Code`.str.contains("{code+year}-S3C")'
        query_batch2_code_df = df.query(query_batch2_code, engine='python')
        count_query_batch2_code = len(query_batch2_code_df)

        if code == "ED":
            scanned_data_rows.append({'File Name': file, 'Centre': code, 'Batch': Batch2, 'Count': count_query_batch2_code, 'Type': DataType3, 'Scan Status': ScanTrue})
        else:
            scanned_data_rows.append({'File Name': file, 'Centre': code, 'Batch': Batch2, 'Count': count_query_batch2_code, 'Type': DataType1, 'Scan Status': ScanTrue})

        query_batch3_code = f'`Class: Class Code`.str.contains("{code+year}-P3") or `Class: Class Code`.str.contains("{code+year}-P4") or `Class: Class Code`.str.contains("{code+year}-P5")'
        query_batch3_code_df = df.query(query_batch3_code, engine='python')
        count_query_batch3_code = len(query_batch3_code_df)

        if code == "ED":
            scanned_data_rows.append({'File Name': file, 'Centre': code, 'Batch': Batch3, 'Count': count_query_batch3_code, 'Type': DataType3, 'Scan Status': ScanTrue})
        else:
            scanned_data_rows.append({'File Name': file, 'Centre': code, 'Batch': Batch3, 'Count': count_query_batch3_code, 'Type': DataType1, 'Scan Status': ScanTrue})

        query_batch4_code = f'`Class: Class Code`.str.contains("{code+year}-S1S") or `Class: Class Code`.str.contains("{code+year}-S2S") or `Class: Class Code`.str.contains("{code+year}-S3S")'
        query_batch4_code_df = df.query(query_batch4_code, engine='python')
        count_query_batch4_code = len(query_batch4_code_df)

        if code == "ED":
            scanned_data_rows.append({'File Name': file, 'Centre': code, 'Batch': Batch4, 'Count': count_query_batch4_code, 'Type': DataType3, 'Scan Status': ScanTrue})
        else:
            scanned_data_rows.append({'File Name': file, 'Centre': code, 'Batch': Batch4, 'Count': count_query_batch4_code, 'Type': DataType1, 'Scan Status': ScanTrue})

# Iterate Excel files - Outstanding
for file in outstanding_excel_files:
    # Read the Excel file
    df = pd.read_excel(file)

    for code in batch_codes:
        query_batch1_code = f'`Class: Class Code`.str.contains("{code+year}-P6") or `Class: Class Code`.str.contains("{code+year}-OL")'
        query_batch1_code_df = df.query(query_batch1_code, engine='python')
        count_query_batch1_code = len(query_batch1_code_df)

        if code == "ED":
            outstanding_data_rows.append({'File Name': file, 'Centre': code, 'Batch': Batch1, 'Count': count_query_batch1_code, 'Type': DataType3, 'Scan Status': ScanFalse})
        else:
            outstanding_data_rows.append({'File Name': file, 'Centre': code, 'Batch': Batch1, 'Count': count_query_batch1_code, 'Type': DataType1, 'Scan Status': ScanFalse})

        query_batch2_code = f'`Class: Class Code`.str.contains("{code+year}-S1C") or `Class: Class Code`.str.contains("{code+year}-S2C") or `Class: Class Code`.str.contains("{code+year}-S3C")'
        query_batch2_code_df = df.query(query_batch2_code, engine='python')
        count_query_batch2_code = len(query_batch2_code_df)

        if code == "ED":
            outstanding_data_rows.append({'File Name': file, 'Centre': code, 'Batch': Batch2, 'Count': count_query_batch2_code, 'Type': DataType3, 'Scan Status': ScanFalse})
        else:
            outstanding_data_rows.append({'File Name': file, 'Centre': code, 'Batch': Batch2, 'Count': count_query_batch2_code, 'Type': DataType1, 'Scan Status': ScanFalse})

        query_batch3_code = f'`Class: Class Code`.str.contains("{code+year}-P3") or `Class: Class Code`.str.contains("{code+year}-P4") or `Class: Class Code`.str.contains("{code+year}-P5")'
        query_batch3_code_df = df.query(query_batch3_code, engine='python')
        count_query_batch3_code = len(query_batch3_code_df)

        if code == "ED":
            outstanding_data_rows.append({'File Name': file, 'Centre': code, 'Batch': Batch3, 'Count': count_query_batch3_code, 'Type': DataType3, 'Scan Status': ScanFalse})
        else:
            outstanding_data_rows.append({'File Name': file, 'Centre': code, 'Batch': Batch3, 'Count': count_query_batch3_code, 'Type': DataType1, 'Scan Status': ScanFalse})

        query_batch4_code = f'`Class: Class Code`.str.contains("{code+year}-S1S") or `Class: Class Code`.str.contains("{code+year}-S2S") or `Class: Class Code`.str.contains("{code+year}-S3S")'
        query_batch4_code_df = df.query(query_batch4_code, engine='python')
        count_query_batch4_code = len(query_batch4_code_df)

        if code == "ED":
            outstanding_data_rows.append({'File Name': file, 'Centre': code, 'Batch': Batch4, 'Count': count_query_batch4_code, 'Type': DataType3, 'Scan Status': ScanFalse})
        else:
            outstanding_data_rows.append({'File Name': file, 'Centre': code, 'Batch': Batch4, 'Count': count_query_batch4_code, 'Type': DataType1, 'Scan Status': ScanFalse})

# Iterate Excel files - Converted
for file in converted_excel_files:
    # Read the Excel file
    df = pd.read_excel(file)

    for code in batch_codes:
        query_batch1_code = f'`Class: Class Code`.str.contains("{code+year}-P6") or `Class: Class Code`.str.contains("{code+year}-OL")'
        query_batch1_code_df = df.query(query_batch1_code, engine='python')
        count_query_batch1_code = len(query_batch1_code_df)

        if code == "ED":
            converted_data_rows.append({'File Name': file, 'Centre': code, 'Batch': Batch1, 'Count': count_query_batch1_code, 'Type': DataType3, 'Conversion Status': ConvertTrue})
        else:
            converted_data_rows.append({'File Name': file, 'Centre': code, 'Batch': Batch1, 'Count': count_query_batch1_code, 'Type': DataType1, 'Conversion Status': ConvertTrue})

        query_batch2_code = f'`Class: Class Code`.str.contains("{code+year}-S1C") or `Class: Class Code`.str.contains("{code+year}-S2C") or `Class: Class Code`.str.contains("{code+year}-S3C")'
        query_batch2_code_df = df.query(query_batch2_code, engine='python')
        count_query_batch2_code = len(query_batch2_code_df)

        if code == "ED":
            converted_data_rows.append({'File Name': file, 'Centre': code, 'Batch': Batch2, 'Count': count_query_batch2_code, 'Type': DataType3, 'Conversion Status': ConvertTrue})
        else:
            converted_data_rows.append({'File Name': file, 'Centre': code, 'Batch': Batch2, 'Count': count_query_batch2_code, 'Type': DataType1, 'Conversion Status': ConvertTrue})

        query_batch3_code = f'`Class: Class Code`.str.contains("{code+year}-P3") or `Class: Class Code`.str.contains("{code+year}-P4") or `Class: Class Code`.str.contains("{code+year}-P5")'
        query_batch3_code_df = df.query(query_batch3_code, engine='python')
        count_query_batch3_code = len(query_batch3_code_df)

        if code == "ED":
            converted_data_rows.append({'File Name': file, 'Centre': code, 'Batch': Batch3, 'Count': count_query_batch3_code, 'Type': DataType3, 'Conversion Status': ConvertTrue})
        else:
            converted_data_rows.append({'File Name': file, 'Centre': code, 'Batch': Batch3, 'Count': count_query_batch3_code, 'Type': DataType1, 'Conversion Status': ConvertTrue})

        query_batch4_code = f'`Class: Class Code`.str.contains("{code+year}-S1S") or `Class: Class Code`.str.contains("{code+year}-S2S") or `Class: Class Code`.str.contains("{code+year}-S3S")'
        query_batch4_code_df = df.query(query_batch4_code, engine='python')
        count_query_batch4_code = len(query_batch4_code_df)

        if code == "ED":
            converted_data_rows.append({'File Name': file, 'Centre': code, 'Batch': Batch4, 'Count': count_query_batch4_code, 'Type': DataType3, 'Conversion Status': ConvertTrue})
        else:
            converted_data_rows.append({'File Name': file, 'Centre': code, 'Batch': Batch4, 'Count': count_query_batch4_code, 'Type': DataType1, 'Conversion Status': ConvertTrue})

# Iterate Excel files - Pending Conversion
for file in pending_excel_files:
    # Read the Excel file
    df = pd.read_excel(file)

    for code in batch_codes:
        query_batch1_code = f'`Class: Class Code`.str.contains("{code+year}-P6") or `Class: Class Code`.str.contains("{code+year}-OL")'
        query_batch1_code_df = df.query(query_batch1_code, engine='python')
        count_query_batch1_code = len(query_batch1_code_df)

        if code == "ED":
            pending_data_rows.append({'File Name': file, 'Centre': code, 'Batch': Batch1, 'Count': count_query_batch1_code, 'Type': DataType3, 'Conversion Status': ConvertFalse})
        else:
            pending_data_rows.append({'File Name': file, 'Centre': code, 'Batch': Batch1, 'Count': count_query_batch1_code, 'Type': DataType1, 'Conversion Status': ConvertFalse})

        query_batch2_code = f'`Class: Class Code`.str.contains("{code+year}-S1C") or `Class: Class Code`.str.contains("{code+year}-S2C") or `Class: Class Code`.str.contains("{code+year}-S3C")'
        query_batch2_code_df = df.query(query_batch2_code, engine='python')
        count_query_batch2_code = len(query_batch2_code_df)

        if code == "ED":
            pending_data_rows.append({'File Name': file, 'Centre': code, 'Batch': Batch2, 'Count': count_query_batch2_code, 'Type': DataType3, 'Conversion Status': ConvertFalse})
        else:
            pending_data_rows.append({'File Name': file, 'Centre': code, 'Batch': Batch2, 'Count': count_query_batch2_code, 'Type': DataType1, 'Conversion Status': ConvertFalse})

        query_batch3_code = f'`Class: Class Code`.str.contains("{code+year}-P3") or `Class: Class Code`.str.contains("{code+year}-P4") or `Class: Class Code`.str.contains("{code+year}-P5")'
        query_batch3_code_df = df.query(query_batch3_code, engine='python')
        count_query_batch3_code = len(query_batch3_code_df)

        if code == "ED":
            pending_data_rows.append({'File Name': file, 'Centre': code, 'Batch': Batch3, 'Count': count_query_batch3_code, 'Type': DataType3, 'Conversion Status': ConvertFalse})
        else:
            pending_data_rows.append({'File Name': file, 'Centre': code, 'Batch': Batch3, 'Count': count_query_batch3_code, 'Type': DataType1, 'Conversion Status': ConvertFalse})

        query_batch4_code = f'`Class: Class Code`.str.contains("{code+year}-S1S") or `Class: Class Code`.str.contains("{code+year}-S2S") or `Class: Class Code`.str.contains("{code+year}-S3S")'
        query_batch4_code_df = df.query(query_batch4_code, engine='python')
        count_query_batch4_code = len(query_batch4_code_df)

        if code == "ED":
            pending_data_rows.append({'File Name': file, 'Centre': code, 'Batch': Batch4, 'Count': count_query_batch4_code, 'Type': DataType3, 'Conversion Status': ConvertFalse})
        else:
            pending_data_rows.append({'File Name': file, 'Centre': code, 'Batch': Batch4, 'Count': count_query_batch4_code, 'Type': DataType1, 'Conversion Status': ConvertFalse})

# Iterate Excel files - Scanned Absentees
for file in abs_scanned_excel_files:
    # Read the Excel file
    df = pd.read_excel(file)

    for code in batch_codes:
        query_batch1_code = f'`Class: Class Code`.str.contains("{code+year}-P6") or `Class: Class Code`.str.contains("{code+year}-OL")'
        query_batch1_code_df = df.query(query_batch1_code, engine='python')
        count_query_batch1_code = len(query_batch1_code_df)
        
        if code != "ED":
            abs_scanned_data_rows.append({'File Name': file, 'Centre': code, 'Batch': Batch1, 'Count': count_query_batch1_code, 'Type': DataType2, 'Scan Status': ScanTrue})

        query_batch2_code = f'`Class: Class Code`.str.contains("{code+year}-S1C") or `Class: Class Code`.str.contains("{code+year}-S2C") or `Class: Class Code`.str.contains("{code+year}-S3C")'
        query_batch2_code_df = df.query(query_batch2_code, engine='python')
        count_query_batch2_code = len(query_batch2_code_df)

        if code != "ED":
            abs_scanned_data_rows.append({'File Name': file, 'Centre': code, 'Batch': Batch2, 'Count': count_query_batch2_code, 'Type': DataType2, 'Scan Status': ScanTrue})

        query_batch3_code = f'`Class: Class Code`.str.contains("{code+year}-P3") or `Class: Class Code`.str.contains("{code+year}-P4") or `Class: Class Code`.str.contains("{code+year}-P5")'
        query_batch3_code_df = df.query(query_batch3_code, engine='python')
        count_query_batch3_code = len(query_batch3_code_df)

        if code != "ED":
            abs_scanned_data_rows.append({'File Name': file, 'Centre': code, 'Batch': Batch3, 'Count': count_query_batch3_code, 'Type': DataType2, 'Scan Status': ScanTrue})

        query_batch4_code = f'`Class: Class Code`.str.contains("{code+year}-S1S") or `Class: Class Code`.str.contains("{code+year}-S2S") or `Class: Class Code`.str.contains("{code+year}-S3S")'
        query_batch4_code_df = df.query(query_batch4_code, engine='python')
        count_query_batch4_code = len(query_batch4_code_df)

        if code != "ED":
            abs_scanned_data_rows.append({'File Name': file, 'Centre': code, 'Batch': Batch4, 'Count': count_query_batch4_code, 'Type': DataType2, 'Scan Status': ScanTrue})
        
# Iterate Excel files - Outstanding Absentees
for file in abs_outstanding_excel_files:
    # Read the Excel file
    df = pd.read_excel(file)

    for code in batch_codes:
        query_batch1_code = f'`Class: Class Code`.str.contains("{code+year}-P6") or `Class: Class Code`.str.contains("{code+year}-OL")'
        query_batch1_code_df = df.query(query_batch1_code, engine='python')
        count_query_batch1_code = len(query_batch1_code_df)
        
        if code != "ED":
            abs_outstanding_data_rows.append({'File Name': file, 'Centre': code, 'Batch': Batch1, 'Count': count_query_batch1_code, 'Type': DataType2, 'Scan Status': ScanFalse})

        query_batch2_code = f'`Class: Class Code`.str.contains("{code+year}-S1C") or `Class: Class Code`.str.contains("{code+year}-S2C") or `Class: Class Code`.str.contains("{code+year}-S3C")'
        query_batch2_code_df = df.query(query_batch2_code, engine='python')
        count_query_batch2_code = len(query_batch2_code_df)
        
        if code != "ED":
            abs_outstanding_data_rows.append({'File Name': file, 'Centre': code, 'Batch': Batch2, 'Count': count_query_batch2_code, 'Type': DataType2, 'Scan Status': ScanFalse})

        query_batch3_code = f'`Class: Class Code`.str.contains("{code+year}-P3") or `Class: Class Code`.str.contains("{code+year}-P4") or `Class: Class Code`.str.contains("{code+year}-P5")'
        query_batch3_code_df = df.query(query_batch3_code, engine='python')
        count_query_batch3_code = len(query_batch3_code_df)

        if code != "ED":
            abs_outstanding_data_rows.append({'File Name': file, 'Centre': code, 'Batch': Batch3, 'Count': count_query_batch3_code, 'Type': DataType2, 'Scan Status': ScanFalse})

        query_batch4_code = f'`Class: Class Code`.str.contains("{code+year}-S1S") or `Class: Class Code`.str.contains("{code+year}-S2S") or `Class: Class Code`.str.contains("{code+year}-S3S")'
        query_batch4_code_df = df.query(query_batch4_code, engine='python')
        count_query_batch4_code = len(query_batch4_code_df)

        if code != "ED":
            abs_outstanding_data_rows.append({'File Name': file, 'Centre': code, 'Batch': Batch4, 'Count': count_query_batch4_code, 'Type': DataType2, 'Scan Status': ScanFalse})

# Iterate Excel files - Converted Absentees
for file in abs_converted_excel_files:
    # Read the Excel file
    df = pd.read_excel(file)

    for code in batch_codes:
        query_batch1_code = f'`Class: Class Code`.str.contains("{code+year}-P6") or `Class: Class Code`.str.contains("{code+year}-OL")'
        query_batch1_code_df = df.query(query_batch1_code, engine='python')
        count_query_batch1_code = len(query_batch1_code_df)
        
        if code != "ED":
            abs_converted_data_rows.append({'File Name': file, 'Centre': code, 'Batch': Batch1, 'Count': count_query_batch1_code, 'Type': DataType2, 'Conversion Status': ConvertTrue})

        query_batch2_code = f'`Class: Class Code`.str.contains("{code+year}-S1C") or `Class: Class Code`.str.contains("{code+year}-S2C") or `Class: Class Code`.str.contains("{code+year}-S3C")'
        query_batch2_code_df = df.query(query_batch2_code, engine='python')
        count_query_batch2_code = len(query_batch2_code_df)
        
        if code != "ED":
            abs_converted_data_rows.append({'File Name': file, 'Centre': code, 'Batch': Batch2, 'Count': count_query_batch2_code, 'Type': DataType2, 'Conversion Status': ConvertTrue})

        query_batch3_code = f'`Class: Class Code`.str.contains("{code+year}-P3") or `Class: Class Code`.str.contains("{code+year}-P4") or `Class: Class Code`.str.contains("{code+year}-P5")'
        query_batch3_code_df = df.query(query_batch3_code, engine='python')
        count_query_batch3_code = len(query_batch3_code_df)

        if code != "ED":
            abs_converted_data_rows.append({'File Name': file, 'Centre': code, 'Batch': Batch3, 'Count': count_query_batch3_code, 'Type': DataType2, 'Conversion Status': ConvertTrue})

        query_batch4_code = f'`Class: Class Code`.str.contains("{code+year}-S1S") or `Class: Class Code`.str.contains("{code+year}-S2S") or `Class: Class Code`.str.contains("{code+year}-S3S")'
        query_batch4_code_df = df.query(query_batch4_code, engine='python')
        count_query_batch4_code = len(query_batch4_code_df)

        if code != "ED":
            abs_converted_data_rows.append({'File Name': file, 'Centre': code, 'Batch': Batch4, 'Count': count_query_batch4_code, 'Type': DataType2, 'Conversion Status': ConvertTrue})

# Iterate Excel files - Pending Conversion Absentees
for file in abs_converted_excel_files:
    # Read the Excel file
    df = pd.read_excel(file)

    for code in batch_codes:
        query_batch1_code = f'`Class: Class Code`.str.contains("{code+year}-P6") or `Class: Class Code`.str.contains("{code+year}-OL")'
        query_batch1_code_df = df.query(query_batch1_code, engine='python')
        count_query_batch1_code = len(query_batch1_code_df)
        
        if code != "ED":
            abs_pending_data_rows.append({'File Name': file, 'Centre': code, 'Batch': Batch1, 'Count': count_query_batch1_code, 'Type': DataType2, 'Conversion Status': ConvertFalse})

        query_batch2_code = f'`Class: Class Code`.str.contains("{code+year}-S1C") or `Class: Class Code`.str.contains("{code+year}-S2C") or `Class: Class Code`.str.contains("{code+year}-S3C")'
        query_batch2_code_df = df.query(query_batch2_code, engine='python')
        count_query_batch2_code = len(query_batch2_code_df)
        
        if code != "ED":
            abs_pending_data_rows.append({'File Name': file, 'Centre': code, 'Batch': Batch2, 'Count': count_query_batch2_code, 'Type': DataType2, 'Conversion Status': ConvertFalse})

        query_batch3_code = f'`Class: Class Code`.str.contains("{code+year}-P3") or `Class: Class Code`.str.contains("{code+year}-P4") or `Class: Class Code`.str.contains("{code+year}-P5")'
        query_batch3_code_df = df.query(query_batch3_code, engine='python')
        count_query_batch3_code = len(query_batch3_code_df)

        if code != "ED":
            abs_pending_data_rows.append({'File Name': file, 'Centre': code, 'Batch': Batch3, 'Count': count_query_batch3_code, 'Type': DataType2, 'Conversion Status': ConvertFalse})

        query_batch4_code = f'`Class: Class Code`.str.contains("{code+year}-S1S") or `Class: Class Code`.str.contains("{code+year}-S2S") or `Class: Class Code`.str.contains("{code+year}-S3S")'
        query_batch4_code_df = df.query(query_batch4_code, engine='python')
        count_query_batch4_code = len(query_batch4_code_df)

        if code != "ED":
            abs_pending_data_rows.append({'File Name': file, 'Centre': code, 'Batch': Batch4, 'Count': count_query_batch4_code, 'Type': DataType2, 'Conversion Status': ConvertFalse})

# Concatenate the data rows into the file_counts DataFrame
file_counts = pd.concat([file_counts, pd.DataFrame(scanned_data_rows)], ignore_index=True)
file_counts = pd.concat([file_counts, pd.DataFrame(outstanding_data_rows)], ignore_index=True)
file_counts = pd.concat([file_counts, pd.DataFrame(converted_data_rows)], ignore_index=True)
file_counts = pd.concat([file_counts, pd.DataFrame(pending_data_rows)], ignore_index=True)
file_counts = pd.concat([file_counts, pd.DataFrame(abs_scanned_data_rows)], ignore_index=True)
file_counts = pd.concat([file_counts, pd.DataFrame(abs_outstanding_data_rows)], ignore_index=True)
file_counts = pd.concat([file_counts, pd.DataFrame(abs_converted_data_rows)], ignore_index=True)
file_counts = pd.concat([file_counts, pd.DataFrame(abs_pending_data_rows)], ignore_index=True)

# Write the file_counts DataFrame to a new Excel file
file_counts.to_excel('z_Final.xlsx', sheet_name='Scanned', index=False)