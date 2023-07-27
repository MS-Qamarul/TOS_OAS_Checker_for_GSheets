import subprocess

print("========== INFO: Program is running, please wait... ==========")

# Run other Python files
try:
    subprocess.run(["python", "1_Franchisor_OAS.py"])
    subprocess.run(["python", "2_Franchisee_OAS.py"])
    subprocess.run(["python", "3_Franchisor_Absentees.py"])
    subprocess.run(["python", "4_Franchisee_Absentees.py"])
    subprocess.run(["python", "5_PO_OAS.py"])
    subprocess.run(["python", "6_Final.py"])
except:
    print("========== ERROR: Something went wrong! ==========")

print("========== INFO: The program ran successfully! ==========")