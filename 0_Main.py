import subprocess
import logging

# Create log file
logging.basicConfig(filename="log.txt", level=logging.DEBUG, format="%(asctime)s %(message)s", filemode="w")

print("========== Program is running, please wait ==========")

# Run other Python files
try:
    subprocess.run(["python", "1_Franchisor_OAS.py"])
    print("========== 1_Franchisor_OAS.py ran successfully ==========")
    subprocess.run(["python", "2_Franchisee_OAS.py"])
    print("========== 2_Franchisee_OAS.py ran successfully ==========")
    subprocess.run(["python", "3_Franchisor_Absentees.py"])
    print("========== 3_Franchisor_Absentees.py ran successfully ==========")
    subprocess.run(["python", "4_Franchisee_Absentees.py"])
    print("========== 4_Franchisee_Absentees.py ran successfully ==========")
    subprocess.run(["python", "5_Count_By_Group.py"])
    print("========== 5_Count_By_Group.py ran successfully ==========")
except:
    print("Error running the program!")

print("========== All Python files ran successfully! ==========")