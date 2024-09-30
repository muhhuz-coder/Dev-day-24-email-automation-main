import pandas as pd
import os
import sys

# reads data from an Excel file containing member information (form responses),
# sorts the members into individual sheets based on their on-day and off-day teams,
# removes unwanted columns from each sheet, and saves the modified data to a new Excel file.

inputFilePath="DEVDAY'24 MEMBERS.xlsx"
tempFilePath= "DEVDAY'24 MEMBERS(grouped) temp.xlsx"
outputFilePath = "DEVDAY'24 MEMBERS(grouped).xlsx"

try:
    inputXls = pd.read_excel(inputFilePath)
except:
    print(f"Cannot read input excel file: {inputFilePath}")
    sys.exit()

OnDayGroups = inputXls.groupby('SELECT ON-DAY TEAM')          # Grouping members by their teams
OffDayGroups = inputXls.groupby('SELECT OFF-DAY TEAM')         
 
try:
    if os.path.exists(tempFilePath):                    # Removing existing temporary file if present
        os.remove(tempFilePath)
    
    with pd.ExcelWriter(tempFilePath) as writer:
        for group_name, group_df in OnDayGroups:
            group_df.to_excel(writer, sheet_name=group_name, index=False)       # Writing data to a temp excel file
        for group_name, group_df in OffDayGroups:
            group_df.to_excel(writer, sheet_name=group_name, index=False)

except PermissionError:
    print(f"Permission denied for {tempFilePath}! File may be in use or already open")
    sys.exit()
except Exception as e:
    print(f"An error occurred: {e}")
    sys.exit()

# Removing unwanted columns from all sheets in the temporary Excel file
tempXls = pd.ExcelFile(tempFilePath)
columns_to_remove = ['Timestamp', 'SELECT ON-DAY TEAM', 'SELECT OFF-DAY TEAM', 'Past EXPERIENCE']  # Columns to be removed

try:
    if os.path.exists(outputFilePath):
        os.remove(outputFilePath)                   # Removing existing output file if present
            
    with pd.ExcelWriter(outputFilePath) as writer:
        for sheet_name in tempXls.sheet_names:
            df = pd.read_excel(tempXls, sheet_name=sheet_name)
            df = df.drop(columns=columns_to_remove, errors='ignore')
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    print('All members data successfully written to output file')
except PermissionError:
    print(f"Permission denied for {outputFilePath}! File may be in use or already open")
    sys.exit()
except Exception as e:
    print(f"An error occurred: {e}")
    sys.exit()

tempXls.close()

# Removing temp file created
if os.path.exists(tempFilePath):    
    os.remove(tempFilePath)
