import pandas as pd
import os

#parameters
folder_path = "/Users/ganapathisubramaniam/Downloads/OPEB"
file_name = "Hinsdale_excel.xlsx"
assumptions_sheet = "Assumptions"
mortality_sheet = "Mortality"
improvement_sheet = "MP-2021"
actives_sheet = "Active Census"
retirees_sheet = "Retiree Census"
soa_sheet = "SOA_Adj"

# Construct file path
file_path = os.path.join(folder_path, file_name)

# Read the specific range A1:C15 from the Assumptions sheet
assump_raw = pd.read_excel(
    file_path,
    sheet_name=assumptions_sheet,
    header=None,           
    usecols="A:C",         # columns A to C
    nrows=15,              # rows 1 to 15
    engine="openpyxl"
)

#Checks o/p --> Matches w/ input
print(assump_raw)