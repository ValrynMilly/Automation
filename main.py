import pandas as pd
import matplotlib.pyplot as plt
print("Pandas: "+pd.__version__)

#First verify you can read an Excel file
xls = pd.ExcelFile('Automation\dti2020.xlsx', engine='openpyxl')
print("Utalizing: "+xls.engine)

#Identifying sheet names
sh = " "
print("Sheets: "+sh.join(xls.sheet_names))

#To use one excel sheet