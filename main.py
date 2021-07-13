from ast import Str
from typing import List
from numpy.core.getlimits import _register_known_types
import pandas as pd #Pandas for data handling
import matplotlib.pyplot as plt #Matplot for Visuals
import numpy as np
from pandas.core.algorithms import isin
from pandas.core.frame import DataFrame #Numpy for Math
from pandas.io.pytables import dropna_doc
print("Pandas: "+pd.__version__) # 1.1.0

#First verify you can read an Excel file
#FP is file path
fp = pd.ExcelFile('Automation\hiring_muti.xlsx', engine='openpyxl')
print("Utalizing: "+fp.engine)

#Sheet to use identified as 'sh' Prints available sheets in document.
sh = " "
print("Sheets: "+sh.join(fp.sheet_names))

#Use a specific excel sheet in file. Making a dataframe for each tab and cleaning data.
ReportSheet = pd.read_excel(fp, sheet_name='Report2').dropna()
AccountSheet = pd.read_excel(fp, sheet_name='Sheet1').dropna()
#print(ReportSheet)
#print(AccountSheet)

#  #  #  #  #  #  First Task  #  #  #  #  #  #  #  #  
# Generate new sheet with values from first report minus accounts on the second sheet.
# To do this I chose to first find the values im looking for & store them.

#Skipping first Two rows & cleaning Database Names
ReportSheet = ReportSheet.iloc[1: , :]
ReportSheet.columns = ['Trade_Reference','Cpty_Code','Account_Number', 'Counterparty_Fullname','Cpty_Nationality']

#Cleaning Account Sheet Structure
col1 = AccountSheet.columns[0]
col2 = AccountSheet.columns[1]
AccountSheet.columns = ['Account_Number', 'Risk']
new_row = {'Account_Number':col1, 'Risk':col2} 
AccountSheet = AccountSheet.append(new_row, ignore_index=True)

#Saving all accounts as integers and storing them in a list.
Accounts = AccountSheet['Account_Number'].astype(int).tolist()
#Filtering The Data 
Task_One = ReportSheet[~ReportSheet['Account_Number'].astype(int).isin(Accounts)]

print(ReportSheet)

#The Below code saves changes but only to a new file, even if I specify same file new tab 
#writer = pd.ExcelWriter('Automation\output.xlsx', engine='openpyxl')
#Task_One.to_excel(writer, 'Task_One', index = False)
#writer.save()

#  #  #  #  #  #  Second Task  #  #  #  #  #  #  #  #  
blue = [610734, 500364, 610515, 200220]
red = 600411
yellow = [100131, 608818, 602484, 605225, 350967]

TaskTwo = ReportSheet.style.apply(lambda x: ['background:blue' if x == blue else 'background:white' for x in ReportSheet['Account_Number'].astype(int)], axis=0)

writer = pd.ExcelWriter('Automation\output.xlsx', engine='openpyxl')
TaskTwo.to_excel(writer, 'Task_One', index = False)
writer.save()