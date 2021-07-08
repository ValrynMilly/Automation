import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
print("Pandas: "+pd.__version__)

#First verify you can read an Excel file
xls = pd.ExcelFile('Automation\dti2020.xlsx', engine='openpyxl')
print("Utalizing: "+xls.engine)

#Identifying sheet names
sh = " "
print("Sheets: "+sh.join(xls.sheet_names))

#To use a specific excel sheet in file.
df = pd.read_excel(xls, sheet_name='Exports')

#Plot the information (In this case its a Bar Chart)
plt.rcdefaults()
ax = plt.subplot()

#Specify specific values to show in dataframe
values = df[['Country of Destination', 'ValueGBP']].dropna()


# Example data
y_pos = np.arange(len(values['Country of Destination']))
Price = 3 + 10 * np.random.rand(len(values['ValueGBP']))


ax.barh(y_pos, Price, align='center')
ax.set_yticks(y_pos)
ax.set_yticklabels(values['Country of Destination'])
ax.invert_yaxis()  # labels read top-to-bottom
ax.set_xlabel('Price (£ - GBP)')
ax.set_title('NON-EU Export Data in the Energy Sector.')

plt.show()