import os
from pathlib import Path
from UliPlot.XLSX import auto_adjust_xlsx_column_width
import pandas as pd


data = pd.read_excel('Output.xlsx')
for index,row in data.iterrows():
    if row['Are batteries included?']=='':
        continue
    row['Batteries Included?'] = row['Are batteries included?']
    break
data = data.drop(columns=['Are batteries included?','Batteries Included'])

# Exporting the dataframe in excel file with auto adjusting excel columns
with pd.ExcelWriter("Output2.xlsx") as writer:
    data.to_excel(writer, sheet_name="MySheet")
    auto_adjust_xlsx_column_width(data, writer, sheet_name="MySheet", margin=0)

# Openeing the excel file
absolutePath = Path('Output2.xlsx').resolve()
os.system(f'start Output2.xlsx "{absolutePath}"')