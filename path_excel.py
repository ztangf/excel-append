import os
import pandas as pd
from openpyxl import load_workbook

current_dir = os.path.dirname(__file__)

files = os.listdir(current_dir)

#file_path = 'C:\Users\Rodrigo Dácio\Documents\MyCode\Projects\Test\data_test_excel_mod.xlsx'

df = pd.read_excel('data_test_excel_mod.xlsx', sheet_name='Sheet1')
#wb = load_workbook('data_test_excel_mod.xlsx')
#ws = wb['Sheet1']

df = pd.read_excel('data_test_excel_mod.xlsx', sheet_name='Sheet1')

#last_row = ws.max_row

data = {"Consumo": [4537] ,"Dias": [25]}
new_df = pd.DataFrame(data)
df = pd.concat([df,new_df],ignore_index=True)

data2 = {"Consumo": [8000] ,"Dias": [21]}
new_df2 = pd.DataFrame(data2)
df = pd.concat([df,new_df2],ignore_index=True)

df['Teste Formula'] = df['Consumo'] + df['Dias']

print(df)

# for r, row in enumerate(df.values.tolist(), 2):
#     for c, value in enumerate(row,start=1):
#         ws.cell(row=r,column=c).value = value

#wb.save('data_test_excel_mod.xlsx')
df.to_excel('data_test_excel_mod.xlsx', index = False)
