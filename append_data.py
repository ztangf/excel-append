import os
import pandas as pd
from openpyxl import load_workbook


current_dir = os.path.dirname(__file__)

files = os.listdir(current_dir)

#file_path = 'C:\Users\Rodrigo Dácio\Documents\MyCode\Projects\Test\data_test_excel_mod.xlsx'


wb = load_workbook('Consumo água_01.xlsx')
ws = wb['Consumo_água']


#Lendo planilha
df = pd.read_excel('Consumo água_01.xlsx', sheet_name='Consumo_água')
print(df)

# Data to be appended as rows
data = [{"Consumo": 4537, "Dias": 25}, {"Consumo": 8000, "Dias": 21}]


# Create a pandas DataFrame
df = pd.DataFrame(data)

# Calculate the new column
df['Teste Formula'] = df['Consumo'] + df['Dias']

# Append the DataFrame to the Excel sheet
for row in df.values:
    ws.append(row.tolist())

wb.save('new_Consumo água_01.xlsx')
