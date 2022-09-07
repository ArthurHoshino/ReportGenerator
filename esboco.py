# we need to get the articles' titles from a folder
# and separate them into in three different sheets in Excel

import os
import creds
import pandas as pd

directory = creds.user_path
folderList = []
fileList = []
positiveList = []
negativeList = []
neutralList = []

x = os.chdir(directory)
folderList.append(os.listdir(x))

for a in folderList[0]:
    y = os.chdir(directory + a)

    for b in os.listdir(y):
        fileList.append(b)

# print(folderList)
# print(fileList)

# create Data-Frame
data = pd.DataFrame({'Titles': fileList})

# write to Excel
toExcel = pd.ExcelWriter(creds.excel_path, engine='xlsxwriter')

data.to_excel(toExcel, sheet_name='Titles', index=False)

toExcel.save()


# excelTable = pd.read_excel(creds.excel_path, sheet_name=['Titles', 'Positivo', 'Negattivo', 'Neutro'])
excelTable = pd.read_excel(creds.excel_path)
excelTable.reset_index()

for index, row in excelTable.iterrows():
    z = row['Titles'].strip()

    if z.startswith('1'):
        positiveList.append(z.replace('1', '').strip().capitalize())
    elif z.startswith('2'):
        negativeList.append(z.replace('2', '').strip().capitalize())
    else:
        neutralList.append(z.replace('3', '').strip().capitalize())

positiveData = pd.DataFrame({'Positivo': positiveList})
negativeData = pd.DataFrame({'Negativo': negativeList})
neutralData = pd.DataFrame({'Neutro': neutralList})

with pd.ExcelWriter(creds.excel_path2) as writer:
    positiveData.to_excel(writer, sheet_name='Positivo', index=False)
    negativeData.to_excel(writer, sheet_name='Negativo', index=False)
    neutralData.to_excel(writer, sheet_name='Neutro', index=False)