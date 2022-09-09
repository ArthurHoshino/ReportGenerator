# we need to get the articles' titles from a folder
# and separate them into in three different sheets in Excel

import os
import creds
import pandas as pd

directory = creds.user_path
folderList = []
fileList = []
positiveList = []
positiveListBrodcaster = []
negativeList = []
negativeListBrodcaster = []
neutralList = []
neutralListBrodcaster = []

def addBroadcaster(list: list, broadcaster: str):
    if broadcaster.startswith('SBT'):
        list.append('SBT')
    elif broadcaster.startswith('ALERTA SINOP'):
        list.append('NOVATV')
    elif broadcaster.startswith('BALANÇA SINOP'):
        list.append('TV CIDADE VERDE')
    elif broadcaster.startswith('BALANÇO GERAL'):
        list.append('RECORD')
    elif broadcaster.startswith('MT1'):
        list.append('TVCA')
    elif broadcaster.startswith('BOM DIA NORTÃO'):
        list.append('TVCA')
    elif broadcaster.startswith('CIDADE ALERTA'):
        list.append('RECORD')
    elif broadcaster.startswith('DE FRENTE COM A VERDADE'):
        list.append('NOVATV')
    else:
        print('No broadscaster found')


x = os.chdir(directory)
folderList.append(os.listdir(x))

for a in folderList[0]:
    y = os.chdir(directory + a)

    for b in os.listdir(y):
        fileList.append(b)

# create Data-Frame
data = pd.DataFrame({'Titles': fileList})

# write to Excel
toExcel = pd.ExcelWriter(creds.excel_path, engine='xlsxwriter')

data.to_excel(toExcel, sheet_name='Titles', index=False)

toExcel.save()

excelTable = pd.read_excel(creds.excel_path)
excelTable.reset_index()

for index, row in excelTable.iterrows():
    z = row['Titles'].strip()

    if z.startswith('1'):
        variable = z.replace('1', '', 1).strip()
        positiveList.append(variable.replace('.txt', ''))
        addBroadcaster(positiveListBrodcaster, variable)
    elif z.startswith('2'):
        variable = z.replace('2', '', 1).strip()
        negativeList.append(variable.replace('.txt', ''))
        addBroadcaster(negativeListBrodcaster, variable)
    else:
        variable = z.replace('3', '', 1).strip()
        neutralList.append(variable.replace('.txt', ''))
        addBroadcaster(neutralListBrodcaster, variable)

positiveData = pd.DataFrame({'Positivo': positiveList, 'Emissora': positiveListBrodcaster})
negativeData = pd.DataFrame({'Negativo': negativeList, 'Emissora': negativeListBrodcaster})
neutralData = pd.DataFrame({'Neutro': neutralList, 'Emissora': neutralListBrodcaster})

with pd.ExcelWriter(creds.excel_path2) as writer:
    positiveData.to_excel(writer, sheet_name='Positivo', index=False)
    negativeData.to_excel(writer, sheet_name='Negativo', index=False)
    neutralData.to_excel(writer, sheet_name='Neutro', index=False)