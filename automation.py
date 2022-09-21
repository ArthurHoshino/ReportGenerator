# we need to get the articles' titles from a folder
# and separate them into in three different sheets in Excel

import os
import creds
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import xlsxwriter

import cv2
from time import strftime
from time import gmtime

# variables
directory = creds.user_path
folderList = []
fileList = []
fileCountList = []
fileBroadcasterList = []
fileCategoriesList = []
positiveList = []
positiveListBrodcaster = []
positiveListDuration = []
negativeList = []
negativeListBrodcaster = []
negativeListDuration = []
neutralList = []
neutralListBrodcaster = []
neutralListDuration = []
count = 0

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
for c in os.listdir(x):
    folderList.append(c)

for a in folderList:
    y = os.chdir(directory + a)
    count = 0

    for b in os.listdir(y):
        fileList.append(b)

        data = cv2.VideoCapture(directory + a + '\\' + b)
        frames = data.get(cv2.CAP_PROP_FRAME_COUNT)
        fps = data.get(cv2.CAP_PROP_FPS)
        seconds = round(frames / fps)

        fileCountList.append(strftime('%H:%M:%S', gmtime(seconds)))

        if b.startswith('1'):
            variable = b.replace('1', '', 1).strip()
            addBroadcaster(fileBroadcasterList, variable)
            fileCategoriesList.append('Positivo')

        elif b.startswith('2'):
            variable = b.replace('2', '', 1).strip()
            addBroadcaster(fileBroadcasterList, variable)
            fileCategoriesList.append('Negativo')

        else:
            variable = b.replace('3', '', 1).strip()
            addBroadcaster(fileBroadcasterList, variable)
            fileCategoriesList.append('Neutro')


# create Data-Frame
data = pd.DataFrame({'Titles': fileList, 'Duration': fileCountList, 'Broadcaster': fileBroadcasterList})

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
        positiveList.append(variable.replace('.mp4', ''))
        # addBroadcaster(positiveListBrodcaster, variable)

        positiveListDuration.append(fileCountList[count])
        positiveListBrodcaster.append(fileBroadcasterList[count])
        count += 1

    elif z.startswith('2'):
        variable = z.replace('2', '', 1).strip()
        negativeList.append(variable.replace('.mp4', ''))
        # addBroadcaster(negativeListBrodcaster, variable)

        negativeListDuration.append(fileCountList[count])
        negativeListBrodcaster.append(fileBroadcasterList[count])
        count += 1

    else:
        variable = z.replace('3', '', 1).strip()
        neutralList.append(variable.replace('.mp4', ''))
        # addBroadcaster(neutralListBrodcaster, variable)

        neutralListDuration.append(fileCountList[count])
        neutralListBrodcaster.append(fileBroadcasterList[count])
        count += 1

colors = sns.color_palette('pastel')[0:5]
labels = ['Positivo', 'Negativo', 'Neutro']

positivo = fileCategoriesList.count('Positivo')
negativo = fileCategoriesList.count('Negativo')
neutro = fileCategoriesList.count('Neutro')

tudo = [positivo, negativo, neutro]

plt.pie(tudo, labels=labels, colors=colors, autopct='%.0f%%')
plt.savefig('teste.png')


positiveData = pd.DataFrame({'Positivo': positiveList, 'Duração': positiveListDuration, 'Emissora': positiveListBrodcaster})
negativeData = pd.DataFrame({'Negativo': negativeList, 'Duração': negativeListDuration, 'Emissora': negativeListBrodcaster})
neutralData = pd.DataFrame({'Neutro': neutralList, 'Duração': neutralListDuration, 'Emissora': neutralListBrodcaster})
geral = pd.DataFrame({'Títulos': fileList, 'Duração': fileCountList, 'Emissora': fileBroadcasterList, 'Categoria': fileCategoriesList})
resumo = pd.DataFrame()

with pd.ExcelWriter(creds.excel_path2) as writer:
    resumo.to_excel(writer, sheet_name='Resumo', index=False)
    positiveData.to_excel(writer, sheet_name='Positivo', index=False)
    negativeData.to_excel(writer, sheet_name='Negativo', index=False)
    neutralData.to_excel(writer, sheet_name='Neutro', index=False)
    geral.to_excel(writer, sheet_name='Geral', index=False)

    worksheetResumo = writer.sheets['Resumo']
    worksheetResumo.insert_image('A1', 'teste.png')

    worksheet = writer.sheets['Geral']
    worksheet.set_column('A:A', 40)
    worksheet.set_column('B:D', 15)

