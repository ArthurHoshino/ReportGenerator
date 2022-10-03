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
allInfoNameList = ['Matérias positivas', 'Matérias negativas', 'Matérias neutras', 'Tempo positivo', 'Tempo negativo', 'Tempo neutro', 'Tempo total']
allInfoCountList = []
count = 0
positiveTime = 0
negativeTime = 0
neutralTime = 0
allTime = 0

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
        if (b.endswith('.mp4')):
            fileList.append(b)
            if b.startswith('1'):
                variable = b.replace('1', '', 1).strip()
                addBroadcaster(fileBroadcasterList, variable)
                fileCategoriesList.append('Positivo')

                data = cv2.VideoCapture(directory + a + '\\' + b)
                frames = data.get(cv2.CAP_PROP_FRAME_COUNT)
                fps = data.get(cv2.CAP_PROP_FPS)
                seconds = round(frames / fps)
                positiveTime += seconds
                allTime += seconds
                fileCountList.append(strftime('%H:%M:%S', gmtime(seconds)))

            elif b.startswith('2'):
                variable = b.replace('2', '', 1).strip()
                addBroadcaster(fileBroadcasterList, variable)
                fileCategoriesList.append('Negativo')

                data = cv2.VideoCapture(directory + a + '\\' + b)
                frames = data.get(cv2.CAP_PROP_FRAME_COUNT)
                fps = data.get(cv2.CAP_PROP_FPS)
                seconds = round(frames / fps)
                negativeTime += seconds
                allTime += seconds
                fileCountList.append(strftime('%H:%M:%S', gmtime(seconds)))

            else:
                variable = b.replace('3', '', 1).strip()
                addBroadcaster(fileBroadcasterList, variable)
                fileCategoriesList.append('Neutro')

                data = cv2.VideoCapture(directory + a + '\\' + b)
                frames = data.get(cv2.CAP_PROP_FRAME_COUNT)
                fps = data.get(cv2.CAP_PROP_FPS)
                seconds = round(frames / fps)
                neutralTime += seconds
                allTime += seconds
                fileCountList.append(strftime('%H:%M:%S', gmtime(seconds)))
        else:
            print('Arquivo não é um vídeo')


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

allInfoCountList.append(fileCategoriesList.count('Positivo'))
allInfoCountList.append(fileCategoriesList.count('Negativo'))
allInfoCountList.append(fileCategoriesList.count('Neutro'))
allInfoCountList.append(strftime('%H:%M:%S', gmtime(positiveTime)))
allInfoCountList.append(strftime('%H:%M:%S', gmtime(negativeTime)))
allInfoCountList.append(strftime('%H:%M:%S', gmtime(neutralTime)))
allInfoCountList.append(strftime('%H:%M:%S', gmtime(allTime)))

positiveData = pd.DataFrame({'Positivo': positiveList, 'Duração': positiveListDuration, 'Emissora': positiveListBrodcaster})
negativeData = pd.DataFrame({'Negativo': negativeList, 'Duração': negativeListDuration, 'Emissora': negativeListBrodcaster})
neutralData = pd.DataFrame({'Neutro': neutralList, 'Duração': neutralListDuration, 'Emissora': neutralListBrodcaster})
geral = pd.DataFrame({'Títulos': fileList, 'Duração': fileCountList, 'Emissora': fileBroadcasterList})
resumo = pd.DataFrame({'Tipo': allInfoNameList, 'Info': allInfoCountList})

with pd.ExcelWriter(creds.excel_path2) as writer:
    resumo.to_excel(writer, sheet_name='Resumo', index=False)
    positiveData.to_excel(writer, sheet_name='Positivo', index=False)
    negativeData.to_excel(writer, sheet_name='Negativo', index=False)
    neutralData.to_excel(writer, sheet_name='Neutro', index=False)
    geral.to_excel(writer, sheet_name='Geral', index=False)

    worksheet1 = writer.sheets['Resumo']
    worksheet2 = writer.sheets['Positivo']
    worksheet3 = writer.sheets['Negativo']
    worksheet4 = writer.sheets['Neutro']
    worksheet5 = writer.sheets['Geral']
    
    worksheet1.set_column('A:B', 30)

    worksheet2.set_column('A:A', 40)
    worksheet2.set_column('B:D', 15)

    worksheet3.set_column('A:A', 40)
    worksheet3.set_column('B:D', 15)

    worksheet4.set_column('A:A', 40)
    worksheet4.set_column('B:D', 15)

    worksheet5.set_column('A:A', 40)
    worksheet5.set_column('B:D', 15)
