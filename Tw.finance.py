import os

import numpy
import requests
import datetime
import time
import math
import pandas as pd
import functions
import xlwt
import numpy as np
from tqdm import tqdm
import gspread
from gspread_dataframe import set_with_dataframe
from oauth2client.service_account import ServiceAccountCredentials

debug_mode = False
save_local_file = False
jump_phase_two = False
start_index = 800

currentDate = datetime.datetime.utcnow()
dateStr = currentDate.strftime("%Y-%m-%d") if not debug_mode else "Debug-" + currentDate.strftime("%Y-%m-%d")

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('tw-finance-f09c6b5d4a8c.json', scope)
gc = gspread.authorize(credentials)
sh = gc.open('Tw-finance')

try:
    if debug_mode:
        try:
            ws = sh.worksheet(dateStr)
            sh.del_worksheet(ws)
            print("Delete exist sheet: " + dateStr)
        except:
            print("Create new sheet: " + dateStr)

    ws = sh.add_worksheet(title=dateStr, rows='1000', cols='12')
except Exception as e:
    print(e)
    print("Cannot add worksheet. Please check if the sheet already exist.")
    exit(1)


pbar = tqdm(total=972)
now = datetime.datetime.now()
dayStart = str(int(time.time()))
dayEnd = str(int(time.time()) - 8640000)
monthEnd = str(int(time.time()) - 686400000)

all = functions.readAll()

resultDic = {}
idArr = []
tempArr = []
nameArr = []
dayWilliamsRArr = []
dayRSIArr = []
monthRSIArr = []
monthMTMArr = []
monthDMIArr_plus = []
monthDMIArr_minus = []
process = 0

# print(all.keys())

for value in all.values:
    pbar.update(1)

    if debug_mode and pbar.n < start_index:
        continue

    tempArr.append(value[0])
    nameArr.append(value[1])

    responseDay = functions.getFinanceData(value[0], dayStart, dayEnd, "1d")
    try:
        dataArrayDay = functions.dataTextToArray(responseDay.text)
    except:
        sh.del_worksheet(ws)
        print()
        print("ERROR: dataTextToArray responseDay. Invalid cookie.")
        break #exit(1)

    arrWilliamsR = functions.arrayWilliamsR(dataArrayDay, 50)
    arrRSI = functions.arrayRSI(dataArrayDay, 60)

    dayWilliamsR = arrWilliamsR[len(arrWilliamsR) - 1][9]
    dayRSI = arrRSI[len(arrRSI) - 1][7]

    dayWilliamsRArr.append(dayWilliamsR)
    dayRSIArr.append(dayRSI)

    responseMonth = functions.getFinanceData(value[0], dayStart, monthEnd, "1mo")

    try:
        dataArrayMonth = functions.dataTextToArray(responseMonth.text)
    except:
        sh.del_worksheet(ws)
        print()
        print("ERROR: dataTextToArray responseMonth. Invalid cookie.")
        break #exit(1)

    arrSize = len(dataArrayMonth)
    if arrSize >= 2:
        if dataArrayMonth[arrSize - 1][2] < dataArrayMonth[arrSize - 2][2]:
            dataArrayMonth[arrSize - 1][2] = dataArrayMonth[arrSize - 2][2]
        if dataArrayMonth[arrSize - 1][3] > dataArrayMonth[arrSize - 2][3]:
            dataArrayMonth[arrSize - 1][3] = dataArrayMonth[arrSize - 2][3]

    dataArrayMonth = np.delete(dataArrayMonth, len(dataArrayMonth) - 2, axis=0)

    # print (responseMonth.text)
    # print (dataArrayMonth)

    arrRSIMonth = functions.arrayRSI(dataArrayMonth, 4)
    arrDMIMonth = functions.arrayDMI(dataArrayMonth, 1)
    arrMTMMonth = functions.arrayMTM(dataArrayMonth, 3, 2)
    if len(arrRSIMonth) <= 1:
        monthRSI = None
    else:
        monthRSI = arrRSIMonth[len(arrRSIMonth) - 1][7]

    if len(arrDMIMonth) <= 1:
        monthDMI = None
    else:
        monthDMI_plus = arrDMIMonth[len(arrDMIMonth) - 1][7]
        monthDMI_minus = arrDMIMonth[len(arrDMIMonth) - 1][8]

    if len(arrMTMMonth) <= 1:
        monthMTM = None
    else:
        monthMTM = arrMTMMonth[len(arrMTMMonth) - 1][9]

    monthRSIArr.append(monthRSI)
    monthMTMArr.append(monthMTM)
    monthDMIArr_plus.append(monthDMI_plus)
    monthDMIArr_minus.append(monthDMI_minus)

    process = process + 1

    if debug_mode and process > 30:
        break

resultDic['monthRSI'] = monthRSIArr
resultDic['monthMTM'] = monthMTMArr
resultDic['monthDMI_plus'] = monthDMIArr_plus
resultDic['monthDMI_minus'] = monthDMIArr_minus
resultDic['dayRSI'] = dayRSIArr
resultDic['dayWilliamsR'] = dayWilliamsRArr
resultDic[all.keys()[1]] = nameArr
resultDic[all.keys()[0]] = tempArr

resultDF = pd.DataFrame(resultDic)
pbar.close()
# print (resultDF)


resultDF = resultDF.reindex(
    columns=['證券代號', '證券名稱', 'dayWilliamsR', 'dayRSI', 'monthRSI', 'monthDMI_plus', 'monthDMI_minus', 'monthMTM'])
accordDic = resultDF[resultDF.monthRSI > 77]
accordDic = accordDic[accordDic.dayRSI > 57]
accordDic = accordDic[accordDic.dayWilliamsR < 20]

# print(accordDic)

if save_local_file:
    resultDF.to_excel('all_results_last.xls', sheet_name=dateStr)
    functions.append_df_to_excel('log_results.xlsx', accordDic, sheet_name=dateStr, index=False)

set_with_dataframe(ws, accordDic, row=1, col=1, include_index=True, include_column_header=True)
# print(accordDic)

listMACDWeekDiff = []
listMACDWeekDirection = []

pbar_MACD = tqdm(total=len(accordDic))


for index, row in accordDic.iterrows():
    # print(index, row['證券代號'], row['證券名稱'])
    responseWeek = functions.getFinanceData(row['證券代號'], dayStart, monthEnd, "1mo")

    try:
        dataArrayWeek = functions.dataTextToArray(responseWeek.text)
    except:
        # sh.del_worksheet(ws)
        print()
        print("ERROR: dataTextToArray responseMonth. Invalid cookie.")
        exit(1)

    arrMACDWeek = functions.arrayMACD(dataArrayWeek, 12, 26, 9)
    if len(arrMACDWeek)>0:
        #print(arrMACDWeek[len(arrMACDWeek)-1])
        listMACDWeekDiff.append(arrMACDWeek[len(arrMACDWeek)-1][9])
        listMACDWeekDirection.append(arrMACDWeek[len(arrMACDWeek)-1][10])
    pbar_MACD.update(1)

accordDic['MACD_Diff'] = list(pd.Series(listMACDWeekDiff))
accordDic['MACD_Direction'] = list(pd.Series(listMACDWeekDirection))

#print(accordDic)
set_with_dataframe(ws, accordDic, row=1, col=1, include_index=True, include_column_header=True)

pbar_MACD.close()
