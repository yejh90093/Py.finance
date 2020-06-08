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
monthDMIArr_plus = []
monthDMIArr_minus = []
process = 0

# print(all.keys())

for value in all.values:
    # value[0] = "4142"
    # print (process, value[0], value[1])

    tempArr.append(value[0])
    nameArr.append(value[1])

    responseDay = functions.getFinanceData(value[0], dayStart, dayEnd, "1d")
    dataArrayDay = functions.dataTextToArray(responseDay.text)
    # print (dataArrayDay)

    arrWilliamsR = functions.arrayWilliamsR(dataArrayDay, 50)
    arrRSI = functions.arrayRSI(dataArrayDay, 60)

    dayWilliamsR = arrWilliamsR[len(arrWilliamsR) - 1][9]
    dayRSI = arrRSI[len(arrRSI) - 1][7]

    dayWilliamsRArr.append(dayWilliamsR)
    dayRSIArr.append(dayRSI)

    responseMonth = functions.getFinanceData(value[0], dayStart, monthEnd, "1mo")
    # responseMonth = functions.getFinanceData('2330', dayStart, monthEnd, "1mo")
    dataArrayMonth = functions.dataTextToArray(responseMonth.text)

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
    if len(arrRSIMonth) <= 1:
        monthRSI = None
    else:
        monthRSI = arrRSIMonth[len(arrRSIMonth) - 1][7]

    if len(arrDMIMonth) <= 1:
        monthDMI = None
    else:
        monthDMI_plus = arrDMIMonth[len(arrDMIMonth) - 1][7]
        monthDMI_minus = arrDMIMonth[len(arrDMIMonth) - 1][8]

    monthRSIArr.append(monthRSI)
    monthDMIArr_plus.append(monthDMI_plus)
    monthDMIArr_minus.append(monthDMI_minus)

    process = process + 1
    pbar.update(1)

    # if process > 100:
    #    break

resultDic['monthRSI'] = monthRSIArr
resultDic['monthDMI_plus'] = monthDMIArr_plus
resultDic['monthDMI_minus'] = monthDMIArr_minus
resultDic['dayRSI'] = dayRSIArr
resultDic['dayWilliamsR'] = dayWilliamsRArr
resultDic[all.keys()[1]] = nameArr
resultDic[all.keys()[0]] = tempArr

resultDF = pd.DataFrame(resultDic)
pbar.close()
# print (resultDF)


currentDate = datetime.datetime.utcnow()
currentDateStr = currentDate.strftime("%Y-%m-%d")
resultDF = resultDF.reindex(
    columns=['證券代號', '證券名稱', 'dayWilliamsR', 'dayRSI', 'monthRSI', 'monthDMI_plus', 'monthDMI_minus'])

resultDF.to_excel('all_results_last.xls', sheet_name=currentDateStr)

accordDic = resultDF[resultDF.monthRSI > 77]
accordDic = accordDic[accordDic.dayRSI > 57]
accordDic = accordDic[accordDic.dayWilliamsR < 20]

accordDic = accordDic.reindex(
    columns=['證券代號', '證券名稱', 'dayWilliamsR', 'dayRSI', 'monthRSI', 'monthDMI_plus', 'monthDMI_minus'])
pd.set_option('display.max_columns', 8)
print(accordDic)
functions.append_df_to_excel('log_results.xlsx', accordDic, sheet_name=currentDateStr, index=False)
