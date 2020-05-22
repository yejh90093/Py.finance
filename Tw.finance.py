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
process = 0

#print(all.keys())

for value in all.values:
    #value[0] = "4142"
    #print (process, value[0], value[1])

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
    dataArrayMonth = np.delete(dataArrayMonth, len(dataArrayMonth) - 2, axis=0)

    # print (dataArrayMonth)

    arrRSIMonth = functions.arrayRSI(dataArrayMonth, 4)
    if len(arrRSIMonth) <= 1:
        monthRSI = None
    else:
        monthRSI = arrRSIMonth[len(arrRSIMonth) - 1][7]

    monthRSIArr.append(monthRSI)

    process = process + 1
    pbar.update(1)

    # if process > 10:
    #    break


resultDic['monthRSI'] = monthRSIArr
resultDic['dayRSI'] = dayRSIArr
resultDic['dayWilliamsR'] = dayWilliamsRArr
resultDic[all.keys()[1]] = nameArr
resultDic[all.keys()[0]] = tempArr

resultDF = pd.DataFrame(resultDic)
pbar.close()

# print (resultDF)
resultDF.to_excel('excel_output.xls',sheet_name='biubiu')

accordDic = resultDF[resultDF.monthRSI > 77]
accordDic = accordDic[accordDic.dayRSI > 57]
accordDic = accordDic[accordDic.dayWilliamsR < 20]

print (accordDic)