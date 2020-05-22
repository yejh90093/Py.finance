import datetime
import numpy
import requests
from io import StringIO
import pandas as pd
import numpy as np


def readAll():
    currentDate = datetime.datetime.utcnow()
    condition = True
    while condition:
        # loop body here
        currentDate = currentDate - datetime.timedelta(days=1)
        currentDateStr = currentDate.strftime("%Y%m%d")
        r = requests.post(
            'http://www.twse.com.tw/exchangeReport/MI_INDEX?response=csv&date=' + currentDateStr + '&type=ALL')
        if r.text != "":
            condition = False

    str_list = []
    for i in r.text.split('\n'):
        if len(i.split('",')) == 17 and i[0] != '=':
            i = i.strip(",\r\n")
            str_list.append(i)

    df = pd.read_csv(StringIO("\n".join(str_list)))
    pd.set_option('display.max_rows', None)
    return df


# Valid intervals: [1m, 2m, 5m, 15m, 30m, 60m, 90m, 1h, 1d, 5d, 1wk, 1mo, 3mo]
def getFinanceData(id, start, end, interval):
    site = "https://query1.finance.yahoo.com/v7/finance/download/" + id + ".TW?" \
                                                                          "period1=" + end + \
           "&period2=" + start + \
           "&interval=" + interval + \
           "&events=history" \
           "&crumb=hP2rOschxO0"
    response = requests.get(site)
    # print(site)
    return response


def dataTextToArray(text):
    df = pd.read_csv(StringIO(text))
    return df.values


def arrayWilliamsR(dataArray, nDay):
    arrWilliamsR = []
    NH = []
    NL = []
    n = 0
    nDay = nDay - 1

    for day in dataArray:

        if n >= nDay:
            NH = []
            NL = []
            for i in range(n, n - nDay, -1):
                NH.append(dataArray[i][2])
                NL.append(dataArray[i][3])

        if len(NH) == 0:
            B = dataArray[n]

            A = numpy.array([None, None, None])
            B = numpy.append(B, A)
        else:
            B = dataArray[n]
            W = (max(NH) - dataArray[n][4]) / (max(NH) - min(NL)) * 100

            A = numpy.array([max(NH), min(NL), W])

            B = numpy.append(B, A)

        # print (n, dataArray[n][0], A)
        arrWilliamsR.append(B)
        n = n + 1

    return arrWilliamsR


def arrayRSI(dataArray, nDay):
    arrRSI = []
    n = 0
    nDay = nDay - 1

    lastPlus = None
    lastMinus = None

    for day in dataArray:
        sumPlus = 0
        sumMinus = 0

        if n >= nDay:

            if n == nDay:
                for i in range(n, n - nDay, -1):
                    dayDiff = dataArray[i][4] - dataArray[i - 1][4]
                    if numpy.math.isnan(dayDiff): dayDiff = 0

                    if dayDiff < 0:
                        sumMinus = sumMinus + abs(dayDiff)
                    else:
                        sumPlus = sumPlus + abs(dayDiff)
            else:
                dayDiff = dataArray[n][4] - dataArray[n - 1][4]
                if numpy.math.isnan(dayDiff): dayDiff = 0

                if dayDiff < 0:
                    sumPlus = (lastPlus * nDay / (nDay + 1))
                    sumMinus = (lastMinus * nDay / (nDay + 1)) + (abs(dayDiff) / (nDay + 1))
                else:
                    sumPlus = (lastPlus * nDay / (nDay + 1)) + (abs(dayDiff) / (nDay + 1))
                    sumMinus = (lastMinus * nDay / (nDay + 1))

            lastPlus = sumPlus
            lastMinus = sumMinus

            avgPlus = sumPlus / (nDay + 1)
            avgMinus = sumMinus / (nDay + 1)

            if avgPlus + avgMinus == 0:
                dayRSI = None
                A = numpy.array([None])
            else:
                dayRSI = avgPlus / (avgPlus + avgMinus)
                A = numpy.array([dayRSI * 100])

            B = dataArray[n]
            B = numpy.append(B, A)
            arrRSI.append(B)
        else:
            B = dataArray[n]
            A = numpy.array([None])
            B = numpy.append(B, A)
            arrRSI.append(B)

        # print (n, dataArray[n][0], A)
        n = n + 1

    return arrRSI
