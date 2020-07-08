import webbrowser

import Define_variable


def callback(url):
    webbrowser.open_new(url)


def checkSheetExist(sheetList, sheetName):
    exist = False
    for worksheet in sheetList:
        if worksheet.title == sheetName:
            exist = True
    return exist


def normal_run():
    Define_variable.result_label.configure(text="Normal Run Start")
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

    debug_mode = Define_variable.var_debugMode.get()
    save_local_file = Define_variable.var_saveLocal.get()
    overwrite_file = Define_variable.var_overwrite.get()
    start_index = int(Define_variable.start_index_entry.get())
    assign_start = start_index != 0

    insert_row_index = 1

    currentDate = datetime.datetime.utcnow()
    dateStr = currentDate.strftime("%Y-%m-%d") if not debug_mode else "Debug-" + currentDate.strftime("%Y-%m-%d")

    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name('tw-finance-f09c6b5d4a8c.json', scope)
    gc = gspread.authorize(credentials)
    sh = gc.open('Tw-finance')
    sheetExist = checkSheetExist(sh.worksheets(), dateStr);

    try:

        if debug_mode:
            print("debug_mode")
            try:
                ws = sh.worksheet(dateStr)
                sh.del_worksheet(ws)
                print("Delete exist sheet: " + dateStr)
            except:
                print("Create new sheet: " + dateStr)

        if sheetExist and overwrite_file:
            print(("sheetExist and overwrite_file"))
            ws = sh.worksheet(dateStr)
            sh.del_worksheet(ws)

        if assign_start:
            print(("assign_start"))
            insert_row_index = len(ws.get_all_records()) + 2

        ws = sh.add_worksheet(title=dateStr, rows='1000', cols='12')
        ws_url = "https://docs.google.com/spreadsheets/d/1eFs26OMNcxY-t2r73WPq14J4qCpyq9Ezv-Ec6vDmBiU/edit#gid=" + str(
            ws.id)
        Define_variable.window_label.bind("<Button-1>", lambda e: callback(ws_url))

    except Exception as e:
        print(e)
        print("Cannot add worksheet. Please check if the sheet already exist.")
        exit(1)

    total = 972  # - start_index
    pbar = tqdm(total=total)
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
            pbar.update(-1)
            continue

        if assign_start and pbar.n < start_index:
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
            break  # exit(1)

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
            break  # exit(1)

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

    set_with_dataframe(ws, accordDic, row=insert_row_index, col=1, include_index=True, include_column_header=True)
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
        if len(arrMACDWeek) > 0:
            # print(arrMACDWeek[len(arrMACDWeek)-1])
            listMACDWeekDiff.append(arrMACDWeek[len(arrMACDWeek) - 1][9])
            listMACDWeekDirection.append(arrMACDWeek[len(arrMACDWeek) - 1][10])
        pbar_MACD.update(1)

    accordDic['MACD_Diff'] = list(pd.Series(listMACDWeekDiff))
    accordDic['MACD_Direction'] = list(pd.Series(listMACDWeekDirection))

    if save_local_file:
        resultDF.to_excel('all_results_last.xls', sheet_name=dateStr)
        functions.append_df_to_excel('log_results.xlsx', accordDic, sheet_name=dateStr, index=False)

    set_with_dataframe(ws, accordDic, row=insert_row_index, col=1, include_index=True, include_column_header=True)

    pbar_MACD.close()
