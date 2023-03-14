import os
import time
import glob
import pandas as pd
from datetime import datetime
import G_PC, G_SP, Y_PC, Y_SP

def concat(*args):
    # function concat for dataFrame
    return pd.concat([x for x in args if not x.empty], ignore_index=True).set_index('keyword')

def getMaster():
    # get target keywords from master file and put into list
    basedir = os.path.dirname(os.path.abspath(__file__))
    master_files = glob.glob(os.path.join(basedir, 'SEO_Autobot.xlsx'))

    dataFrame = pd.read_excel(master_files[0], sheet_name='Master', header=6, usecols='C:G')
    GPC_list = []
    GSP_list = []
    YPC_list = []
    YSP_list = []

    for ind in dataFrame.index:
        if dataFrame['PC'][ind] == 'ON':
            GPC_list.append(dataFrame['Target_keyword'][ind])

        if dataFrame['SP'][ind] == 'ON':
            GSP_list.append(dataFrame['Target_keyword'][ind])

        if dataFrame['PC.1'][ind] == 'ON':
            YPC_list.append(dataFrame['Target_keyword'][ind])

        if dataFrame['SP.1'][ind] == 'ON':
            YSP_list.append(dataFrame['Target_keyword'][ind])

    return {
            'GPC_list': GPC_list, 
            'GSP_list': GSP_list, 
            'YPC_list': YPC_list, 
            'YSP_list': YSP_list,
            }

def makeData(kw_list):
    # from lists of keyword, get ads data
    if len(kw_list['GPC_list']) > 0:
        try:
            GPC_list = G_PC.mainRun(kw_list['GPC_list'])
        except Exception:
            print('Error at Google PC!')
            pass
    else:
        print('Google PC is empty!')
        GPC_list = pd.DataFrame()

    if len(kw_list['GSP_list']) > 0:
        try:
            GSP_list = G_SP.mainRun(kw_list['GSP_list'])
        except Exception:
            print('Error at Google SP!')
            pass
    else:
        print('Google SP is empty!')
        GSP_list = pd.DataFrame()

    if len(kw_list['YPC_list']) > 0:
        try:
            YPC_list = Y_PC.mainRun(kw_list['YPC_list'])
        except Exception:
            print('Error at Yahoo PC!')
            pass
    else:
        print('Yahoo PC is empty!')
        YPC_list = pd.DataFrame()

    if len(kw_list['YSP_list']) > 0:
        try:
            YSP_list = Y_SP.mainRun(kw_list['YSP_list'])
        except Exception:
            print('Error at Yahoo SP!')
            pass
    else:
        print('Yahoo PC is empty!')
        YSP_list = pd.DataFrame()
    
    return [GPC_list, GSP_list, YPC_list, YSP_list]

def saveToCSV(data):
    # save dataFrame to csv
    dataDF = concat(*[data[0], data[1], data[2], data[3]])
    
    # filter url that contains rikunabi
    final_data = dataDF[dataDF['title'].str.contains('next|NEXT') == False]   

    path = os.getcwd() + '\\Output\\'
    try:
        os.makedirs(path)
    except FileExistsError:
        pass

    currDate = datetime.today().strftime('%Y%m%d')
    outName = f'SEO_Report_{currDate}.csv'

    final_data.to_csv(f'{path}{outName}', encoding='cp932')

    return final_data

def masterRun():
    # start timer
    start = time.time()

    # Get list keywords from master file
    kw_list = getMaster()

    # make dataFrame from list of keywords
    data = makeData(kw_list)

    # save data to CSV
    final_data = saveToCSV(data)

    end = time.time()
    hours, rem = divmod(end-start, 3600)
    minutes, seconds = divmod(rem, 60)
    print("Time to completion: {:0>2}:{:0>2}:{:0>2.0f}".format(int(hours),int(minutes),int(seconds)))
