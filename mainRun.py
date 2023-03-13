import os
import time
import glob
import pandas as pd
from datetime import datetime
import G_PC, G_SP, Y_PC, Y_SP

def getMaster():
    basedir = os.path.dirname(os.path.abspath(__file__))
    master_files = glob.glob(os.path.join(basedir, 'SEO_Autobot.xlsx'))

    dataframe = pd.read_excel(master_files[0])
    kw_list = dataframe['Target_keyword'].tolist()

    return kw_list

def main():
    start = time.time()
    kw_list = getMaster()
    try:
       PC1 = G_PC.mainRun(kw_list)
    except Exception:
        raise
    
    try:
        SP1 = G_SP.mainRun(kw_list)
    except Exception:
        raise
    
    try:
        PC2 = Y_PC.mainRun(kw_list)
    except Exception:
        raise

    try:
        SP2 = Y_SP.mainRun(kw_list)
    except Exception:
        raise

    data = pd.concat([PC1, SP1, PC2, SP2], ignore_index=True).set_index('keyword')
    
    # filter url that contains rikunabi
    final_data = data[data['title'].str.contains('next|NEXT') == False]
    
    end = time.time()
    hours, rem = divmod(end-start, 3600)
    minutes, seconds = divmod(rem, 60)
    print("Time to completion: {:0>2}:{:0>2}:{:0>2.0f}".format(int(hours),int(minutes),int(seconds)))

    return final_data

final_data = main()

path = os.getcwd() + '\\Output\\'
currDate = datetime.today().strftime('%Y%m%d')
outName = f'SEO_Report_{currDate}.csv'

final_data.to_csv(f'{path}{outName}', encoding='cp932')
