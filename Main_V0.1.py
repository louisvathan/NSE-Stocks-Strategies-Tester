import pandas as pd
import os
import time
import datetime
from datetime import timezone
from datetime import datetime
import requests
from openpyxl import Workbook
import yfinance as yf

start = time.time()

from Strategy_Management import *
from Excel_Processing import full_summary
import Stocks

now = datetime.now()
date = str(now.year) + '{:02d}'.format(now.month) + '{:02d}'.format(now.day)

stocks = Stocks.mystocks
period1 = str(int(datetime((int(now.year)-5), (int(now.month)), (int(now.day)), 0, 0, tzinfo=timezone.utc).timestamp()))
period2 = str(int(datetime((int(now.year)), (int(now.month)), (int(now.day)), 0, 0, tzinfo=timezone.utc).timestamp()))

#CREATING MAIN AND CSV FOLDERS.
current_location = os.getcwd()

foldername = 'Results(' + date + ')'
csv_location = foldername + '/CSV FILES'
plot_location = foldername + '/PLOTS'
pkl_location = foldername + '/PICKLE'


folders = [foldername, csv_location, plot_location, pkl_location]

for f in folders:
    try:
        os.mkdir(f)
    except:
        pass

csv_full_location = current_location + '/' + csv_location
plot_full_location = current_location + '/' + plot_location
pkl_full_location = current_location + '/' + pkl_location
folder_full_location = current_location + '/' + foldername

summary = pd.DataFrame(index=stocks)

for stock in stocks:
    csv_name = '{}-{}-{}.csv'.format(stock, period1, period2)
    csv_file = csv_full_location + '/' + csv_name

    pkl_name = '{}-{}-{}.pkl'.format(stock, period1, period2)
    pkl_file = pkl_full_location + '/' + pkl_name

    for root, dir, files in os.walk(pkl_full_location):
        if pkl_name in files:
            df_lt = pd.read_pickle(pkl_file)
            delay = 0
        else:
            df_lt = yf.download(stock+'.NS', start="{}-{}-{}".format(int(now.year)-5, int(now.month), int(now.day)),
                                    end="{}-{}-{}".format(int(now.year), int(now.month), int(now.day))).to_csv(csv_file)
            df_lt = pd.read_csv(csv_file)
            df_lt = df_lt.dropna()
            df_lt = df_lt.reset_index()
            
            if df_lt.columns[0] == 'index':
                df_lt.drop(['index'], axis=1, inplace=True)
            elif df_lt.columns[0] == 'level_0':
                df_lt.drop(['level_0'], axis=1, inplace=True)
            df_lt.to_pickle(pkl_file)
            print(df_lt)
            delay = 2

    df_st = df_lt.tail(150)
    df_st = df_st.reset_index()
    df_st.drop(['index'], axis=1, inplace=True)
    
    #Creating Excel Document in the name of the current Stock.
    wb = Workbook()
    
    summary = Strategy_Management(stock, df_st, df_lt, plot_full_location, wb, summary)
    
    wb.save('{}/{}/{}.xlsx'.format(current_location, foldername, stock))

    print('{} Done.'.format(stock))

    time.sleep(delay)

wb_summary = Workbook()
full_summary(wb_summary, summary, date, folder_full_location)
wb_summary.save('{}/{}/Summary ({}).xlsx'.format(current_location, foldername, date))

end = time.time()
print(f"Runtime of the program is {end - start}")
