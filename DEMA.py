import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
plt.style.use('seaborn')

from Status import *
from BackTest import BackTest


def plot_dema(stock, plt_df, plot_full_location, cur_descp):
    plt_df = plt_df.set_index(pd.DatetimeIndex(plt_df['Date'].values))

    #Plotting DEMA Chart.
    plt.figure(figsize=(12.2, 4.5))
    column_list = ['DEMA_short', 'DEMA_long', 'Close']
    plt_df[column_list].plot(figsize=(12.2, 6.4))
    plt.title('Close price for {} - {}'.format(stock, cur_descp))
    plt.xlabel('Price in INR.')
    plt.xlabel('Date')
    plt.legend(loc='upper left')
    plot_name = '{}_DEMA.png'.format(stock)
    plot_file = plot_full_location + '/' + stock + '/' + plot_name
    plt.savefig(plot_file, dpi=150)
    dema_hlink = '=HYPERLINK("{}","DEMA Plot")'.format(plot_file)
    plt.close('all')

    #Visually show the stock buy and sell signals.
    plt.figure(figsize=(12.2, 4.5))
    plt.scatter(plt_df.index, plt_df['Buy_Signal_Price'], color = 'green', label='Buy Signal', marker='^', alpha=1)
    plt.scatter(plt_df.index, plt_df['Sell_Signal_Price'], color = 'red', label='Sell Signal', marker='v', alpha=1)
    plt.plot(plt_df['Close'], label='Close Price', alpha=0.35)
    plt.plot(plt_df['DEMA_short'], label='DEMA_short', alpha=0.35)
    plt.plot(plt_df['DEMA_long'], label='DEMA_long', alpha=0.35)
    plt.xticks(rotation=45)
    plt.title('Close Price Buy and Sell Signals ({})'.format(stock))
    plt.xlabel('Date', fontsize=18)
    plt.ylabel('Close Price in INR', fontsize = 18)
    plt.legend(loc='upper left')
    plot_name = '{}_DEMA_Signals.png'.format(stock)
    plot_file = plot_full_location + '/' + stock + '/' + plot_name
    plt.savefig(plot_file, dpi=150)
    signal_hlink = '=HYPERLINK("{}","DEMA Signal Plot")'.format(plot_file)
    plt.close('all')

    return (dema_hlink, signal_hlink)


def DEMA_Calc(data, time_period, column):
    #Calculate the Exponential Moving Average for some time period.
    EMA = data[column].ewm(span=time_period, adjust=False).mean()
    #Calculate the DEMA.
    DEMA = 2 * EMA - EMA.ewm(span=time_period, adjust=False).mean()

    return DEMA


def buy_sell_DEMA(data):
    buy_list = []
    sell_list = []
    flag = False
    #Loop through the data.
    for i in range(0, len(data)):
        if data['DEMA_short'][i] > data['DEMA_long'][i] and flag == False:
            buy_list.append(data['Close'][i])
            sell_list.append(np.nan)
            flag = True
        elif data['DEMA_short'][i] < data['DEMA_long'][i] and flag == True:
            buy_list.append(np.nan)
            sell_list.append(data['Close'][i])
            flag = False
        else:
            buy_list.append(np.nan)
            sell_list.append(np.nan)

    #Store the buy and sell signals lists into the data set.
    data['Buy_Signal_Price'] = buy_list
    data['Sell_Signal_Price'] = sell_list
    
    return data


def DEMA(stock, input_var, plot_full_location, wb, cur_descp, stock_summary, stkplt_hlink):

    period_short_DEMA = input_var[0]
    period_long_DEMA = input_var[1]
    column = input_var[2]
##    pd.set_option('display.max_columns', None)
##    df_DEMA = input_var[3]
    df = input_var[3]
    df_DEMA = pd.DataFrame()
    columns = ['Date', 'Open', 'High', 'Low', 'Close', 'Adj Close', 'Volume']
    for col in columns:
        df_DEMA[col] = df[col]

    #Store the short term DEMA (20 day period) and the long term DEMA (50 day period) into the data set.
    df_DEMA['DEMA_short'] = DEMA_Calc(df_DEMA, period_short_DEMA, column)
    df_DEMA['DEMA_long'] = DEMA_Calc(df_DEMA, period_long_DEMA, column)

    df_DEMA = buy_sell_DEMA(df_DEMA)

    strategy_hlinks = plot_dema(stock, df_DEMA, plot_full_location, cur_descp)
    stock_status = stock_current_status(df_DEMA)
    bt_data = BackTest('DEMA', stock, df_DEMA, plot_full_location, wb, cur_descp)
    print_report(stock, cur_descp, df_DEMA, bt_data, stock_summary,
                 stock_status, stkplt_hlink, strategy_hlinks)

    return stock_status
