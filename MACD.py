import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
plt.style.use('seaborn')

from Status import *
from BackTest import BackTest


def plot_macd(stock, plt_df, plot_full_location, cur_descp):
    plt_df = plt_df.set_index(pd.DatetimeIndex(plt_df['Date'].values))

    #Plotting MACD chart.
    plt.figure(figsize=(12.2, 4.5))
    plt.plot(plt_df.index, plt_df['MACD'], label = '{} MACD'.format(stock), color = 'red')
    plt.plot(plt_df.index, plt_df['Signal Line'], label = 'Signal Line', color = 'blue')
    plt.title('{} - {}'.format(stock, cur_descp))
    plt.legend(loc='upper left')
    plot_name = '{}_MACD.png'.format(stock)
    plot_file = plot_full_location + '/' + stock + '/' + plot_name
    plt.savefig(plot_file, dpi=150)
    macd_hlink = '=HYPERLINK("{}","MACD Plot")'.format(plot_file)
    plt.close('all')

    #Plotting Buy & Sell Signals.
    plt.figure(figsize=(12.2, 4.5))
    plt.scatter(plt_df.index, plt_df['Buy_Signal_Price'], color='green', label='Buy', marker='^', alpha=1)
    plt.scatter(plt_df.index, plt_df['Sell_Signal_Price'], color='red', label='Sell', marker='v', alpha=1)
    plt.plot(plt_df['Close'], label='Close Price', alpha=0.35)
    plt.title('Close Price Buy & Sell Signals ({} - {})'.format(stock, cur_descp))
    plt.xlabel('Date')
    plt.ylabel('Close Price in INR')
    plt.legend(loc='upper left')
    plot_name = '{}_MACD_Signals.png'.format(stock)
    plot_file = plot_full_location + '/' + stock + '/' + plot_name
    plt.savefig(plot_file, dpi=150)
    signal_hlink = '=HYPERLINK("{}","MACD Signal Plot")'.format(plot_file)
    plt.close('all')

    return (macd_hlink, signal_hlink)


def buy_sell(signal):
    #Creating Buy and Sell Signals.
    Buy = []
    Sell = []
    flag = -1
    
    for i in range(0, len(signal)):
        if signal['MACD'][i] > signal['Signal Line'][i]:
            Sell.append(np.nan)
            if flag != 1:
                Buy.append(signal['Close'][i])
                flag = 1
            else:
                Buy.append(np.nan)
        elif signal['MACD'][i] < signal['Signal Line'][i]:
            Buy.append(np.nan)
            if flag != 0:
                Sell.append(signal['Close'][i])
                flag = 0
            else:
                Sell.append(np.nan)
        else:
            Buy.append(np.nan)
            Sell.append(np.nan)
    
    return(Buy, Sell)


def MACD(stock, input_var, plot_full_location, wb, cur_descp, stock_summary, stkplt_hlink):
    macd_short_period = input_var[0]
    macd_long_period = input_var[1]
    macd_signal_period = input_var[2]
    df_MACD = input_var[3]
    #Calculate the MACF and signal line indicators.
    #Calculate the Short Term Exponential Moving Average (EMA)
    ShortEMA = df_MACD.Close.ewm(span=macd_short_period, adjust=False).mean()
    #Calculate the Long Term Exponential Moving Average (EMA)
    LongEMA = df_MACD.Close.ewm(span=macd_long_period, adjust=False).mean()
    #Calsulate the MACD line.
    MACD = ShortEMA - LongEMA
    #Calsulate the signal line.
    signal = MACD.ewm(span=macd_signal_period, adjust=False).mean()

    df_MACD['MACD'] = MACD
    df_MACD['Signal Line'] = signal

    a = buy_sell(df_MACD)
    df_MACD['Buy_Signal_Price'] = a[0]
    df_MACD['Sell_Signal_Price'] = a[1]

    strategy_hlinks = plot_macd(stock, df_MACD, plot_full_location, cur_descp)
    stock_status = stock_current_status(df_MACD)
    bt_data = BackTest('MACD', stock, df_MACD, plot_full_location, wb, cur_descp)
    print_report(stock, cur_descp, df_MACD, bt_data, stock_summary,
                 stock_status, stkplt_hlink, strategy_hlinks)

    return stock_status
