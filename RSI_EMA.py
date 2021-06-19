import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
plt.style.use('seaborn')

from Status import *
from BackTest import BackTest


def plot_rsi_ema(stock, plt_df, plot_full_location, cur_descp, period_rsi,
                 rsi_low, rsi_high, period_short, period_long, column):
    plt_df = plt_df.set_index(pd.DatetimeIndex(plt_df['Date'].values))

    ax = ('ax1', 'ax2')
    fig, ax = plt.subplots(nrows=2, ncols=1, sharex=True, gridspec_kw={'height_ratios': [1, 1]})
    
    #Plotting RSI and EMA Charts.
    ax[0].set_title(cur_descp)
    ax[0].plot(plt_df.index, plt_df['RSI'])
    ax[0].axhline(0, linestyle='--', alpha = 0.5, color='gray')
    ax[0].axhline(10, linestyle='--', alpha = 0.5, color='orange')
    ax[0].axhline(20, linestyle='--', alpha = 0.5, color='green')
    ax[0].axhline(30, linestyle='--', alpha = 0.5, color='red')
    ax[0].axhline(70, linestyle='--', alpha = 0.5, color='red')
    ax[0].axhline(80, linestyle='--', alpha = 0.5, color='green')
    ax[0].axhline(90, linestyle='--', alpha = 0.5, color='red')
    ax[0].axhline(100, linestyle='--', alpha = 0.5, color='gray')
    ax[1].set_ylabel('RSI')

    ax[1].plot(plt_df.index, plt_df[column], label=column)
    ax[1].plot(plt_df.index, plt_df['EMA{}'.format(period_short)], label='EMA{}'.format(period_short))
    ax[1].plot(plt_df.index, plt_df['EMA{}'.format(period_long)], label='EMA{}'.format(period_long))
    ax[1].set_xlabel('{} to {}'.format(plt_df['Date'][0], plt_df['Date'].iloc[-1]))
    ax[1].legend(title="{} - EMA {} & {}".format(stock, period_short, period_long), fontsize='small', fancybox=True)
    
    for a in ax:
        a.patch.set_edgecolor('black')  
        a.patch.set_linewidth('1')
        a.yaxis.set_label_position("right")
        a.yaxis.tick_right()

    plt.subplots_adjust(wspace=0, hspace=0)
    plot_name = '{}_RSI_EMA.png'.format(stock)
    plot_file = plot_full_location + '/' + stock + '/' + plot_name
    plt.savefig(plot_file, dpi=150)
    rsi_ema_hlink = '=HYPERLINK("{}","RSI_EMA Plot")'.format(plot_file)
    plt.close('all')

    #Plot the Buy and Sell Prices.
    plt.figure(figsize=(12.2, 4.5))
    plt.plot(plt_df[column], label=column, alpha=0.35)
    if not plt_df['Buy_Signal_Price'].isnull().all():
        plt.scatter(plt_df.index, plt_df['Buy_Signal_Price'], label='Buy_Signal', marker='^', alpha=1, color='green')
    if not plt_df['Sell_Signal_Price'].isnull().all():
        plt.scatter(plt_df.index, plt_df['Sell_Signal_Price'], label='Sell_Signal', marker='v', alpha=1, color='red')
    plt.title('Close Price Buy & Sell Signals ({} - {})'.format(stock, cur_descp))
    plt.xlabel('{} to {}'.format(plt_df['Date'][0], plt_df['Date'].iloc[-1]))
    plt.ylabel('Close Price in INR')
    plt.legend(loc='upper left')
    plot_name = '{}_RSI_EMA_Signals.png'.format(stock)
    plot_file = plot_full_location + '/' + stock + '/' + plot_name
    plt.savefig(plot_file, dpi=150)
    signal_hlink = '=HYPERLINK("{}","RSI_EMA Signal Plot")'.format(plot_file)
    plt.close('all')

    return (rsi_ema_hlink, signal_hlink)


def buy_sell_rsi_ema(data, period_short, period_long, rsi_low, rsi_high, column):
    sigPriceBuy = []
    sigPriceSell = []
    flag = -1

    for i in range(len(data)):
        if data['EMA{}'.format(period_short)][i] > data['EMA{}'.format(period_long)][i]:
            if data['RSI'][i] < rsi_low:
                if flag != 1:
                    sigPriceBuy.append(data[column][i])
                    sigPriceSell.append(np.nan)
                    flag = 1
                else:
                    sigPriceBuy.append(np.nan)
                    sigPriceSell.append(np.nan)
            else:
                sigPriceBuy.append(np.nan)
                sigPriceSell.append(np.nan)
        elif data['EMA{}'.format(period_short)][i] < data['EMA{}'.format(period_long)][i]:
            if data['RSI'][i] > rsi_high:
                if flag != 0:
                    sigPriceBuy.append(np.nan)
                    sigPriceSell.append(data[column][i])
                    flag = 0
                else:
                    sigPriceBuy.append(np.nan)
                    sigPriceSell.append(np.nan)
            else:
                sigPriceBuy.append(np.nan)
                sigPriceSell.append(np.nan)
        else:
            sigPriceBuy.append(np.nan)
            sigPriceSell.append(np.nan)
            
    return (sigPriceBuy, sigPriceSell)


def RSI(period_rsi, column, df_RSI_EMA):
    #Intake dataframe and converts into RSI files(new_df).
    delta = df_RSI_EMA[column].diff(1)
    delta = delta.dropna()

    up = delta.copy()
    down = delta.copy()

    up[up<0] = 0
    down[down>0] = 0

    avg_gain = up.rolling(window=period_rsi).mean()
    avg_loss = abs(down.rolling(window=period_rsi).mean())

    rs = avg_gain/avg_loss
    rsi = 100.0 - (100.0 / (1.0 + rs))

    df_RSI_EMA['RSI'] = rsi
    
    return df_RSI_EMA


def EMA(df_RSI_EMA, time_period, column):
    #Calculate the Exponential Moving Average for some time period.
    EMA = df_RSI_EMA[column].ewm(span=time_period, adjust=False).mean()

    return EMA


def RSI_EMA(stock, input_var, plot_full_location, wb, cur_descp, stock_summary, stkplt_hlink):
    period_rsi = input_var[0]
    rsi_low = input_var[1]
    rsi_high = input_var[2]
    period_short = input_var[3]
    period_long = input_var[4]
    column = input_var[5]
    df = input_var[6]
    df_RSI_EMA = pd.DataFrame()
    columns = ['Date', 'Open', 'High', 'Low', 'Close', 'Adj Close', 'Volume']
    for col in columns:
        df_RSI_EMA[col] = df[col]

    #Calculating RSI.
    df_RSI_EMA = RSI(period_rsi, column, df_RSI_EMA)

    #Calculating Short and Long Term EMA.
    df_RSI_EMA['EMA{}'.format(period_short)] = EMA(df_RSI_EMA, period_short, column)
    df_RSI_EMA['EMA{}'.format(period_long)] = EMA(df_RSI_EMA, period_long, column)

    x = buy_sell_rsi_ema(df_RSI_EMA, period_short, period_long, rsi_low, rsi_high, column)
    df_RSI_EMA['Buy_Signal_Price'] = x[0]
    df_RSI_EMA['Sell_Signal_Price'] = x[1]

    strategy_hlinks = plot_rsi_ema(stock, df_RSI_EMA, plot_full_location, cur_descp, period_rsi,
                                   rsi_low, rsi_high, period_short, period_long, column)
    stock_status = stock_current_status(df_RSI_EMA)
    bt_data = BackTest('RSI_EMA', stock, df_RSI_EMA, plot_full_location, wb, cur_descp)
    print_report(stock, cur_descp, df_RSI_EMA, bt_data, stock_summary,
                 stock_status, stkplt_hlink, strategy_hlinks)

    return stock_status
