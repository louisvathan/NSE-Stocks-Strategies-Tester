import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
plt.style.use('seaborn')

from Status import *
from BackTest import BackTest


def plot_st(stock, plt_df, plot_full_location, cur_descp, ATR_Length, multiplier, df_stsig):
    plt_df = plt_df.set_index(pd.DatetimeIndex(plt_df['Date'].values))

    #Plotting Supertrend with close price.
    plt.figure(figsize=(12.2, 4.5))
    plt.plot(plt_df['Close'], label='Close')
    plt.plot(plt_df['Supertrend({}, {})'.format(ATR_Length, multiplier)], label='Supertrend({}, {})'.format(ATR_Length, multiplier))
    plt.scatter(plt_df.index, df_stsig['Buy_Signal'], label='Buy_Signal', marker='^', alpha=1, color='green')
    plt.scatter(plt_df.index, df_stsig['Sell_Signal'], label='Sell_Signal', marker='v', alpha=1, color='red')
    plt.title('{} SuperTrend Chart'.format(stock))
    plt.xlabel('{} to {}'.format(plt_df['Date'][0], plt_df['Date'].iloc[-1]))
    plt.ylabel('Close Price in INR')
    plot_name = '{}_SuperTrend.png'.format(stock)
    plot_file = plot_full_location + '/' + stock + '/' + plot_name
    plt.savefig(plot_file, dpi=150)
    st_hlink = '=HYPERLINK("{}","SuperTrend Plot")'.format(plot_file)
    plt.close('all')

    #Plot the Buy and Sell Prices.
    plt.figure(figsize=(12.2, 4.5))
    plt.plot(plt_df['Close'], label='Close', alpha=0.35)
    plt.scatter(plt_df.index, plt_df['Buy_Signal_Price'], label='Buy_Signal', marker='^', alpha=1, color='green')
    plt.scatter(plt_df.index, plt_df['Sell_Signal_Price'], label='Sell_Signal', marker='v', alpha=1, color='red')
    plt.title('Close Price Buy & Sell Signals ({} - {})'.format(stock, cur_descp))
    plt.xlabel('{} to {}'.format(plt_df['Date'][0], plt_df['Date'].iloc[-1]))
    plt.ylabel('Close Price in INR')
    plt.legend(loc='upper left')
    plot_name = '{}_SuperTrend_Signals.png'.format(stock)
    plot_file = plot_full_location + '/' + stock + '/' + plot_name
    plt.savefig(plot_file, dpi=150)
    signal_hlink = '=HYPERLINK("{}","SuperTrend Signal Plot")'.format(plot_file)
    plt.close('all')

    return (st_hlink, signal_hlink)


def buy_sell_st(data, ATR_Length, multiplier):
    sigPriceBuy = []
    sigPriceSell = []
    sigBuy = []
    sigSell = []
    flag = -1

    for i in range(len(data)):
        if data['Supertrend({}, {})'.format(ATR_Length, multiplier)][i] < data['Close'][i]:
            if flag != 1:
                sigPriceBuy.append(data['Close'][i])
                sigPriceSell.append(np.nan)
                sigBuy.append(data['Supertrend({}, {})'.format(ATR_Length, multiplier)][i])
                sigSell.append(np.nan)
                flag = 1
            else:
                sigPriceBuy.append(np.nan)
                sigPriceSell.append(np.nan)
                sigBuy.append(np.nan)
                sigSell.append(np.nan)
        elif data['Supertrend({}, {})'.format(ATR_Length, multiplier)][i] > data['Close'][i]:
            if flag != 0:
                sigPriceBuy.append(np.nan)
                sigPriceSell.append(data['Close'][i])
                sigBuy.append(np.nan)
                sigSell.append(data['Supertrend({}, {})'.format(ATR_Length, multiplier)][i])
                flag = 0
            else:
                sigPriceBuy.append(np.nan)
                sigPriceSell.append(np.nan)
                sigBuy.append(np.nan)
                sigSell.append(np.nan)
        else:
            sigPriceBuy.append(np.nan)
            sigPriceSell.append(np.nan)
            sigBuy.append(np.nan)
            sigSell.append(np.nan)
    return (sigPriceBuy, sigPriceSell, sigBuy, sigSell)


def Supertrend(stock, input_var, plot_full_location, wb, cur_descp, stock_summary, stkplt_hlink):
    ATR_Length = input_var[0]
    multiplier = input_var[1]
    df_ST = input_var[2]

    #Calculating SuperTrend.
    df_ST = df_ST.reindex(['Date', 'Open', 'High', 'Low', 'Close', 'Adj Close', 'Volume', 'TR', 'ATR',
                     'Upper Band Basic', 'Lower Band Basic', 'Upper Band', 'Lower Band', 'Supertrend({}, {})'.format(ATR_Length, multiplier)], axis=1)

    for i in range(len(df_ST)):
        if i == 0:
            df_ST.loc[i, 'TR'] = np.nan
        else:
            #Calculating True Range(TR).
            #TR = MAX ((High - Low), ABS(High - Close), ABS(Low - Close))
            HL = df_ST.loc[i, 'High'] - df_ST.loc[i, 'Low']
            HPC = abs(df_ST.loc[i, 'High'] - df_ST.loc[i-1, 'Close'])
            LPC = abs(df_ST.loc[i, 'Low'] - df_ST.loc[i-1, 'Close'])
            df_ST.loc[i, 'TR'] = max(HL, HPC, LPC)
            if i>ATR_Length-1:
                #Calculating Average True Range(ATR).
                total_tr = 0
                for j in range(0, ATR_Length):
                    total_tr = total_tr + df_ST.loc[i-j, 'TR']
                df_ST.loc[i, 'ATR'] = total_tr/ATR_Length
                df_ST.loc[i, 'Upper Band Basic'] = ((df_ST.loc[i, 'High']+df_ST.loc[i, 'Low'])/2)+(multiplier*df_ST.loc[i, 'ATR'])
                df_ST.loc[i, 'Lower Band Basic'] = ((df_ST.loc[i, 'High']+df_ST.loc[i, 'Low'])/2)-(multiplier*df_ST.loc[i, 'ATR'])
                df_ST.fillna(0, inplace=True)
            
                if df_ST.loc[i, 'Upper Band Basic'] < df_ST.loc[i-1, 'Upper Band'] or df_ST.loc[i-1, 'Close'] > df_ST.loc[i-1, 'Upper Band']:
                    df_ST.loc[i, 'Upper Band'] = float(df_ST.loc[i, 'Upper Band Basic'])
                else:
                    df_ST.loc[i, 'Upper Band'] = float(df_ST.loc[i-1, 'Upper Band'])
            
                if df_ST.loc[i, 'Lower Band Basic'] > df_ST.loc[i-1, 'Lower Band'] or df_ST.loc[i-1, 'Close'] < df_ST.loc[i-1, 'Lower Band']:
                    df_ST.loc[i, 'Lower Band'] = float(df_ST.loc[i, 'Lower Band Basic'])
                else:
                    df_ST.loc[i, 'Lower Band'] = float(df_ST.loc[i-1, 'Lower Band'])

                if df_ST.loc[i-1, 'Supertrend({}, {})'.format(ATR_Length, multiplier)]==df_ST.loc[i-1, 'Upper Band'] and df_ST.loc[i, 'Close']<df_ST.loc[i, 'Upper Band']:
                    df_ST.loc[i, 'Supertrend({}, {})'.format(ATR_Length, multiplier)] = df_ST.loc[i, 'Upper Band']
                elif df_ST.loc[i-1, 'Supertrend({}, {})'.format(ATR_Length, multiplier)]==df_ST.loc[i-1, 'Upper Band'] and df_ST.loc[i, 'Close']>df_ST.loc[i, 'Upper Band']:
                    df_ST.loc[i, 'Supertrend({}, {})'.format(ATR_Length, multiplier)] = df_ST.loc[i, 'Lower Band']
                elif df_ST.loc[i-1, 'Supertrend({}, {})'.format(ATR_Length, multiplier)]==df_ST.loc[i-1, 'Lower Band'] and df_ST.loc[i, 'Close']>df_ST.loc[i, 'Lower Band']:
                    df_ST.loc[i, 'Supertrend({}, {})'.format(ATR_Length, multiplier)] = df_ST.loc[i, 'Lower Band']
                elif df_ST.loc[i-1, 'Supertrend({}, {})'.format(ATR_Length, multiplier)]==df_ST.loc[i-1, 'Lower Band'] and df_ST.loc[i, 'Close']<df_ST.loc[i, 'Lower Band']:
                    df_ST.loc[i, 'Supertrend({}, {})'.format(ATR_Length, multiplier)] = df_ST.loc[i, 'Upper Band']
                else:
                    df_ST.loc[i, 'Supertrend({}, {})'.format(ATR_Length, multiplier)] = np.nan

    df_ST.replace(0, np.nan, inplace=True)

    a = buy_sell_st(df_ST, ATR_Length, multiplier)
    df_ST['Buy_Signal_Price'] = a[0]
    df_ST['Sell_Signal_Price'] = a[1]

    df_stsig = pd.DataFrame()
    df_stsig['Date'] = df_ST['Date']
    df_stsig['Buy_Signal'] = a[2]
    df_stsig['Sell_Signal'] = a[3]

    strategy_hlinks = plot_st(stock, df_ST, plot_full_location, cur_descp, ATR_Length, multiplier, df_stsig)
    stock_status = stock_current_status(df_ST)
    bt_data = BackTest('SuperTrend', stock, df_ST, plot_full_location, wb, cur_descp)
    print_report(stock, cur_descp, df_ST, bt_data, stock_summary,
                 stock_status, stkplt_hlink, strategy_hlinks)

    return stock_status
