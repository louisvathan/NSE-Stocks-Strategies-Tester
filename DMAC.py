import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
plt.style.use('seaborn')

from Status import *
from BackTest import BackTest


def plot_dmac(stock, plt_df, plot_full_location, cur_descp):
    words = cur_descp.split(" ", 1)
    values = words[1].replace('(', '').replace(')', '')
    variables = tuple(map(int, values.split(', ')))
    dmac_short = variables[0]
    dmac_long = variables[1]
    strategy = words[0]
    
    plt_df = plt_df.set_index(pd.DatetimeIndex(plt_df['Date'].values))

    #Plotting DMAC lines.
    plt.figure(figsize= (12.5, 4.5))
    plt.plot(plt_df['Adj Close'], label=stock)
    plt.plot(plt_df['SMA{}'.format(dmac_short)], label='SMA{}'.format(dmac_short))
    plt.plot(plt_df['SMA{}'.format(dmac_long)], label='SMA{}'.format(dmac_long))
    plt.title('{} Adj. Close Price History'.format(stock))
    plt.xlabel('{} to {}'.format(plt_df['Date'][0], plt_df['Date'].iloc[-1]))
    plt.ylabel('Adj Close Price INR')
    plt.legend(loc='upper left')
    plot_name = '{}_{}.png'.format(stock, strategy)
    plot_file = plot_full_location + '/' + stock + '/' + plot_name
    plt.savefig(plot_file, dpi=150)
    dmac_hlink = '=HYPERLINK("{}","{} Plot")'.format(plot_file, strategy)
    plt.close('all')

    #Plotting Buy and Sell Signals
    plt.figure(figsize=(12.6, 4.6))
    plt.plot(plt_df['Close'], label=stock, alpha=0.35)
    plt.plot(plt_df['SMA{}'.format(dmac_short)], label='SMA{}'.format(dmac_short), alpha=0.35)
    plt.plot(plt_df['SMA{}'.format(dmac_long)], label='SMA{}'.format(dmac_long), alpha=0.35)
    if not plt_df['Buy_Signal_Price'].isnull().all():
        plt.scatter(plt_df.index, plt_df['Buy_Signal_Price'], label='Buy', marker='^', color='green')
    if not plt_df['Sell_Signal_Price'].isnull().all():
        plt.scatter(plt_df.index, plt_df['Sell_Signal_Price'], label='Sell', marker='v', color='red')
    plt.title('Close Price History Buy & Sell Signals ({} - {})'.format(stock, cur_descp))
    plt.xlabel('{} to {}'.format(plt_df['Date'][0], plt_df['Date'].iloc[-1]))
    plt.ylabel('Close Price INR')
    plt.legend(loc='upper left')
    plot_name = '{}_{}_Signals.png'.format(stock, strategy)
    plot_file = plot_full_location + '/' + stock + '/' + plot_name
    plt.savefig(plot_file, dpi=150)
    signal_hlink = '=HYPERLINK("{}","{} Signal Plot")'.format(plot_file, strategy)
    plt.close('all')

    return (dmac_hlink, signal_hlink)
    

def buy_sell_dmac(data, dmac_short, dmac_long):
    sigPriceBuy = []
    sigPriceSell = []
    flag = -1

    for i in range(len(data)):
        if data['SMA{}'.format(dmac_short)][i] > data['SMA{}'.format(dmac_long)][i]:
            if flag != 1:
                sigPriceBuy.append(data['Close'][i])
                sigPriceSell.append(np.nan)
                flag = 1
            else:
                sigPriceBuy.append(np.nan)
                sigPriceSell.append(np.nan)
        elif data['SMA{}'.format(dmac_short)][i] < data['SMA{}'.format(dmac_long)][i]:
            if flag != 0:
                sigPriceBuy.append(np.nan)
                sigPriceSell.append(data['Close'][i])
                flag = 0
            else:
                sigPriceBuy.append(np.nan)
                sigPriceSell.append(np.nan)
        else:
            sigPriceBuy.append(np.nan)
            sigPriceSell.append(np.nan)
    return (sigPriceBuy, sigPriceSell)


def DMAC(stock, input_var, plot_full_location, wb, cur_descp, stock_summary, stkplt_hlink):
    dmac_short = input_var[0]
    dmac_long = input_var[1]
    df_DMAC = input_var[2]

    sma_short = pd.DataFrame()
    sma_short['Adj Close Price'] = df_DMAC['Adj Close'].rolling(window=dmac_short).mean()
    
    sma_long = pd.DataFrame()
    sma_long['Adj Close Price'] = df_DMAC['Adj Close'].rolling(window=dmac_long).mean()
    
    data = pd.DataFrame()
    data['Close'] = df_DMAC['Adj Close']
    data['SMA{}'.format(dmac_short)] = sma_short['Adj Close Price']
    data['SMA{}'.format(dmac_long)] = sma_long['Adj Close Price']
    
    a = buy_sell_dmac(data, dmac_short, dmac_long)
    data['Buy_Signal_Price'] = a[0]
    data['Sell_Signal_Price'] = a[1]

    df_DMAC['SMA{}'.format(dmac_short)] = data['SMA{}'.format(dmac_short)]
    df_DMAC['SMA{}'.format(dmac_long)] = data['SMA{}'.format(dmac_long)]
    df_DMAC['Buy_Signal_Price'] = data['Buy_Signal_Price']
    df_DMAC['Sell_Signal_Price'] = data['Sell_Signal_Price']
    
    strategy_hlinks = plot_dmac(stock, df_DMAC, plot_full_location, cur_descp)
    stock_status = stock_current_status(df_DMAC)
    bt_data = BackTest('DMAC', stock, df_DMAC, plot_full_location, wb, cur_descp)
    print_report(stock, cur_descp, df_DMAC, bt_data, stock_summary,
                 stock_status, stkplt_hlink, strategy_hlinks)

    return stock_status
