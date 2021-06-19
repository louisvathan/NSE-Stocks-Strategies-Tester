import pandas as pd
import math
import matplotlib as mpl
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
from matplotlib.widgets import Cursor, MultiCursor
import numpy as np

from Ploting import reformat_large_tick_values

plt.style.use('seaborn')

def btPlot(indicator, stock, df, value_df, result_df, plot_full_location, cur_descp):
    df = df.set_index(pd.DatetimeIndex(df['Date'].values))
##    pd.set_option('display.max_rows', None)
##    pd.set_option('display.max_columns', None)
    value_df = value_df.set_index(pd.DatetimeIndex(value_df['Date'].values))
    result_df = result_df.set_index(pd.DatetimeIndex(result_df['Date'].values))

    ax = ('ax1', 'ax2', 'ax3', 'ax4')
    fig, ax = plt.subplots(nrows=4, ncols=1, sharex=True, gridspec_kw={'height_ratios': [1, 1, 3, 1]})
    
    ax[0].plot(value_df.index, value_df['Cash'], label='Cash', color='red')
    ax[0].plot(value_df.index, value_df['Portfolio_Value'], label='Total Portfolio Value', color='blue')
    ax[0].set_title('{} - {} Backtest Results'.format(stock, cur_descp))
    ax[0].legend(title="Broker Account", fontsize='small', fancybox=True)
    ax[0].ticklabel_format(useOffset=False, style='plain', axis='y')
    
    ##ax[1].scatter(x=result_df.index, y=result_df['Profit'], marker='o', label='Profit', color='green')
    if not result_df['Profit'].isnull().all():
        ax[1].scatter(x=result_df.index, y=result_df['Profit'], marker='o', color='green')
    if not result_df['Loss'].isnull().all():
        ax[1].scatter(x=result_df.index, y=result_df['Loss'], marker='o', color='red')
    ax[1].axhline(0, linestyle='--', alpha=0.5, color='yellow')
    ax[1].legend(title="Trades - Net Profit/Loss", fontsize='small', fancybox=True)
    ax[1].ticklabel_format(useOffset=False, style='plain', axis='y')

    ax[2].plot(df.index, df['Close'], label='Close Price')
    ax[2].plot(df.index, df['Buy_Signal_Price'], color='green', marker='^', alpha=1, label='Buy')
    ax[2].plot(df.index, df['Sell_Signal_Price'], color='red', marker='v', alpha=1, label='Sell')
    ax[2].legend(title="{} - {} (1 Day)".format(stock, indicator), fontsize='small', fancybox=True)
    ax[2].ticklabel_format(useOffset=False, style='plain', axis='y')

    ax[3].bar(df.index, df['Volume'], label='Trade Volume')
    ax[3].set_xlabel('Dates', labelpad=16)
    ax[3].legend(title="Volume", fontsize='small', fancybox=True)
    ax[3].yaxis.set_major_formatter(ticker.FuncFormatter(reformat_large_tick_values))

    for a in ax:
        a.patch.set_edgecolor('black')  
        a.patch.set_linewidth('1')
        a.yaxis.set_label_position("right")
        a.yaxis.tick_right()

    plt.subplots_adjust(wspace=0, hspace=0)
    btplot_name = '{}-{}-Backtest_Plot.png'.format(stock, indicator)
    btplot_file = plot_full_location + '/' + stock + '/' + btplot_name
    bt_hlink = '=HYPERLINK("{}","Backtest Report")'.format(btplot_file)
##    plt.show()
    plt.savefig(btplot_file, dpi=150)
    plt.close('all')

    return bt_hlink
