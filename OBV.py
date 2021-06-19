import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
plt.style.use('seaborn')

from Status import *
from BackTest import BackTest


def plot_obv(stock, plt_df, plot_full_location, cur_descp, OBV_EMA):
    plt_df = plt_df.set_index(pd.DatetimeIndex(plt_df['Date'].values))

    #Plotting the OBV and OBV EMA lines.
    plt.figure(figsize=(12.2, 4.5))
    plt.plot(plt_df['OBV'], label='OBV', color='orange')
    plt.plot(plt_df[OBV_EMA], label=OBV_EMA, color='purple')
    plt.title('{} OBV/{} Chart'.format(stock, OBV_EMA))
    plt.xlabel('{} to {}'.format(plt_df['Date'][0], plt_df['Date'].iloc[-1]))
    plt.ylabel('Close Price in INR')
    plot_name = '{}_OBV.png'.format(stock)
    plot_file = plot_full_location + '/' + stock + '/' + plot_name
    plt.savefig(plot_file, dpi=150)
    obv_hlink = '=HYPERLINK("{}","OBV Plot")'.format(plot_file)
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
    plot_name = '{}_OBV_Signals.png'.format(stock)
    plot_file = plot_full_location + '/' + stock + '/' + plot_name
    plt.savefig(plot_file, dpi=150)
    signal_hlink = '=HYPERLINK("{}","OBV Signal Plot")'.format(plot_file)
    plt.close('all')

    return (obv_hlink, signal_hlink)


def buy_sell_OBV(signal, col1, col2):
  sigPriceBuy = []
  sigPriceSell = []
  flag = -1
  #Loop through the length of the data.
  for i in range(0, len(signal)):
    #If OBV > OBV_EMA Then Buy. --> col1 => 'OBV' and col2 => 'OBV_EMA'
    if signal[col1][i] > signal[col2][i] and flag != 1:
      sigPriceBuy.append(signal['Close'][i])
      sigPriceSell.append(np.nan)
      flag = 1
    #If OBV > OBV_EMA Then Sell
    elif signal[col1][i] < signal[col2][i] and flag != 0:
      sigPriceSell.append(signal['Close'][i])
      sigPriceBuy.append(np.nan)
      flag = 0
    else:
      sigPriceBuy.append(np.nan)
      sigPriceSell.append(np.nan)
  return (sigPriceBuy, sigPriceSell)


def OBV(stock, input_var, plot_full_location, wb, cur_descp, stock_summary, stkplt_hlink):
    period_OBV_EMA = input_var[0]
    df_OBV = input_var[1]

    #Calculate the On Balance Volume(OBV)
    OBV = []
    OBV.append(0)

    #Loop through the data set (Close Price) from the second row (index 1) to the end of the data set.
    for i in range(1, len(df_OBV.Close)):
      if df_OBV.Close[i] > df_OBV.Close[i-1]:
        OBV.append(OBV[-1] + df_OBV.Volume[i])
      elif df_OBV.Close[i] < df_OBV.Close[i-1]:
        OBV.append(OBV[-1] - df_OBV.Volume[i])
      else:
        OBV.append(OBV[-1])

    #Store the OBV and OBV Exponential Moving Average (EMA) into new columns.
    OBV_EMA = 'OBV_EMA({})'.format(period_OBV_EMA)
    df_OBV['OBV'] = OBV
    df_OBV[OBV_EMA] = df_OBV['OBV'].ewm(span=period_OBV_EMA).mean()

    #Create buy and sell columns.
    x = buy_sell_OBV(df_OBV, 'OBV', OBV_EMA)
    df_OBV['Buy_Signal_Price'] = x[0]
    df_OBV['Sell_Signal_Price'] = x[1]

    strategy_hlinks = plot_obv(stock, df_OBV, plot_full_location, cur_descp, OBV_EMA)
    stock_status = stock_current_status(df_OBV)
    bt_data = BackTest('OBV', stock, df_OBV, plot_full_location, wb, cur_descp)
    print_report(stock, cur_descp, df_OBV, bt_data, stock_summary,
                 stock_status, stkplt_hlink, strategy_hlinks)

    return stock_status
