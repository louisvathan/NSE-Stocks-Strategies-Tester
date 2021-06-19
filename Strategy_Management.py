import pandas as pd
import os
import ast

#Importing Functions from other files.
#from RSI import *
from MACD import MACD
from DMAC import DMAC
from DEMA import DEMA
from OBV import OBV
from RSI_EMA import RSI_EMA
from Supertrend import Supertrend
from BackTest import *
from Excel_Processing import stock_summary_sheet
from Ploting import stock_plot

def Strategy_Management(stock, df_st, df_lt, plot_full_location, wb, summary):

    strategies = ['MACD', 'DMAC', 'DEMA', 'OBV', 'RSI_EMA', 'SuperTrend']
    dataset = ['short', 'long', 'short', 'short', 'long', 'long']
    
    try:
        os.mkdir('{}/{}'.format(plot_full_location, stock))
    except:
        pass
    
    #Creating Stock Plots.
    stkplt_hlink_short = stock_plot(df_st, plot_full_location, stock)
    stkplt_hlink_long = stock_plot(df_lt, plot_full_location, stock)
    
    #Variables for Moving Average Convergence Divergence (MACD) Indicator.
    macd_short_period = 12
    macd_long_period = 26
    macd_signal_period = 9
    #global MACD_var
    

    #Variables for Dual Moving Average Crossover (DMAC) Indicator.
    dmac_short = 30
    dmac_long = 100
    #global DMAC_var

    #Variables for Double Exponential Moving Average (DEMA) Indicator.
    period_short_DEMA = 20
    period_long_DEMA = 50
    column = 'Close'
    #global DEMA_var

    #Variables for On-Balance Volume (OBV) Indicator.
    period_OBV_EMA = 20

    #Variables for RSI_EMA Indicator.
    period_rsi = 14
    rsi_low = 40
    rsi_high = 60
    period_short = 5
    period_long = 10
    column = 'Close'

    #Variables for SuperTrend.
    ATR_Length_ST = 10
    multiplier = 3
    
    #Creating summary dataframe.
    index = ['Current Stock Details', 'Last Signal', 'Last Signal Price(Rs.)', 'Last Close Price(Rs.)', 'Price Difference(%)', 'Remarks',
             'Stock Plot', 'Strategy Plot', 'Signals Plot',
             'Backtest Results', 'Starting Date', 'Ending Date', 'Invested Money(Rs.)', 'Percent used to trade(%)',
             'Final Portfolio Value(Rs.)', 'Profit & Loss(%)', 'Positive Trade Percentage(%)',
             'Backtest Plot']
    startegy_descp = [
                        '{} ({}, {}, {})'.format(strategies[0], macd_short_period, macd_long_period, macd_signal_period),
                        '{} ({}, {})'.format(strategies[1], dmac_short, dmac_long),
                        '{} ({}, {}, {})'.format(strategies[2], period_short_DEMA, period_long_DEMA, column),
                        '{} ({})'.format(strategies[3], period_OBV_EMA),
                        '{} ({}, {}, {}) ({}, {}, {})'.format(strategies[4], period_rsi, rsi_low, rsi_high, period_short, period_long, column),
                        '{} ({}, {})'.format(strategies[5], ATR_Length_ST, multiplier)
                    ]
    stock_summary = pd.DataFrame(index=index, columns=startegy_descp)
    
    
    #Running Startegies.
    #Moving Average Convergence Divergence (MACD)
    MACD_var = (macd_short_period, macd_long_period, macd_signal_period, df_st)
    stock_status = MACD(stock, MACD_var, plot_full_location, wb, startegy_descp[strategies.index('MACD')], stock_summary, stkplt_hlink_short)
    summary.loc[stock, startegy_descp[strategies.index('MACD')]] = stock_status[4]
    
    #Dual Moving Average Crossover (DMAC)
    DMAC_var = (dmac_short, dmac_long, df_lt)
    stock_status = DMAC(stock, DMAC_var, plot_full_location, wb, startegy_descp[strategies.index('DMAC')], stock_summary, stkplt_hlink_long)
    summary.loc[stock, startegy_descp[strategies.index('DMAC')]] = stock_status[4]
    
    #Double Exponential Moving Average (DEMA)
    DEMA_var = (period_short_DEMA, period_long_DEMA, column, df_st)
    stock_status = DEMA(stock, DEMA_var, plot_full_location, wb, startegy_descp[strategies.index('DEMA')], stock_summary, stkplt_hlink_short)
    summary.loc[stock, startegy_descp[strategies.index('DEMA')]] = stock_status[4]

    #On-Balance Volume (OBV).
    OBV_var = (period_OBV_EMA, df_st)
    stock_status = OBV(stock, OBV_var, plot_full_location, wb, startegy_descp[strategies.index('OBV')], stock_summary, stkplt_hlink_short)
    summary.loc[stock, startegy_descp[strategies.index('OBV')]] = stock_status[4]

    #RSI and EMA.
    RSI_EMA_var = (period_rsi, rsi_low, rsi_high, period_short, period_long, column, df_lt)
    stock_status = RSI_EMA(stock, RSI_EMA_var, plot_full_location, wb, startegy_descp[strategies.index('RSI_EMA')], stock_summary, stkplt_hlink_short)
    summary.loc[stock, startegy_descp[strategies.index('RSI_EMA')]] = stock_status[4]

    #SuperTrend.
    ST_var = (ATR_Length_ST, multiplier, df_lt)
    stock_status = Supertrend(stock, ST_var, plot_full_location, wb, startegy_descp[strategies.index('SuperTrend')], stock_summary, stkplt_hlink_long)
    summary.loc[stock, startegy_descp[strategies.index('SuperTrend')]] = stock_status[4]
    
    last_strat = False
    for strategy in strategies:
        
        cur_descp = startegy_descp[strategies.index(strategy)]
        cur_data = dataset[strategies.index(strategy)]

        #stock_status = eval(strategy + '(stock, {}_var, plot_full_location, wb, {}, stock_summary, stkplt_hlink_{})'.format(strategy, cur_descp, cur_data))
        
        '''
##        dataset = 'df_{}'.format(strategy)
##        indi_name = '{}_var'.format(strategy)
##        if cur_data == 'short':
##            globals()[dataset] = df_st
##            globals()[dataset] = globals()[dataset]['Date', 'Open', 'High', 'Low', 'Close', 'Adj Close', 'Volume']
##            stkplt_hlink = stkplt_hlink_short
##        else:
##            globals()[dataset] = df_lt
##            globals()[dataset] = globals()[dataset]['Date', 'Open', 'High', 'Low', 'Close', 'Adj Close', 'Volume']
##            stkplt_hlink = stkplt_hlink_long
##        print(globals()[dataset])
##        
##        method_name = strategy
##        possibles = globals().copy()
##        possibles.update(locals())
##        method = possibles.get(method_name)
##        df = method(stock, globals()[indi_name], globals()[dataset])
##        #print(df)
        
        #df = eval(strategy + '(stock, {}_var)'.format(strategy))
        #stock_status = stock_current_status(df)
        stock_status = eval('stock_current_status' + '(df_{})'.format(strategy))
        #bt_data = BackTest(strategy, stock, df, plot_full_location, wb, cur_descp)
        bt_data = eval('BackTest' + '(strategy, stock, df_{}, plot_full_location, wb, cur_descp)'.format(strategy))
        #print_report(stock, cur_descp, df, bt_data, stock_summary, stock_status, stkplt_hlink)
        eval('print_report' + '(stock, cur_descp, df_{}, bt_data, stock_summary, stock_status, stkplt_hlink)'.format(strategy))'''

        if cur_data == 'short':
            stkplt_hlink = stkplt_hlink_short
        else:
            stkplt_hlink = stkplt_hlink_long
            
        #Adding Remarks to Summary DataFrame.
        #summary.loc[stock, cur_descp] = stock_status[4]

        if strategies.index(strategy)+1 == len(strategies):
            print(stock_summary)
            stock_summary_sheet(stock, wb, stock_summary, index, startegy_descp, stkplt_hlink)

    return summary
