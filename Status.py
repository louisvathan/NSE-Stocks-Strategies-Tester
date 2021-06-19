import pandas as pd

def percent_trades_profit(result_df):
    total_trades = 0
    profit_trades = 0
    for i in range(len(result_df)):
        if result_df.loc[i, 'P&L'] > 0: #Profit
            total_trades += 1
            profit_trades += 1
        elif result_df.loc[i, 'P&L'] < 0:  #Loss
            total_trades += 1
        else:  #No Trade.
            pass
    if total_trades > 0:
        profitable_trades = (profit_trades/total_trades) * 100
    else:
        profitable_trades = 0

    return profitable_trades
    

def stock_current_status(df):
    stock_status = []
    buy_location = 0
    sell_location = 0
    for i in reversed(range(len(df))):
        if buy_location == 0 & sell_location == 0:
            if float(df.loc[i, 'Adj Close']) - float(df.loc[i, 'Buy_Signal_Price']) == 0 or float(df.loc[i, 'Close']) - float(df.loc[i, 'Buy_Signal_Price']) == 0:
                buy_location = i
            if float(df.loc[i, 'Adj Close']) - float(df.loc[i, 'Sell_Signal_Price']) == 0 or float(df.loc[i, 'Close']) - float(df.loc[i, 'Sell_Signal_Price']) == 0:
                sell_location = i
        else:
            break

    if buy_location > sell_location:
        #1.SIGNAL TYPE.
        actual_signal = 'Buy Signal'
        stock_status.append('Buy')
        #2.SIGNAL PRICE.
        stock_status.append(float(df.loc[buy_location, 'Buy_Signal_Price']))
        diff = float(df['Adj Close'].iloc[-1]) - float(df.loc[buy_location, 'Buy_Signal_Price'])
        days = len(df) - buy_location
    if sell_location > buy_location:
        #1.SIGNAL TYPE.
        actual_signal = 'Sell Signal'
        stock_status.append('Sell')
        #2.SIGNAL PRICE.
        stock_status.append(float(df.loc[sell_location, 'Sell_Signal_Price']))
        diff = float(df['Adj Close'].iloc[-1]) - float(df.loc[sell_location, 'Sell_Signal_Price'])
        days = len(df) - sell_location
    if sell_location == buy_location:
        actual_signal = 'No Signal'
        stock_status.append('NIL')
        stock_status.append(float(0))
        diff = 0
        days = 'N/A'

    #3.CLOSE PRICE
    stock_status.append(float(df['Adj Close'].iloc[-1]))

    #4.PERCENTAGE DIFFERENCE.
    percent_decimals = (diff / float(df['Adj Close'].iloc[-1]))
    percent = percent_decimals * 100
    stock_status.append(float(percent_decimals))

    #5.REMARKS.
    remarks = '{}, {} days before.'.format(actual_signal, days)
    stock_status.append(remarks)

    return stock_status


def print_report(stock, col, df, bt_data, stock_summary, stock_status, stkplt_hlink, strategy_hlinks):
    result_df, cash_in_hand = bt_data[0], bt_data[1]
    stake, portfolio_values = bt_data[2], bt_data[3]
    bt_hlink = bt_data[4]
    stock_summary.loc['Current Stock Details', col] = ''
    stock_summary.loc['Last Signal', col] = stock_status[0]
    stock_summary.loc['Last Signal Price(Rs.)', col] = stock_status[1]
    stock_summary.loc['Last Close Price(Rs.)', col] = stock_status[2]
    stock_summary.loc['Price Difference(%)', col] = stock_status[3]
    stock_summary.loc['Remarks', col] = stock_status[4]
    stock_summary.loc['Stock Plot', col] = stkplt_hlink
    stock_summary.loc['Strategy Plot', col] = strategy_hlinks[0]
    stock_summary.loc['Signals Plot', col] = strategy_hlinks[1]
    stock_summary.loc['Backtest Results', col] = ''
    stock_summary.loc['Starting Date', col] = str(df.loc[0, 'Date'])
    stock_summary.loc['Ending Date', col] = str(df.loc[len(df)-1, 'Date'])
    stock_summary.loc['Invested Money(Rs.)', col] = cash_in_hand
    stock_summary.loc['Percent used to trade(%)', col] = stake
    stock_summary.loc['Final Portfolio Value(Rs.)', col] = portfolio_values[-1]
    stock_summary.loc['Profit & Loss(%)', col] = ((portfolio_values[-1]-cash_in_hand)/cash_in_hand)
    stock_summary.loc['Positive Trade Percentage(%)', col] = percent_trades_profit(result_df)/100
    stock_summary.loc['Backtest Plot', col] = bt_hlink
    '''print('-----------BackTesting Results-----------')
    print('Strategy                             : {}'.format(strategy))
    print('Stock                                : {}'.format(stock))
    print('Starting Date                        : {}'.format(str(df.loc[0, 'Date'])))
    print('Ending Date                          : {}'.format(str(df.loc[len(df)-1, 'Date'])))
    print('Invested Money                       : Rs.{:.2f}'.format(cash_in_hand))
    print('Percent used to trade                : {:.2f}%'.format(stake * 100))
    print('Final Portfolio Value                : Rs.{:.2f}'.format(portfolio_values[-1]))
    print('Profit & Loss                        : {:.2f}%'.format(((portfolio_values[-1]-cash_in_hand)/cash_in_hand)*100))
    print('Percentage of Trades ended in Profit : {:.2f}%'.format(percent_trades_profit(result_df)))'''
