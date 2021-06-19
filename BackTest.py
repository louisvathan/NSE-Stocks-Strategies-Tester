import pandas as pd
import math
import numpy as np

from BacktestPlot import btPlot
from Excel_Processing import excel_conv

import logging
logger = logging.getLogger()
logger.setLevel(logging.CRITICAL)

def cnc_transaction_charges(transaction_details):
    buy_price = transaction_details[0]
    sell_price = transaction_details[1]
    qty = transaction_details[2]
    buy_transaction = buy_price * qty
    sell_transaction = sell_price * qty
    turnover = buy_transaction + sell_transaction
    brokerage = 0
    stt = turnover * (0.1 / 100)
    transaction_charges = turnover * (0.00325 / 100)
    gst = (brokerage + transaction_charges) * 0.18
    sebi_charges = turnover * (10 / 100000000)
    stamp_duty = buy_transaction * (0.015/100)

    total_charges = brokerage + stt + transaction_charges + gst + sebi_charges + stamp_duty

    dp_charges = 15.93

    total_charges = total_charges+dp_charges

    return total_charges

def portfolio_calc(cash, value, broker_charges):
    return cash+value-broker_charges

def p_and_l(transaction_details, total_charges):
    buy_price = transaction_details[0]
    sell_price = transaction_details[1]
    qty = transaction_details[2]
    buyvalue = buy_price * qty
    sellvalue = sell_price * qty
    price_diff = sellvalue + total_charges - buyvalue

    return price_diff


def broker_trade(df, cash_in_hand, stake, profit_loss, portfolio_values):
    broker_df = pd.DataFrame(columns=['Date', 'Open', 'Close'])
    broker_df['Date'] = df['Date']
    broker_df['Open'] = df['Open']
    broker_df['Close'] = df['Close']
    broker_df['Buy_Signal_Price'] = df['Buy_Signal_Price']
    broker_df['Sell_Signal_Price'] = df['Sell_Signal_Price']

    transaction_values = []
    transaction_details = []
    good_to_buy = True
    good_to_sell = False
    for i in range(len(df)):
        if i == 0:
            broker_df.loc[i, 'Buy_Sell_Price'] = np.nan
            broker_df.loc[i, 'Cash_In_Hand'] = cash_in_hand
            broker_df.loc[i, 'Value'] = 0
            broker_df.loc[i, 'Status'] = 'None'
            broker_df.loc[i, 'Transaction_Charges'] = 0
            broker_df.loc[i, 'Holding'] = 0
            profit_loss.append(0)
        elif i == 1:
            broker_df.loc[i, 'Buy_Sell_Price'] = np.nan
            broker_df.loc[i, 'Cash_In_Hand'] = cash_in_hand
            broker_df.loc[i, 'Value'] = 0
            broker_df.loc[i, 'Transaction_Charges'] = 0
            broker_df.loc[i, 'Status'] = 'None'
            broker_df.loc[i, 'Holding'] = 0
            profit_loss.append(0)
        else:
            if not math.isnan(float(df.loc[i-1, 'Buy_Signal_Price'])) and broker_df.loc[i-1, 'Holding'] == 0:
                if math.isnan(float(broker_df.loc[i, 'Status'])) or broker_df.loc[i, 'Status'] == 'Sell Completed':
                    broker_df.loc[i, 'Buy_Sell_Price'] = df.loc[i, 'Open']
                    amount_to_invest = broker_df.loc[i-1, 'Cash_In_Hand'] * stake
                    size = math.floor(amount_to_invest / broker_df.loc[i, 'Buy_Sell_Price'])
                    buy_transaction = size * broker_df.loc[i, 'Buy_Sell_Price']
                    broker_df.loc[i, 'Value'] = size * df.loc[i, 'Close']
                    broker_df.loc[i, 'Holding'] = size
                    broker_df.loc[i, 'Cash_In_Hand'] = broker_df.loc[i-1, 'Cash_In_Hand'] - buy_transaction
                    broker_df.loc[i, 'Transaction_Charges'] = 0
                    transaction_details.append(broker_df.loc[i, 'Buy_Sell_Price'])
                    broker_df.loc[i, 'Status'] = 'Order Placed'
                    profit_loss.append(0)
                    good_to_buy = False

            elif not math.isnan(float(df['Buy_Signal_Price'][i-2])) and broker_df.loc[i-2, 'Holding'] == 0:
                if broker_df.loc[i-1, 'Status'] == 'Order Placed':
                    broker_df.loc[i, 'Status'] = 'Order Completed'
                    broker_df.loc[i, 'Value'] = size * df.loc[i, 'Close']
                    broker_df.loc[i, 'Holding'] = size
                    broker_df.loc[i, 'Cash_In_Hand'] = broker_df.loc[i-1, 'Cash_In_Hand']
                    broker_df.loc[i, 'Buy_Sell_Price'] = np.nan
                    broker_df.loc[i, 'Transaction_Charges'] = 0
                    profit_loss.append(0)
                    good_to_sell = True

            elif not math.isnan(float(df['Sell_Signal_Price'][i-1])):
                if broker_df.loc[i-1, 'Status'] == 'Order Completed':
                    broker_df.loc[i, 'Buy_Sell_Price'] = df.loc[i, 'Open']
                    sell_transaction = size * broker_df.loc[i, 'Buy_Sell_Price']
                    broker_df.loc[i, 'Value'] = 0
                    broker_df.loc[i, 'Holding'] = 0
                    transaction_details.append(broker_df.loc[i, 'Buy_Sell_Price'])
                    transaction_details.append(size)
                    total_charges = cnc_transaction_charges(transaction_details)
                    broker_df.loc[i, 'Transaction_Charges'] = total_charges

                    trade_pl = p_and_l(transaction_details, total_charges)
                    profit_loss.append(trade_pl)

                    broker_df.loc[i, 'Cash_In_Hand'] = broker_df.loc[i-1, 'Cash_In_Hand'] + sell_transaction - total_charges
                    broker_df.loc[i, 'Status'] = 'Sell Completed'
                    transaction_details.clear()
                    good_to_buy = True
                    good_to_sell = False


                elif broker_df.loc[i-1, 'Status'] == 'None':
                    broker_df.loc[i, 'Buy_Sell_Price'] = np.nan
                    broker_df.loc[i, 'Cash_In_Hand'] = broker_df.loc[i-1, 'Cash_In_Hand']
                    broker_df.loc[i, 'Value'] = 0
                    broker_df.loc[i, 'Holding'] = broker_df.loc[i-1, 'Holding']
                    broker_df.loc[i, 'Status'] = 'None'
                    broker_df.loc[i, 'Transaction_Charges'] = 0
                    profit_loss.append(0)

            else:
                if broker_df.loc[i-1, 'Status'] == 'None' or broker_df.loc[i-1, 'Status'] == 'Sell Completed':
                    broker_df.loc[i, 'Buy_Sell_Price'] = np.nan
                    broker_df.loc[i, 'Cash_In_Hand'] = broker_df.loc[i-1, 'Cash_In_Hand']
                    broker_df.loc[i, 'Value'] = 0
                    broker_df.loc[i, 'Holding'] = broker_df.loc[i-1, 'Holding']
                    broker_df.loc[i, 'Status'] = 'None'
                    broker_df.loc[i, 'Transaction_Charges'] = 0
                    profit_loss.append(0)
                elif broker_df.loc[i-1, 'Status'] == 'Order Completed':
                    broker_df.loc[i, 'Buy_Sell_Price'] = np.nan
                    broker_df.loc[i, 'Cash_In_Hand'] = broker_df.loc[i-1, 'Cash_In_Hand']
                    broker_df.loc[i, 'Value'] = size * df.loc[i, 'Close']
                    broker_df.loc[i, 'Holding'] = broker_df.loc[i-1, 'Holding']
                    broker_df.loc[i, 'Status'] = broker_df.loc[i-1, 'Status']
                    broker_df.loc[i, 'Transaction_Charges'] = 0
                    profit_loss.append(0)

        portfolio_values.append(portfolio_calc(broker_df.loc[i, 'Cash_In_Hand'], broker_df.loc[i, 'Value'], broker_df.loc[i, 'Transaction_Charges']))

    return broker_df

def value_track(broker_df, portfolio_values):
    value_df = pd.DataFrame(columns=['Date'])
    value_df['Date'] = broker_df['Date']
    value_df['Cash'] = broker_df['Cash_In_Hand']
    value_df['Portfolio_Value'] = portfolio_values

    return value_df

def trade_result(broker_df, profit_loss):
    result_df = pd.DataFrame(columns=['Date'])
    result_df['Date'] = broker_df['Date']
    result_df['P&L'] = profit_loss
    for i in range(len(profit_loss)):
        if profit_loss[i] > 0:
            result_df.loc[i, 'Profit'] = profit_loss[i]
            result_df.loc[i, 'Loss'] = np.nan
        elif profit_loss[i] < 0:
            result_df.loc[i, 'Profit'] = np.nan
            result_df.loc[i, 'Loss'] = profit_loss[i]
        else:
            result_df.loc[i, 'Profit'] = np.nan
            result_df.loc[i, 'Loss'] = np.nan

    return result_df

def BackTest(indicator, stock, df, plot_full_location, wb, cur_descp):
    cash_in_hand = 10000
    stake = 0.95

    profit_loss = []
    portfolio_values = []
    broker_df = broker_trade(df, cash_in_hand, stake, profit_loss, portfolio_values)

    value_df = value_track(broker_df, portfolio_values)

    result_df = trade_result(broker_df, profit_loss)

    bt_hlink = btPlot(indicator, stock, df, value_df, result_df, plot_full_location, cur_descp)

    excel_conv(indicator, df, broker_df, value_df, result_df, wb)

    return (result_df, cash_in_hand, stake, portfolio_values, bt_hlink)
