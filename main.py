from sklearn.neighbors import KernelDensity
from pandas_datareader import data as web
from matplotlib import pyplot as plt
import xlwings as xw
import pandas as pd
import numpy as np


def get_params(book, sheet_name):
    """
    Get ticker, start_date and end_date parameters.
    :param xw.Book book: Excel work book.
    :return: ticker, start_date and end_date parameters.
    """
    sheet = book.sheets[sheet_name]
    ticker = sheet.range('C2').value
    start_date = sheet.range('C3').value
    end_date = sheet.range('C4').value

    return ticker, start_date, end_date


def get_adj_closes(ticker, start_date=None, end_date=None):
    """
    Obtains adjusted close prices from a given ticker within start_date and end_date.
    :param str ticker: Ticker symbol of the company.
    :param str start_date: Dates lower bound.
    :param str end_date: Dates upper bound.
    :return: pd.DataFrame of adjusted close prices.
    """
    closes = web.DataReader(name=ticker,
                            data_source='yahoo',
                            start=start_date,
                            end=end_date)['Adj Close']

    return pd.DataFrame(closes)


def write_closes(closes, ticker, book, sheet_name):
    """
    Write ticker closes data into sheet_name of book.
    :param pd.DataFrame closes: Adjusted close prices.
    :param str ticker: Ticker of the company.
    :param xl.Book book: Excel workbook.
    :param str sheet_name: Name of the sheet.
    :return: None.
    """
    book.sheets.add(name=sheet_name)
    sheet2 = book.sheets['prices']
    sheet2.range('B2').value = ticker
    sheet2.range('B3').value = closes


def montecarlo():
    """
    Run montecarlo simulation. Reads parameters from the excel sheet Params, and perform price simulations
    over one year. Finally, plots the simulation results and adds it to the sheet.
    :return: None
    """
    # Get the Excel work book
    wb = xw.Book.caller()

    # Get params
    ticker, start_date, end_date = get_params(book=wb, sheet_name="Params")

    # Get adj closes
    closes = get_adj_closes(ticker, start_date, end_date)

    # Calculate simple daily returns
    ret = closes.pct_change().dropna()
    # Estimate density with Gaussian kernels
    kde = KernelDensity(kernel='gaussian', bandwidth=0.001).fit(ret)
    # Returns simulation
    n_days, n_sim = 252, 100000
    d_range = pd.date_range(start=closes.index[-1] + pd.Timedelta(days=1), periods=n_days)
    ret_sim = pd.DataFrame(data=kde.sample(n_samples=n_days * n_sim).reshape((n_days, n_sim)),
                           index=d_range)
    # To prices
    closes_sim = (closes.iloc[-1].values[0]) * (1 + ret_sim).cumprod()

    # Get 5% - 95% percentile bands
    band_5 = pd.DataFrame(data={'5% band': np.percentile(closes_sim, 5, axis=1)}, index=d_range)
    band_95 = pd.DataFrame(data={'95% band': np.percentile(closes_sim, 95, axis=1)}, index=d_range)

    # Plot past prices, bands and prices scenarios
    fig = plt.figure(figsize=(6, 4))
    plt.plot(closes.iloc[-100:], label='Historical Adj Close')
    plt.plot(band_5, label='5% Percentile Band')
    plt.plot(band_95, label='95% Percentile Band')
    plt.plot(closes_sim.sample(10, axis=1), label='Price Scenarios')
    plt.xlabel('Time')
    plt.ylabel('Price')
    plt.legend(loc='upper left', bbox_to_anchor=(1.05, 1))

    # Add the plot to the sheet
    sheet = wb.sheets['Params']
    sheet.pictures.add(fig, name='Montecarlo Simulation', update=True,
                       left=sheet.range('B9').left, top=sheet.range('B9').top)
