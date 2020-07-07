# coding: utf-8
#

import argparse
import os

import matplotlib
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib.backends.backend_pdf import PdfPages

matplotlib.style.use('bmh')


def getQuarters(filename):
    xp = pd.ExcelFile(filename)
    nesco = xp.parse('Data Sheet', header=40)
    quaters = nesco.drop(nesco.index[range(9, 52)])

    quaters.set_index('Report Date', inplace=True)
    quaters.index.name = 'Quaters'
    return quaters

#quaters=getQuarters(fpath+fname)

def get_mcap(filename):
    xp = pd.ExcelFile(filename)
    data = xp.parse("Data Sheet")
    mcap = data.iloc[7,1] 
    return mcap

def get_pl_bal_cash_price(filename):
    xp = pd.ExcelFile(filename)
    nesco = xp.parse("Data Sheet", header=15)
    pl = nesco.iloc[:15, :].copy()
    bal = nesco.iloc[39:56, :].copy()
    cash = nesco.iloc[64:69, :].copy()
    price = nesco.iloc[73, :].copy()
    price = price.iloc[1:].copy()

    pl.set_index('Report Date', inplace=True)
    pl.index.name = "PL"

    bal.set_index('Report Date', inplace=True)
    bal.index.name = "BalanceSheet"

    cash.set_index('Report Date', inplace=True)
    cash.index.name = "CashFlow"

    df = pd.concat([pl, bal, cash])
    df = df.T
    df["price"] = price

    df['mcap'] = get_mcap(filename)
    return df

# first = ["Sales","equity", "Profit before tax", "Cash from Operating Activity"]
# margin_roe =["opm", "npm", "roe"]


def add_otherdata(dft):
    exps = ["Power and Fuel", "Other Mfr. Exp",
            "Employee Cost", "Selling and admin", "Other Expenses"]
    dft["equity"] = dft["Equity Share Capital"] + dft["Reserves"]
    dft["expense"] = dft[exps].sum(axis=1) + dft["Change in Inventory"]*-1.0
    dft["OperatingProfit"] = dft["Sales"] - dft["expense"]
    dft["opm"] = (dft["OperatingProfit"]/dft["Sales"])*100.0
    dft["npm"] = (dft["Net profit"]/dft["Sales"])*100.
    dft["d2e"] = dft["Borrowings"]/dft["equity"]
    dft["roe"] = (dft["Net profit"] / dft["equity"])*100.0
    dft["cashbyNP"] = (dft["Cash from Operating Activity"] /
                       dft["Net profit"])*100.0
    dft["GOAR"] = 100.0*(dft["Net profit"].diff() /
                         (dft["Net profit"] - dft["Dividend Amount"]))
    dft["DividendPayout"] = 100.0*(dft["Dividend Amount"]/dft["Net profit"])
    dft["BookValue"] = dft.equity / (dft["No. of Equity Shares"]/1e7)
    dft["siv"] = dft.BookValue*(dft.roe.shift()/12.0)
    dft["siv2"] = dft.BookValue*(dft.roe.shift()/8.0)
    dft["QuickRank"] = (dft.GOAR * (dft.cashbyNP/100.00)
                        * dft.roe) / (1+dft.d2e)
    dft["eps"] = dft["Net profit"] / (dft["No. of Equity Shares"]/1e7)
    dft["cash"] = dft["Cash from Operating Activity"]
    dft["Interest%NP"] = 100.0*(dft["Interest"] /dft["Net profit"])
    return dft

#GOAR = (Net Profit last year -Net Profit preceding year )/(Net Profit last year -Dividend last year ) *100
#quick rank =Growth of Additional Rs * Cash by Net Profit  * Return on equity / (Debt to equity +1)
# cash/np = Cash from operations last year / Net Profit last year
# siv = (Return on equity preceding year/12.00)* Book value


def current_year(df):
    try:
        df.index = pd.to_datetime(df.index)
        year = df.index.year[-1]
    except Exception:
       year = df.index[-1]
    return year


def Average5Year(df):
    return df.iloc[5:, ].mean()


def Average10Year(df):
    return df.mean()


def show_roe_opm_npm(df):
    margin_roe = ["opm", "cashbyNP"]
    second = ["npm", "roe"]
    fig, ax = plt.subplots(1, 2, figsize=(16, 9))
    ax1 = df[margin_roe].T.plot.bar(ax=ax[0], rot=0, title="10 Year History")
    ax1 = df[second].T.plot.bar(ax=ax[1], rot=0, title="10 Year History")
    # ax1.set_xticklabels(map(lambda x: x.year, df.index))
    return fig


def show_roe_opm_npm_line(df):
    fig, ax = plt.subplots(2, 1, figsize=(16, 9), sharex=True)
    ax1 = df[["opm", "cashbyNP"]].plot(ax=ax[0], figsize=(16, 9), rot=0)
    ax1 = df[["npm", "roe"]].plot(ax=ax[1], figsize=(16, 9), rot=0)
    # ax1.set_xticklabels(map(lambda x: x.year, df.index))
    return fig


def sales_any_other_data_line(df):
    fig, ax = plt.subplots(2, 1, sharex=True,  figsize=(16, 9))
    ax1 = df[["Sales", "equity"]].plot(ax=ax[0], rot=0)
    ax2 = df[["Profit before tax", "Cash from Operating Activity"]].plot(
        ax=ax[1], rot=0)
    # ax[0].set_xticklabels(map(lambda x: x.year, df.index))
    return fig


def sales_any_other_data(df):
    first = ["Sales", "equity"]
    second = ["Profit before tax",
              "Cash from Operating Activity"]
    fig, ax = plt.subplots(1, 2, figsize=(16, 9))
    ax1 = df[first].T.plot.bar(ax=ax[0], rot=0, title="10 Year History")
    ax1 = df[second].plot.bar(ax=ax[1], rot=0, title="10 Year History")
    try:
        ax1.set_xticklabels(map(lambda x: x.year, df.index))
    except Exception:
        pass
    return fig


def getHalfYearlyPercent(quaters):
    fig, ax = plt.subplots(1, 1)
    thisyear = quaters[quaters.columns[[-1, -2]]].sum(axis=1)
    lastyear = quaters[quaters.columns[[-5, -6]]].sum(axis=1)
    df = pd.DataFrame({'this': thisyear, 'last': lastyear})
    df = df.T.pct_change().iloc[1]
    p = df.plot(ax=ax, kind='bar', figsize=(16, 9),
                title='Half Yearly %', table=True)
    p.get_xaxis().set_visible(False)
    return fig


def getHalfYearly(quaters):
    fig, ax = plt.subplots(1, 1)
    thisyear = quaters[quaters.columns[[-1, -2]]].sum(axis=1)
    lastyear = quaters[quaters.columns[[-5, -6]]].sum(axis=1)
    df = pd.DataFrame({'this': thisyear, 'last': lastyear})
    p = df.plot(ax=ax, kind='bar', figsize=(16, 9),
                title='Half Yearly', table=True)
    p.get_xaxis().set_visible(False)
    return fig


def getQuaterlyFigures(quaters):
    fig, ax = plt.subplots(1, 1)
    data = quaters[quaters.columns[[-5, -2, -1]]]
    p = data.plot(ax=ax, kind='bar', figsize=(16, 9),
                  rot=0, table=True, title='Quarterly Result')
    p.get_xaxis().set_visible(False)
    return fig


def compareLastQuarterResults(quaters):
    fig, ax = plt.subplots(1, 1)
    data = quaters[quaters.columns[[-5, -1]]]
    p = data.plot(ax=ax, kind='bar', figsize=(16, 9),
                  rot=0, table=True, title='Last Year Quarter')
    p.get_xaxis().set_visible(False)
    return fig


def lastYearPercentChange(quaters):
    fig, ax = plt.subplots(1, 1)
    data = quaters[quaters.columns[[-5, -1]]]
    df = data.T.pct_change().iloc[-1]
    p = df.plot(ax=ax, kind='bar', rot=0,
                title='Last Year % change', figsize=(16, 9), table=True)
    p.get_xaxis().set_visible(False)
    return fig


def getQuarterPercentChange(quaters):
    fig, ax = plt.subplots(1, 1)
    data = quaters[quaters.columns[[-2, -1]]]
    df = data.T.pct_change().iloc[-1]
    p = df.plot(ax=ax, kind='bar', figsize=(16, 9), rot=0,
                title='% Quarter change', table=True)
    p.get_xaxis().set_visible(False)
    return fig


def getLastFiveQuarters(quaters):
    fig, ax = plt.subplots(1, 1)
    data = quaters[quaters.columns[[-5, -4, -3, -2, -1]]]
    p = data.plot(ax=ax, kind='bar', figsize=(
        16, 9), rot=0, title='Last 5 Quarters', table=True)
    p.get_xaxis().set_visible(False)
    return fig


def TitleSlide(text='Thank You'):
    fig = plt.figure(figsize=(16, 6))
    plt.text(0.25, 0.5, text, fontsize=15)
    plt.axis('off')
    return fig


def price_siv(df):
    fig, ax = plt.subplots(1, 1, figsize=(16, 9))
    cols = ["price", "siv", "siv2"]
    ax1 = df[cols].plot(ax=ax, rot=0, title="Price with intrinsic value")
    return fig


def compareLastYearResults(df):
    cols = ["Sales", "expense", "OperatingProfit", "Net profit",
            "cash", "Dividend Amount", "Borrowings", "Interest"]
    fig, ax = plt.subplots(2, 1, figsize=(16, 9))
    data = df.iloc[-2:, :]
    ip = data[cols].T.plot(ax=ax[0], kind='bar',
                           rot=0, title='Last Year Comparision')

    p = data[cols].pct_change().iloc[-1]*100
    ia = p.plot.bar(ax=ax[1], rot=0, title='% change')
    # p.get_xaxis().set_visible(False)
    return fig


def compareLastYearResults_roe(df):
    cols2 = ["opm", "npm", "roe", "GOAR", "Interest%NP", "d2e"]
    fig, ax = plt.subplots(2, 1, figsize=(16, 9))
    data = df.iloc[-2:, :]
    ip = data[cols2].T.plot(ax=ax[0], kind='bar',
                            rot=0, title='Last Year Comparision', alpha=0.8)

    p = data[cols2].pct_change().iloc[-1]*100
    ia = p.plot.bar(ax=ax[1], rot=0, title='% change', alpha=0.8)
    # p.get_xaxis().set_visible(False)
    return fig

# Tables
# fig, ax = plt.subplots(2, 1, figsize=(16, 9))
# ax1 = ax[1].table(cellText=data[cols2].values, colLabels=data[cols2].columns, loc='center', rowLabels=data[cols2].index)
# ax1 = ax[0].table(cellText=data[cols].values, colLabels=data[cols].columns, loc='center', rowLabels=data[cols].index)
# plt.show()


def mainpage(df, cyear):
    text = """
    Sale: {:0.1f}      Net Profit: {:0.1f}      Cash: {:0.1f}   Book Value: {:0.1f} 

    Dividend: {:0.1f}  DebttoEquity: {:0.2f}    OPM: {:0.2f}%   NPM: {:0.2f}%

    ROE: {:0.2f}%     GOAR: {:0.2f}%    Interest paid: {:0.2f}   Market cap: {:0.2f}
    """

    fig, ax = plt.subplots(2, 2, figsize=(16, 6))
    ax = ax.ravel()
    valss = ["Sales", "Net profit", "cash", "BookValue",
             "Dividend Amount", "d2e", "opm", "npm", "roe", "GOAR", "Interest", "mcap"]
    vals = df[valss].values
    ax[0].text(0.1, 0.1, text.format(*vals), fontsize=15)
    ax[0].set_title("For Year {} ".format(cyear))
    ax[0].axis('off')
    ax[1].axis('off')
    ax1 = df[["Sales", "Net profit", "cash", "BookValue",
              "Dividend Amount", "Interest"]].plot.bar(ax=ax[2], rot=0)
    ax1 = df[["opm", "npm", "roe", "GOAR", "Interest%NP"]
             ].plot.bar(ax=ax[3], rot=0)
    ax[2].grid(False)
    ax[3].grid(False)
    # ax[1].axis('off')
    return fig


def get_overall(filename):
    """
    Make a basic annual report diagram for the file.
    """
    figures = []
    # fig = TitleSlide(os.path.basename(filename).replace('.xlsx', ''))
    # figures.append(fig)

    df = get_pl_bal_cash_price(filename)
    df = add_otherdata(df)

    fig = mainpage(df.iloc[-1], current_year(df))
    figures.append(fig)

    fig = sales_any_other_data(df)
    figures.append(fig)

    fig = sales_any_other_data_line(df)
    figures.append(fig)

    fig = show_roe_opm_npm(df)
    figures.append(fig)

    fig = show_roe_opm_npm_line(df)
    figures.append(fig)

    fig = compareLastYearResults(df)
    figures.append(fig)

    fig = compareLastYearResults_roe(df)
    figures.append(fig)

    fig = price_siv(df)
    figures.append(fig)

    return figures


def GetDataAsFigures(filename):
    ''' Takes screenr.in excel sheet and analyses '''
    figures = []

    fig = TitleSlide(os.path.basename(filename).replace('.xlsx', ''))
    figures.append(fig)

    fig = TitleSlide("Annual Summary")
    figures.append(fig)

    figs = get_overall(filename)
    figures.extend(figs)

    fig = TitleSlide("Quarterly Summary")
    figures.append(fig)

    quaters = getQuarters(filename)
    fig = getLastFiveQuarters(quaters)
    figures.append(fig)

    fig = getQuaterlyFigures(quaters)
    figures.append(fig)

    fig = compareLastQuarterResults(quaters)
    figures.append(fig)

    fig = lastYearPercentChange(quaters)
    figures.append(fig)

    fig = getQuarterPercentChange(quaters)
    figures.append(fig)

    lastmonth = quaters.columns[-1]
    lastmonth = lastmonth.month

    if (lastmonth == 9):
        fig = getHalfYearly(quaters)
        figures.append(fig)

        fig = getHalfYearlyPercent(quaters)
        figures.append(fig)

    elif (lastmonth == 12):
        pass

    fig = TitleSlide("Thank You")
    figures.append(fig)

    return figures


def CreatePDFFileFromFigures(figures, filename):
    ''' Creates PDF file from figures list'''
    pdf = PdfPages(filename)
    for fig in figures:
        pdf.savefig(fig)

    pdf.close()


def AnalyzeCreatePDFFile(filename):
    '''Analyses screenr for quaterly results'''
    outfile = filename.replace('.xlsx', '.pdf')

    figures = GetDataAsFigures(filename)
    CreatePDFFileFromFigures(figures, outfile)


def main():
    parser = argparse.ArgumentParser(description="Analyze quarter")
    parser.add_argument("file", type=str,  help="screener file")

    args = parser.parse_args()

    assert os.path.exists(args.file)
    AnalyzeCreatePDFFile(args.file)


if __name__ == "__main__":
    main()
