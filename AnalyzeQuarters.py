# coding: utf-8
#

from matplotlib.backends.backend_pdf import PdfPages
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib
import argparse

import os


matplotlib.style.use('ggplot')


def getQuarters(filename):
    xp = pd.ExcelFile(filename)
    nesco = xp.parse('Data Sheet', header=40)
    quaters = nesco.drop(nesco.index[range(9, 52)])

    quaters.set_index('Report Date', inplace=True)
    quaters.index.name = 'Quaters'
    return quaters

#quaters=getQuarters(fpath+fname)


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


def GetDataAsFigures(filename):
    ''' Takes screenr.in excel sheet and analyses '''
    figures = []

    quaters = getQuarters(filename)
    fig = TitleSlide(os.path.basename(filename).replace('.xlsx', ''))
    figures.append(fig)

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
    import os
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
