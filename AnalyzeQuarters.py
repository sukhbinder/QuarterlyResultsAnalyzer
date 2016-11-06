# coding: utf-8
# 

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib
matplotlib.style.use('ggplot')
from matplotlib.backends.backend_pdf import PdfPages



def getQuarters(filename):
    xp=pd.ExcelFile(filename)
    nesco=xp.parse('Data Sheet',header=40)
    quaters=nesco.drop(nesco.index[range(9,52)])

    quaters.set_index('Report Date',inplace=True)
    quaters.index.name = 'Quaters'
    return quaters

#quaters=getQuarters(fpath+fname)

def getQuaterlyFigures(quaters):
    fig,ax=plt.subplots(1,1)
    p=quaters[[-5,-2,-1]].plot(ax=ax,kind='bar',figsize=(16,9),rot=0,table=True,title='Quarterly Result');
    p.get_xaxis().set_visible(False)
    return fig
	
def compareLastQuarterResults(quaters):
    fig,ax=plt.subplots(1,1)
    p=quaters[[-5,-1]].plot(ax=ax,kind='bar',figsize=(16,9),rot=0,table=True,title='Last Year Quarter');
    p.get_xaxis().set_visible(False)
    return fig

def lastYearPercentChange(quaters):
    fig,ax=plt.subplots(1,1)
    df=quaters[[-5,-1]].T.pct_change().iloc[-1]
    p=df.plot(ax=ax,kind='bar',rot=0,title='Last Year % change',figsize=(16,9),table=True );
    p.get_xaxis().set_visible(False)
    return fig

def getQuarterPercentChange(quaters):
    fig,ax=plt.subplots(1,1)
    df=quaters[[-2,-1]].T.pct_change().iloc[-1]
    p=df.plot(ax=ax,kind='bar',figsize=(16,9),rot=0,title='% Quarter change',table=True);
    p.get_xaxis().set_visible(False)
    return fig

def getLastFiveQuarters(quaters):
    fig,ax=plt.subplots(1,1)
    p=quaters[[-5,-4,-3,-2,-1]].plot(ax=ax,kind='bar',figsize=(16,9),rot=0,title='Last 5 Quarters',table=True);
    p.get_xaxis().set_visible(False)
    return fig


def TitleSlide(text='Thank You'):
    fig=plt.figure(figsize=(16,6))
    plt.text(0.25,0.5,text,fontsize=15)
    plt.axis('off')
    return fig

def GetDataAsFigures(filename):
    ''' Takes screer.in excel sheet and analyses '''
    figures=[]
    
    quaters=getQuarters(filename)  
    fig=TitleSlide(filename.replace('.xlsx',''))
    figures.append(fig)
    
    fig=getLastFiveQuarters(quaters)
    figures.append(fig)
    
    fig=getQuaterlyFigures(quaters)
    figures.append(fig)
    
    fig=compareLastQuarterResults(quaters)
    figures.append(fig)
    
    fig=lastYearPercentChange(quaters)
    figures.append(fig)
    
    fig=getQuarterPercentChange(quaters)
    figures.append(fig)
    
    fig=TitleSlide("Thank You")
    figures.append(fig)
                
    return figures


def CreatePDFFileFromFigures(figures,filename):
    ''' Creates PDF file from figures list'''
    pdf=PdfPages(filename)
    for fig in figures:
        pdf.savefig(fig)
        
    pdf.close()

def AnalyzeCreatePDFFile(filename):
    '''Analyses screenr for quaterly results'''
    import os
    outfile=os.path.split(filename)[1].replace('.xlsx','')+'.pdf'
    figures=GetDataAsFigures(filename)
    CreatePDFFileFromFigures(figures,outfile)
    

if __name__ == "__main__":
    fpath=r'tests/'
    fname='Tasty Bite Eat.xlsx'
    AnalyzeCreatePDFFile(fpath+fname)


