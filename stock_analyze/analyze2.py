import stock_analyze.AnalyzeQuarters as sa
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
import tempfile
import os
import argparse



def get_data_in_excel(fname, figures=False):
    df = sa.get_pl_bal_cash_price(fname)
    df1 = sa.add_otherdata(df)

    wb = openpyxl.Workbook()
    ws = wb.active
    dd =df1.fillna(0).round(1).T
    for r in dataframe_to_rows(dd, index=True, header=True):
        if not "Report Date" in r:
            ws.append(r)

    items=["Sales", "expense", "equity","cash", "Investments", "Interest", "Profit before tax", "Dividend Amount", "opm", "npm", "d2e", "roe", "eps", ]

    dd1=df1[items].fillna(0).round(1).T

    ws1 = wb.create_sheet("Skip_Years")
    for r in dataframe_to_rows(dd1.iloc[:,::-3], index=True, header=True):
        ws1.append(r)


    cs = sa.get_commonsize_analysis(df1)
    ws1 = wb.create_sheet("Common_sizes")
    for r in dataframe_to_rows(cs, index=True, header=True):
        ws1.append(r)

    outfile = os.path.join(os.path.dirname(fname), "AN_{}".format(os.path.basename(fname)))
    
    with tempfile.TemporaryDirectory(prefix="work_") as tempdir:
        if figures:
            figures_in_workbook(fname, wb, tempdir)
        wb.save(outfile)

def figures_in_workbook(fname, wb, tempdir):
    fig = sa.get_overall(fname)
    for i, f in enumerate(fig):
        ws1 = wb.create_sheet(str(i))
        figname = "{}.png".format(i)
        figname = os.path.join(tempdir, figname)
        f.savefig(figname, dpi=150)
        img = openpyxl.drawing.image.Image(figname)
        ws1.add_image(img)
        a=ws1.cell(1,2,1)


def main():
    parser = argparse.ArgumentParser(description="Analyze quarter")
    parser.add_argument("file", type=str,  help="screener file")

    args = parser.parse_args()

    assert os.path.exists(args.file)
    get_data_in_excel(args.file)
