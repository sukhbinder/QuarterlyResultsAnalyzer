import stock_analyze.analyze2 as a2
import stock_analyze.AnalyzeQuarters as a1
import matplotlib.pyplot as plt

from itertools import cycle

import argparse
import os
import numpy as np
#   options: { indexAxis: 'y'}

# scales: {x: { stacked: true, }, y: { stacked: true}}

def make_data_tag(data_dict, lab="", datalabel="data"):
    colors = cycle(['#377eb8', '#ff7f00', '#4daf4a',
                    '#f781bf', '#a65628', '#984ea3',
                    '#999999', '#e41a1c', '#dede00'])
    # colors = cycle(['lightgreen', 'sky', 'gold','orange', 'red', 'purple','pink', 'silver', '#dede00'])
    output_text = "const {1} = {{ labels: {0}, datasets: [\n".format(
        lab, datalabel)
    data_line = " {{ label: '{0}', backgroundColor: '{1}', borderColor: '{1}', data: {2} }},"

    for key, val in data_dict.items():
        color = next(colors)
        output_text += data_line.format(key, color, val) + "\n"
    output_text += "]};"
    return output_text


def make_config_tag(charttype="bar", datalabel="data"):
    output_text = "\nconst config_{1} ={{type: '{0}', data:{1}, options: {{}} }};\n\n".format(
        charttype, datalabel)
    return output_text


scripttag = """

function loadchart(){
    var myChart = new Chart(
    document.getElementById('capital_history'),
    config
  );
  
}

$(document).ready(loadchart);
</script>
</html>
"""


htmlhead = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.5.1/jquery.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <title>{}</title>

</head>
<body>
"""


scripend = """
$(document).ready(loadchart);
</script>
</html>
"""


class Graph:
    def __init__(self, doc, name):
        self._parent = doc
        self._name = name
        self._type = "bar"
        self._label = ""
        self._data = None
        self._config = None

    def gtype(self, gtype):
        assert gtype in ["bar", "line", "doughnot",
                         "bubble"], "Graph type not yet supported"
        self._type = gtype
        return self

    def data(self, data_dict):
        self._data = data_dict
        return self

    def label(self, label):
        self._label = label
        return self

    def get_tag(self):
        output_text = make_data_tag(self._data, self._label, self._name) + \
            make_config_tag(charttype=self._type, datalabel=self._name)
        return output_text


def sanitize(name):
    return name.strip().replace(" ", "_")


class Doc:
    def __init__(self, name="default") -> None:
        name = sanitize(name)
        self._name = name
        self._output_file= "index_{}.html".format(name)
        self._graphs = []

    def add_graph(self, name=None):
        if name is None:
            name = "Div_{}".format(len(self._graphs)+1)
        assert name not in [g._name for g in self._graphs], "Name exists"
        name = sanitize(name)
        graph = Graph(self, name)
        self._graphs.append(graph)
        return graph

    def _add_divs(self):
        output_str = """<div class="chart-container" style="height:70%; width:90%"><h2> {0} </h2><canvas id="{0}"></canvas></div>\n"""

        output_text = "\n"
        for graph in self._graphs:
            output_text += output_str.format(graph._name)

        output_text += "</body>"
        return output_text

    def _get_chartsjs(self):
        output_text = "function loadchart() {"
        text = "var nw{0} = new Chart( document.getElementById('{0}'), config_{0} );"
        for graph in self._graphs:
            output_text += text.format(graph._name)
        output_text += " \n};"
        return output_text

    def _get_chart_data(self):
        output_text = "\n<script>"
        for graph in self._graphs:
            output_text += graph.get_tag()
        return output_text

    def write_doc(self, outputpath=None):
        out_text = htmlhead.format(self._name)
        out_text += self._add_divs()
        out_text += self._get_chart_data()
        out_text += self._get_chartsjs()
        out_text += scripend
        if outputpath is None:
            outfile = self._output_path
        else:
            outfile = os.path.join(outputpath, self._output_file)
        with open(outfile, "w") as fout:
            fout.write(out_text)

def make_plots(fname):
    # fname = r"/Users/sukhbindersingh/Downloads/Banco Products.xlsx"
    df, _ = a1.get_pl_bal_cash_price(fname)
    df1 = a1.add_otherdata(df)

    df1["grossassets"] = df1["Net Block"]+df1["Depreciation"]
    df1["asset_turnover"] = df1.Sales/df1.grossassets
    df1["equity_multiplier"] = df1.grossassets/df1.equity


    df1["sales_growth"] = df1.Sales.pct_change()*100.0
    df1["exp_pect_sales"] = df1.expense/df1.Sales * 100.0
    # replace inf to nan as
  

    trend = a2.get_trend_analysis(df1.T)
    trend.replace([np.inf, -np.inf], np.nan, inplace=True)
    trend.fillna(0, inplace=True)


    df1.replace([np.inf, -np.inf], np.nan, inplace=True)
    df1.fillna(0, inplace=True)

    items = ["Sales", "expense", "equity", "cash", "Investments", "Interest",
            "Profit before tax", "Dividend Amount"]
    ratios = ["opm", "npm", "d2e", "roe", "eps", "sales_growth"]

    capital_history = ["equity", "Borrowings", "Profit before tax"]

    business_out = ["Sales", "expense", "Interest",
                    "Profit before tax", "Net profit", "Dividend Amount"]

    roe_du = ["asset_turnover", "equity_multiplier", "npm", "roe", "sales_growth"]

    cash_PBT = df1[["cash", "Profit before tax"]].iloc[-5:].sum().to_dict()


    #

    a = df1[items].fillna(0).round(1).T.iloc[:, ::-3].T
    labels = [i.strftime("%Y") for i in a.index.to_list()]
    aa = a.to_dict(orient="list")

    outfile = os.path.basename(fname.replace('.xlsx', ''))
    doc = Doc(outfile)

    a = df1[business_out].iloc[:, ::-1]
    labels = [i.strftime("%Y") for i in a.index.to_list()]
    aa = a.to_dict(orient="list")
    doc.add_graph("Business_Performance").gtype("bar").label(labels).data(aa)


    a = df1[capital_history]
    labels = [i.strftime("%Y") for i in a.index.to_list()]
    aa = a.to_dict(orient="list")
    doc.add_graph("Capital_History").gtype("bar").label(labels).data(aa)


    a = df1[items].fillna(0).round(1).T.iloc[:, ::-3].T
    labels = [i.strftime("%Y") for i in a.index.to_list()]
    aa = a.to_dict(orient="list")
    doc.add_graph("Every_3_years").gtype("bar").label(labels).data(aa)

    a = df1[ratios].fillna(0).round(1).T.iloc[:, ::-3].T
    labels = [i.strftime("%Y") for i in a.index.to_list()]
    aa = a.to_dict(orient="list")
    doc.add_graph("Every_3_years_ratios").gtype("bar").label(labels).data(aa)

    a = df1[roe_du]
    labels = [i.strftime("%Y") for i in a.index.to_list()]
    aa = a.to_dict(orient="list")
    doc.add_graph("Ratios").gtype("line").label(labels).data(aa)

    a = df1[["exp_pect_sales"]]
    labels = [i.strftime("%Y") for i in a.index.to_list()]
    aa = a.to_dict(orient="list")
    doc.add_graph("ExpenseVsSales").gtype("bar").label(labels).data(aa)

    a = trend.T[business_out]
    labels = [i.strftime("%Y") for i in a.index.to_list()]
    aa = a.to_dict(orient="list")
    doc.add_graph("Trend").gtype("bar").label(labels).data(aa)


    a = trend.T[ratios]
    
    labels = [i.strftime("%Y") for i in a.index.to_list()]
    aa = a.to_dict(orient="list")
    doc.add_graph("Trend_Ratios").gtype("bar").label(labels).data(aa)

    # a = cash_PBT
    # aa = a.to_dict(orient="list")
    # doc.add_graph("Cash and PBT (5yrs)").gtype("bar").label(labels).data(aa)

    outputpath = os.path.join(os.path.dirname(fname))
    doc.write_doc(outputpath)

def main():
    parser = argparse.ArgumentParser(description="Analyze quarter")
    parser.add_argument("file", type=str,  help="screener file")

    args = parser.parse_args()

    assert os.path.exists(args.file)
    make_plots(args.file)





# with open("index.html", "w") as f_out:
#     f_out.write(htmlheader)
#     f_out.write(
#         make_data_tag(aa,labels, "chist")
#        + make_config_tag(charttype="bar", datalabel="chist")
#        + scripttag)
