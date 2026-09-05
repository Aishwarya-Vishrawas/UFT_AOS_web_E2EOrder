"""Build the NIFTY 50 vs Pidilite Industries daily-return workbook.

Pulls five years of daily bars for both instruments from the Yahoo Finance chart
API, aligns them on common trading sessions, and writes an .xlsx with the raw
prices, live daily-percentage-return formulas, an OLS regression of Pidilite's
returns (y) on NIFTY 50's returns (x), and a scatter chart with a fitted line.

    python3 analysis/build_returns_workbook.py

Requires: pandas, numpy, openpyxl, requests.
"""

import os
import requests, pandas as pd, json, sys

UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"}

def grab(sym):
    url = f"https://query1.finance.yahoo.com/v8/finance/chart/{sym}"
    p = {"range": "5y", "interval": "1d", "includeAdjustedClose": "true", "events": "div|split"}
    r = requests.get(url, params=p, headers=UA, timeout=60)
    r.raise_for_status()
    res = r.json()["chart"]["result"][0]
    ts = res["timestamp"]
    q = res["indicators"]["quote"][0]
    adj = res["indicators"].get("adjclose", [{}])[0].get("adjclose", q["close"])
    df = pd.DataFrame({
        "Date": pd.to_datetime(ts, unit="s", utc=True).tz_convert("Asia/Kolkata").normalize().tz_localize(None),
        "Close": q["close"], "AdjClose": adj,
    }).dropna()
    df = df.drop_duplicates(subset="Date").set_index("Date").sort_index()
    print(sym, df.shape, df.index.min().date(), df.index.max().date())
    return df


from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import ScatterChart, Reference, Series
from openpyxl.chart.trendline import Trendline
from openpyxl.chart.marker import Marker
from openpyxl.comments import Comment

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                   "NIFTY50_Pidilite_Daily_Returns_Beta.xlsx")

n = grab("^NSEI").rename(columns={"Close": "N_Close", "AdjClose": "N_Adj"})
p = grab("PIDILITIND.NS").rename(columns={"Close": "P_Close", "AdjClose": "P_Adj"})
df = n.join(p, how="inner").sort_index()
print("aligned rows:", len(df), df.index.min().date(), "->", df.index.max().date())

FONT = "Arial"
hdr_fill = PatternFill("solid", fgColor="1F3864")
hdr_font = Font(name=FONT, bold=True, color="FFFFFF", size=10)
base = Font(name=FONT, size=10)
bold = Font(name=FONT, size=10, bold=True)
title = Font(name=FONT, size=14, bold=True, color="1F3864")
blue = Font(name=FONT, size=10, color="0000FF")
thin = Side(style="thin", color="BFBFBF")
box = Border(left=thin, right=thin, top=thin, bottom=thin)
key_fill = PatternFill("solid", fgColor="FFF2CC")

import numpy as np
_r = df[["N_Adj", "P_Adj"]].pct_change().dropna()
_x, _y = _r.N_Adj.values, _r.P_Adj.values
_slope, _inter = np.polyfit(_x, _y, 1)
_corr = float(np.corrcoef(_x, _y)[0, 1])
CHECK = (f"n = {len(_x)}  |  slope (beta) = {_slope:.4f}  |  intercept = {_inter*100:.4f}% per day  |  "
         f"R-squared = {_corr**2:.4f}  |  correlation = {_corr:.4f}")
print("CHECK:", CHECK)

wb = Workbook()

# ---------------- Data sheet ----------------
ws = wb["Sheet"]; ws.title = "Data"
headers = ["Date", "NIFTY 50 Close", "NIFTY 50 Adj Close", "Pidilite Close (INR)",
           "Pidilite Adj Close (INR)", "NIFTY 50 Daily Return %", "Pidilite Daily Return %"]
ws.append(headers)
for c in range(1, len(headers) + 1):
    cell = ws.cell(row=1, column=c)
    cell.font = hdr_font; cell.fill = hdr_fill
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
ws.cell(row=1, column=6).comment = Comment(
    "Daily simple return computed on Adjusted Close: (Adj_t / Adj_t-1) - 1.\n"
    "Stored as a fraction, displayed as a percentage.\n"
    "Source: Yahoo Finance chart API, symbol ^NSEI.", "Analysis")
ws.cell(row=1, column=7).comment = Comment(
    "Daily simple return computed on Adjusted Close: (Adj_t / Adj_t-1) - 1.\n"
    "Adjusted Close is split- and dividend-adjusted, so this is a total return.\n"
    "Source: Yahoo Finance chart API, symbol PIDILITIND.NS.", "Analysis")

for i, (dt, row) in enumerate(df.iterrows()):
    r = i + 2
    ws.cell(row=r, column=1, value=dt.date()).number_format = "yyyy-mm-dd"
    ws.cell(row=r, column=2, value=round(float(row.N_Close), 2))
    ws.cell(row=r, column=3, value=round(float(row.N_Adj), 2))
    ws.cell(row=r, column=4, value=round(float(row.P_Close), 2))
    ws.cell(row=r, column=5, value=round(float(row.P_Adj), 2))
    if r > 2:
        ws.cell(row=r, column=6, value=f"=C{r}/C{r-1}-1")
        ws.cell(row=r, column=7, value=f"=E{r}/E{r-1}-1")
    for c in range(1, 8):
        cl = ws.cell(row=r, column=c)
        cl.font = base
        if c in (2, 3, 4, 5):
            cl.number_format = "#,##0.00"
        elif c in (6, 7):
            cl.number_format = "0.00%;(0.00%);-"

last = len(df) + 1          # last data row
first_ret = 3               # first row that has a return
ws.freeze_panes = "B2"
ws.auto_filter.ref = f"A1:G{last}"
for col, w in zip("ABCDEFG", [12, 15, 17, 18, 20, 20, 20]):
    ws.column_dimensions[col].width = w
ws.row_dimensions[1].height = 34

X = f"Data!$F${first_ret}:$F${last}"   # NIFTY 50 returns  (x axis)
Y = f"Data!$G${first_ret}:$G${last}"   # Pidilite returns  (y axis)

# ---------------- Regression sheet ----------------
rs = wb.create_sheet("Regression")
rs["A1"] = "Pidilite Industries vs NIFTY 50 — Daily Return Regression"; rs["A1"].font = title
rs["A2"] = "OLS fit:  Pidilite daily return (y)  =  slope x  NIFTY 50 daily return (x)  +  intercept"
rs["A2"].font = Font(name=FONT, size=10, italic=True)

# Rows are declared with {tags}; the row each tag lands on is resolved below,
# so no cross-reference is hardcoded to a row number.
rows = [
    ("hdr1", "Regression results", None, None, None),
    ("slope", "Slope (beta vs NIFTY 50)", f"=SLOPE({Y},{X})", "0.0000", "Change in Pidilite's daily return per 1 unit change in NIFTY 50's daily return"),
    ("alpha", "Intercept (alpha, daily)", f"=INTERCEPT({Y},{X})", "0.0000%", "Average daily return not explained by the index"),
    ("r2", "R-squared", f"=RSQ({Y},{X})", "0.0000", "Share of Pidilite's daily variance explained by the index"),
    ("corr", "Correlation (Pearson r)", f"=CORREL({Y},{X})", "0.0000", "Square root of R-squared"),
    ("se", "Standard error of slope", f"=STEYX({Y},{X})/SQRT(DEVSQ({X}))", "0.0000", "Sampling error of the slope estimate"),
    ("t", "t-statistic of slope", "=B{slope}/B{se}", "0.00", "Slope divided by its standard error; |t| > 2 is significant at ~5%"),
    ("n", "Observations (n)", f"=COUNT({Y})", "#,##0", "Trading days with a return for both series"),
    (None, None, None, None, None),
    ("hdr2", "Descriptive statistics", None, None, None),
    ("mx", "Mean daily return - NIFTY 50", f"=AVERAGE({X})", "0.0000%", None),
    ("my", "Mean daily return - Pidilite", f"=AVERAGE({Y})", "0.0000%", None),
    ("sx", "Daily volatility - NIFTY 50", f"=STDEV({X})", "0.0000%", "Sample standard deviation of daily returns"),
    ("sy", "Daily volatility - Pidilite", f"=STDEV({Y})", "0.0000%", "Sample standard deviation of daily returns"),
    ("ax", "Annualised volatility - NIFTY 50", "=B{sx}*SQRT($B${td})", "0.00%", None),
    ("ay", "Annualised volatility - Pidilite", "=B{sy}*SQRT($B${td})", "0.00%", None),
    ("cov", "Covariance (sample)", "=SUMPRODUCT((" + X + "-AVERAGE(" + X + "))*(" + Y + "-AVERAGE(" + Y + ")))/(B{n}-1)", "0.00000000", "Sample covariance of the two return series"),
    ("var", "Variance of NIFTY 50 (sample)", f"=VAR({X})", "0.00000000", None),
    ("chk", "Slope cross-check (Cov / Var)", "=B{cov}/B{var}", "0.0000", "Must equal the SLOPE result above"),
    (None, None, None, None, None),
    ("td", "Trading days per year (assumption)", 252, "#,##0", "Standard convention for annualising Indian equity volatility"),
]

# Pass 1: resolve the row each tag occupies.
ROW = {}
_r = 3
for tag, *_ in rows:
    if tag:
        ROW[tag] = _r
    _r += 1

r = 3
for tag, label, val, fmt, note in rows:
    if label is None:
        r += 1; continue
    if isinstance(val, str):
        val = val.format(**ROW)
    c1 = rs.cell(row=r, column=1, value=label)
    if val is None:
        c1.font = Font(name=FONT, size=11, bold=True, color="1F3864")
    else:
        c1.font = base
        c2 = rs.cell(row=r, column=2, value=val)
        c2.font = blue if isinstance(val, (int, float)) else bold
        c2.number_format = fmt
        c2.border = box
        if isinstance(val, (int, float)):
            c2.fill = key_fill
        if note:
            rs.cell(row=r, column=3, value=note).font = Font(name=FONT, size=9, italic=True, color="595959")
    r += 1

rs.cell(row=ROW["slope"], column=1).fill = key_fill
rs.cell(row=ROW["slope"], column=2).font = Font(name=FONT, size=12, bold=True, color="C00000")
rs.column_dimensions["A"].width = 34
rs.column_dimensions["B"].width = 16
rs.column_dimensions["C"].width = 66

# Scatter chart
ch = ScatterChart()
ch.title = "Pidilite vs NIFTY 50 — daily % returns"
ch.style = 13
ch.x_axis.title = "NIFTY 50 daily return %"
ch.y_axis.title = "Pidilite daily return %"
ch.x_axis.numFmt = "0.0%"
ch.y_axis.numFmt = "0.0%"
ch.height = 12; ch.width = 18
ch.legend = None
xref = Reference(ws, min_col=6, min_row=first_ret, max_row=last)
yref = Reference(ws, min_col=7, min_row=first_ret, max_row=last)
s = Series(yref, xref, title="Daily returns")
s.marker = Marker(symbol="circle", size=3)
s.graphicalProperties.line.noFill = True
s.trendline = Trendline(trendlineType="linear", dispEq=True, dispRSqr=True)
ch.series.append(s)
rs.add_chart(ch, "E3")

# ---------------- Notes sheet ----------------
ns = wb.create_sheet("Notes")
ns["A1"] = "Sources, definitions and assumptions"; ns["A1"].font = title
notes = [
    ("Instruments", ""),
    ("NIFTY 50 (x axis)", "Yahoo Finance symbol ^NSEI — NSE India benchmark index."),
    ("Pidilite (y axis)", "Yahoo Finance symbol PIDILITIND.NS — Pidilite Industries Ltd, NSE India. "
                          "Interpreted as the requested 'Pidilite Inc'; there is no separately listed entity by that exact name."),
    ("", ""),
    ("Data", ""),
    ("Source", "Yahoo Finance chart API (query1.finance.yahoo.com/v8/finance/chart), range=5y, interval=1d, retrieved 2026-09-05."),
    ("Period covered", f"{df.index.min().date()} to {df.index.max().date()} (5 years of daily bars)."),
    ("Rows on Data sheet", f"{len(df)} trading days common to both series; "
                           f"{len(df)-1} daily returns (the first date has no prior close)."),
    ("Alignment", "Dates are inner-joined, so only sessions on which both the index and the stock traded are kept. "
                  "This avoids spurious zero or stale returns on non-common holidays."),
    ("Prices", "Both raw Close and Adjusted Close are shown. Adjusted Close is corrected for splits and dividends."),
    ("", ""),
    ("Calculations", ""),
    ("Daily % return", "(Adjusted Close today / Adjusted Close previous session) - 1, computed on Adjusted Close so the "
                       "series is a total return. Stored as a fraction and displayed as a percentage."),
    ("Slope", "Ordinary least squares SLOPE(y, x) with NIFTY 50 returns as x and Pidilite returns as y. "
              "This is Pidilite's raw (unadjusted) beta against the NIFTY 50."),
    ("Units", "Slope is unit-free: it is identical whether both series are expressed as fractions or as percentages."),
    ("Annualisation", "Daily volatility x SQRT(252). 252 trading days per year is an assumption, set in cell Regression!B21."),
    ("", ""),
    ("Verification", ""),
    ("Independent check", "The workbook's cells hold live formulas, not pasted numbers, and Excel recalculates them on open. "
                          "The same regression computed independently in Python (numpy least squares) at build time gave: " + CHECK + ". "
                          "The Regression sheet's 'Slope cross-check (Cov / Var)' cell is a second, in-sheet check that must match the SLOPE result."),
    ("Recalculation", "This file was written by openpyxl, which stores formulas without cached results. "
                      "Full-calculation-on-load is set, so Excel or LibreOffice fills in every value the first time the file is opened."),
    ("", ""),
    ("Caveats", ""),
    ("Raw beta", "No Blume or Vasicek adjustment and no risk-free rate is subtracted; this is a total-return beta, not a CAPM excess-return beta."),
    ("Stability", "Beta estimated over a 5-year window is a single point estimate and is not stable over sub-periods."),
]
r = 3
for a, b in notes:
    if a and not b:
        ns.cell(row=r, column=1, value=a).font = Font(name=FONT, size=11, bold=True, color="1F3864")
    elif a:
        ns.cell(row=r, column=1, value=a).font = bold
        c = ns.cell(row=r, column=2, value=b); c.font = base
        c.alignment = Alignment(wrap_text=True, vertical="top")
    r += 1
ns.column_dimensions["A"].width = 24
ns.column_dimensions["B"].width = 110

wb.calculation.fullCalcOnLoad = True
wb.save(OUT)
print("saved", OUT)
