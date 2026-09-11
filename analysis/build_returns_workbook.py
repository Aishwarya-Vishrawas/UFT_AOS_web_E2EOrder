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


# =====================================================================
# Cost of Equity (CAPM) and Cost of Debt sheets
# =====================================================================
# Every figure below that is NOT derived from the Data sheet is a hardcoded
# input, written in blue and sourced on the Notes sheet.

RF              = 0.0697   # India 10Y G-Sec yield, 10-Sep-2026 (Trading Economics)
DIV_YIELD       = 0.0125   # NIFTY 50 long-run average dividend yield (~1.25%)
ERP_DAMODARAN   = 0.0708   # Damodaran India total equity risk premium, Jan-2026 vintage
NIFTY_10Y_PRICE = 0.1049   # NIFTY 50 price CAGR over the 10 years to 11-Sep-2026
PIDI_10Y_TR     = 0.1646   # Pidilite total-return CAGR over the same 10 years

# Pidilite Industries, CONSOLIDATED, INR crore. Yahoo Finance fundamentals,
# cross-checked line by line against screener.in (both restate the same filed
# consolidated statements). FY label = year ended 31 March.
FY        = ["FY2023", "FY2024", "FY2025", "FY2026"]
INTEREST  = [47.64, 51.19, 50.35, 54.22]      # finance costs
BORROW    = [390.61, 382.47, 454.14, 417.21]  # total borrowings (incl. lease liabilities)
BORROW_EX = [163.26, 131.15, 147.18, 105.91]  # borrowings excluding lease liabilities
PBT       = [1723.24, 2379.35, 2822.70, 3320.17]
TAX       = [434.37, 631.93, 726.53, 849.45]

def put(sh, row, label, value, fmt=None, note=None, kind="formula", indent=0, note_col=3):
    """Write a label/value/note line. kind drives the colour convention."""
    c1 = sh.cell(row=row, column=1, value=("    " * indent) + label)
    c1.font = base
    if value is not None:
        c2 = sh.cell(row=row, column=2, value=value)
        c2.font = {"input": blue, "link": green, "formula": bold}[kind]
        if fmt:
            c2.number_format = fmt
        c2.border = box
        if kind == "input":
            c2.fill = key_fill
    if note:
        sh.cell(row=row, column=note_col, value=note).font = small
    return row + 1

def header(sh, row, text):
    c = sh.cell(row=row, column=1, value=text)
    c.font = Font(name=FONT, size=11, bold=True, color="1F3864")
    return row + 1

def putk(row, label, value, fmt=None, note=None, kind="formula", indent=0):
    """put() bound to the Cost of Equity sheet, with notes parked in column H
    so they never collide with the B:F sensitivity grid."""
    return put(ke, row, label, value, fmt, note, kind, indent, note_col=8)

green = Font(name=FONT, size=10, color="008000")
small = Font(name=FONT, size=9, italic=True, color="595959")

# ---------------- Cost of Equity ----------------
ke = wb.create_sheet("Cost of Equity")
ke["A1"] = "Cost of Equity - CAPM"; ke["A1"].font = title
ke["A2"] = "Ke  =  Risk-free rate  +  Beta x Market risk premium"
ke["A2"].font = Font(name=FONT, size=10, italic=True)
ke["A3"] = "Blue = hardcoded input (sourced on Notes)   Green = linked from another sheet   Black = formula"
ke["A3"].font = small

r = 5
r = header(ke, r, "1. Inputs")
R_RF = r
r = putk(r, "Risk-free rate (India 10Y G-Sec)", RF, "0.00%",
        "Benchmark 10Y G-Sec yield, 10-Sep-2026. Long-dated sovereign yield is the standard Rf for an INR DCF.", "input")
R_BETA = r
r = putk(r, "Beta (Pidilite vs NIFTY 50)", "=Regression!B4", "0.0000",
        "Live link to the OLS slope on the Regression sheet - the beta calculated from 5 years of daily returns.", "link")
R_DIVY = r
r = putk(r, "NIFTY 50 dividend yield", DIV_YIELD, "0.00%",
        "Long-run average (~1.25%). Added to the price-index CAGR because ^NSEI excludes dividends.", "input")

r += 1
r = header(ke, r, "2. Market return and risk premium - three approaches")
ke.cell(row=r, column=1, value="Approach A - trailing 5-year realised (this workbook's own data)").font = Font(name=FONT, size=10, bold=True)
r += 1
R_N5P = r
r = putk(r, "NIFTY 50 price CAGR, 5y", f"=(Data!C{last}/Data!C2)^(365/(Data!A{last}-Data!A2))-1", "0.00%",
        "Computed from the first and last NIFTY 50 adjusted close on the Data sheet.", "formula", 1)
R_RM5 = r
r = putk(r, "Market return Rm (5y, total)", f"=B{R_N5P}+B{R_DIVY}", "0.00%", "Price CAGR plus dividend yield.", "formula", 1)
R_MRP5 = r
r = putk(r, "Market risk premium (5y)", f"=B{R_RM5}-B{R_RF}", "0.00%",
        "Near zero. The 5y window opens close to the Sep-2021 peak, so the realised market return barely clears the G-Sec yield - and on a slightly different start date it falls below it, turning the premium negative.", "formula", 1)
R_KE_A = r
r = putk(r, "Ke - Approach A", f"=B{R_RF}+B{R_BETA}*B{R_MRP5}", "0.00%",
        "Prices equity at roughly the cost of government debt, which cannot be right. An artefact of the start date, not an economic result - shown for completeness, not for use.", "formula", 1)

r += 1
ke.cell(row=r, column=1, value="Approach B - trailing 10-year realised").font = Font(name=FONT, size=10, bold=True)
r += 1
R_N10P = r
r = putk(r, "NIFTY 50 price CAGR, 10y", NIFTY_10Y_PRICE, "0.00%",
        "10 years to 11-Sep-2026, from the full ^NSEI history (outside this workbook's 5y Data sheet).", "input", 1)
R_RM10 = r
r = putk(r, "Market return Rm (10y, total)", f"=B{R_N10P}+B{R_DIVY}", "0.00%", None, "formula", 1)
R_MRP10 = r
r = putk(r, "Market risk premium (10y)", f"=B{R_RM10}-B{R_RF}", "0.00%", "A full cycle, so far less start-date sensitive than 5y.", "formula", 1)
R_KE_B = r
r = putk(r, "Ke - Approach B", f"=B{R_RF}+B{R_BETA}*B{R_MRP10}", "0.00%", None, "formula", 1)

r += 1
ke.cell(row=r, column=1, value="Approach C - forward-looking implied premium (RECOMMENDED)").font = Font(name=FONT, size=10, bold=True)
r += 1
R_ERP = r
r = putk(r, "India equity risk premium", ERP_DAMODARAN, "0.00%",
        "Damodaran implied India ERP, Jan-2026 vintage (mature-market ERP + India country risk premium).", "input", 1)
R_KE_C = r
r = putk(r, "Ke - Approach C", f"=B{R_RF}+B{R_BETA}*B{R_ERP}", "0.00%",
        "Forward-looking and not hostage to the start date. This is the defensible number.", "formula", 1)

r += 1
r = header(ke, r, "3. Selected cost of equity")
R_SEL = r
r = putk(r, "Market risk premium used", f"=B{R_ERP}", "0.00%",
        "Defaults to Approach C. Point this at B{a} or B{b} to switch approach.".format(a=R_MRP5, b=R_MRP10), "formula")
R_KE = r
r = putk(r, "COST OF EQUITY (Ke)", f"=B{R_RF}+B{R_BETA}*B{R_SEL}", "0.00%", "Ke = Rf + Beta x MRP", "formula")
ke.cell(row=R_KE, column=1).font = Font(name=FONT, size=11, bold=True)
ke.cell(row=R_KE, column=2).font = Font(name=FONT, size=12, bold=True, color="C00000")
ke.cell(row=R_KE, column=2).fill = key_fill

r += 1
r = header(ke, r, "4. Pidilite's own realised return ('market rate') vs CAPM")
R_P5 = r
r = putk(r, "Pidilite realised total return, 5y", f"=(Data!E{last}/Data!E2)^(365/(Data!A{last}-Data!A2))-1", "0.00%",
        "From the Data sheet adjusted closes - what the stock actually delivered over the beta window.", "formula")
R_P10 = r
r = putk(r, "Pidilite realised total return, 10y", PIDI_10Y_TR, "0.00%",
        "10 years to 11-Sep-2026, from the full price history.", "input")
r = putk(r, "Realised 5y return less Ke", f"=B{R_P5}-B{R_KE}", "0.00%",
        "Negative means the stock under-delivered against its CAPM required return over the last 5 years.", "formula")

# Sensitivity grid
r += 1
r = header(ke, r, "5. Sensitivity of Ke to beta and market risk premium")
grid_top = r
ke.cell(row=grid_top, column=1, value="Beta \\ MRP").font = bold
mrps = [0.05, 0.06, 0.07, 0.08, 0.09]
for j, m in enumerate(mrps):
    c = ke.cell(row=grid_top, column=2 + j, value=m); c.font = bold; c.number_format = "0.0%"
    c.fill = PatternFill("solid", fgColor="D9E2F3"); c.border = box
betas = [0.55, 0.60, 0.65, 0.70, 0.75, 0.80, 0.85]
for i, b in enumerate(betas):
    rr = grid_top + 1 + i
    c = ke.cell(row=rr, column=1, value=b); c.font = bold; c.number_format = "0.00"
    c.fill = PatternFill("solid", fgColor="D9E2F3"); c.border = box
    for j in range(len(mrps)):
        col = get_column_letter(2 + j)
        cc = ke.cell(row=rr, column=2 + j, value=f"=$B${R_RF}+$A{rr}*{col}${grid_top}")
        cc.number_format = "0.00%"; cc.font = base; cc.border = box
ke.cell(row=grid_top + len(betas) + 2, column=1,
        value="Each cell is Rf + beta x MRP using the risk-free rate in B%d." % R_RF).font = small

ke.column_dimensions["A"].width = 38
for col in "BCDEFG":
    ke.column_dimensions[col].width = 13
ke.column_dimensions["H"].width = 100

# ---------------- Cost of Debt ----------------
kd = wb.create_sheet("Cost of Debt")
kd["A1"] = "Cost of Debt"; kd["A1"].font = title
kd["A2"] = "Kd  =  Total interest (finance costs)  /  Total borrowings        After-tax Kd  =  Kd x (1 - effective tax rate)"
kd["A2"].font = Font(name=FONT, size=10, italic=True)
kd["A3"] = "Pidilite Industries Ltd, consolidated, INR crore. Blue = reported figure (source on Notes); black = formula."
kd["A3"].font = small

hr = 5
kd.cell(row=hr, column=1, value="INR crore").font = hdr_font
kd.cell(row=hr, column=1).fill = hdr_fill
for j, f in enumerate(FY):
    c = kd.cell(row=hr, column=2 + j, value=f)
    c.font = hdr_font; c.fill = hdr_fill; c.alignment = Alignment(horizontal="center")

def line(row, label, values, fmt="#,##0.00", kind="input", note=None, indent=0):
    c1 = kd.cell(row=row, column=1, value=("    " * indent) + label); c1.font = base
    for j, v in enumerate(values):
        if v is None:
            continue
        c = kd.cell(row=row, column=2 + j, value=v)
        c.font = blue if kind == "input" else bold
        c.number_format = fmt; c.border = box
        if kind == "input":
            c.fill = key_fill
    if note:
        kd.cell(row=row, column=2 + len(FY) + 1, value=note).font = small
    return row + 1

r = hr + 1
R_INT = r
r = line(r, "Finance costs (total interest)", INTEREST, note="Consolidated P&L finance costs; includes interest on lease liabilities.")
R_BOR = r
r = line(r, "Total borrowings (closing)", BORROW, note="Balance-sheet borrowings as reported, i.e. including lease liabilities.")
R_BEX = r
r = line(r, "   of which: borrowings excl. leases", BORROW_EX, note="Pidilite is effectively debt-free; most of the balance is Ind AS 116 lease liability.", indent=1)
R_LSE = r
for j in range(len(FY)):
    col = get_column_letter(2 + j)
    c = kd.cell(row=R_LSE, column=2 + j, value=f"={col}{R_BOR}-{col}{R_BEX}")
    c.font = bold; c.number_format = "#,##0.00"; c.border = box
kd.cell(row=R_LSE, column=1, value="   of which: lease liabilities").font = base
kd.cell(row=R_LSE, column=2 + len(FY) + 1, value="Balancing figure.").font = small
r = R_LSE + 1

R_AVG = r
kd.cell(row=r, column=1, value="Average borrowings").font = base
for j in range(1, len(FY)):
    col, pcol = get_column_letter(2 + j), get_column_letter(1 + j)
    c = kd.cell(row=r, column=2 + j, value=f"=AVERAGE({pcol}{R_BOR},{col}{R_BOR})")
    c.font = bold; c.number_format = "#,##0.00"; c.border = box
kd.cell(row=r, column=2, value="n/a").font = small
kd.cell(row=r, column=2 + len(FY) + 1, value="FY2023 needs the FY2022 closing balance, which is outside the pulled series.").font = small
r += 1

r += 1
R_KDC = r
kd.cell(row=r, column=1, value="Cost of debt - closing borrowings").font = bold
for j in range(len(FY)):
    col = get_column_letter(2 + j)
    c = kd.cell(row=r, column=2 + j, value=f"={col}{R_INT}/{col}{R_BOR}")
    c.font = bold; c.number_format = "0.00%"; c.border = box
kd.cell(row=r, column=2 + len(FY) + 1, value="Total interest / total borrowings - the formula as specified.").font = small
r += 1

R_KDA = r
kd.cell(row=r, column=1, value="Cost of debt - average borrowings").font = base
for j in range(1, len(FY)):
    col = get_column_letter(2 + j)
    c = kd.cell(row=r, column=2 + j, value=f"={col}{R_INT}/{col}{R_AVG}")
    c.font = bold; c.number_format = "0.00%"; c.border = box
kd.cell(row=r, column=2 + len(FY) + 1, value="More accurate: interest accrues across the year, not on the closing balance.").font = small
r += 1

R_KDX = r
kd.cell(row=r, column=1, value="Memo: interest / borrowings excl. leases").font = base
for j in range(len(FY)):
    col = get_column_letter(2 + j)
    c = kd.cell(row=r, column=2 + j, value=f"={col}{R_INT}/{col}{R_BEX}")
    c.font = base; c.number_format = "0.00%"; c.border = box
kd.cell(row=r, column=2 + len(FY) + 1, value="Overstated - the numerator still contains lease interest. Shown only to expose the mismatch.").font = small
r += 2

R_PBT = r
r = line(r, "Profit before tax", PBT)
R_TAX = r
r = line(r, "Tax expense", TAX)
R_ETR = r
kd.cell(row=r, column=1, value="Effective tax rate").font = base
for j in range(len(FY)):
    col = get_column_letter(2 + j)
    c = kd.cell(row=r, column=2 + j, value=f"={col}{R_TAX}/{col}{R_PBT}")
    c.font = bold; c.number_format = "0.00%"; c.border = box
kd.cell(row=r, column=2 + len(FY) + 1, value="Tax expense / PBT.").font = small
r += 1

R_ATK = r
kd.cell(row=r, column=1, value="AFTER-TAX COST OF DEBT").font = Font(name=FONT, size=11, bold=True)
for j in range(len(FY)):
    col = get_column_letter(2 + j)
    c = kd.cell(row=r, column=2 + j, value=f"={col}{R_KDC}*(1-{col}{R_ETR})")
    c.font = Font(name=FONT, size=11, bold=True, color="C00000")
    c.number_format = "0.00%"; c.border = box; c.fill = key_fill
kd.cell(row=r, column=2 + len(FY) + 1, value="Kd x (1 - effective tax rate), on closing borrowings.").font = small
r += 2

kd.cell(row=r, column=1, value="Latest year (FY2026) headline").font = Font(name=FONT, size=11, bold=True, color="1F3864")
r += 1
lastcol = get_column_letter(1 + len(FY))
r = put(kd, r, "Cost of debt, pre-tax", f"={lastcol}{R_KDC}", "0.00%", "FY2026 finance costs / FY2026 total borrowings.", note_col=7)
r = put(kd, r, "Effective tax rate", f"={lastcol}{R_ETR}", "0.00%", None)
r = put(kd, r, "Cost of debt, after tax", f"={lastcol}{R_ATK}", "0.00%", "The figure that belongs in a WACC.", note_col=7)

kd.column_dimensions["A"].width = 38
for j in range(len(FY)):
    kd.column_dimensions[get_column_letter(2 + j)].width = 13
kd.column_dimensions[get_column_letter(2 + len(FY) + 1)].width = 95


# ---------------- WACC ----------------
# WACC = We x Ke + Wd x Kd x (1 - t).  Every driver is linked from the
# Cost of Equity / Cost of Debt sheets so nothing is retyped.
SHARES = 1017766288        # ordinary shares outstanding at 31-Mar-2026
CASH_EQ = 232.49           # cash and cash equivalents, INR crore, 31-Mar-2026
CASH_STI = 4215.98         # cash + short-term investments, INR crore, 31-Mar-2026

wc = wb.create_sheet("WACC")
wc["A1"] = "Weighted Average Cost of Capital"; wc["A1"].font = title
wc["A2"] = "WACC  =  We x Ke  +  Wd x Kd x (1 - tax rate)"
wc["A2"].font = Font(name=FONT, size=10, italic=True)
wc["A3"] = "Blue = hardcoded input (sourced on Notes)   Green = linked from another sheet   Black = formula"
wc["A3"].font = small

def putw(row, label, value, fmt=None, note=None, kind="formula", indent=0):
    return put(wc, row, label, value, fmt, note, kind, indent, note_col=8)

r = 5
r = header(wc, r, "1. Market value of equity")
W_SH = r
r = putw(r, "Shares outstanding", SHARES, "#,##0",
         "Ordinary shares at 31-Mar-2026. Reconciles to the Rs 102 crore equity capital at Re 1 face value.", "input")
W_PX = r
r = putw(r, "Share price (INR)", f"=Data!D{last}", "#,##0.00",
         "Latest unadjusted close on the Data sheet. Unadjusted is correct here - market cap needs the traded price.", "link")
W_E = r
r = putw(r, "Market value of equity (INR crore)", f"=B{W_SH}*B{W_PX}/10000000", "#,##0.00",
         "Shares x price, converted to crore (1 crore = 10,000,000).")

r += 1
r = header(wc, r, "2. Market value of debt")
W_D = r
r = putw(r, "Total borrowings (INR crore)", f"='Cost of Debt'!{lastcol}{R_BOR}", "#,##0.00",
         "FY2026 closing borrowings, linked from the Cost of Debt sheet. Book value is the standard proxy for the "
         "market value of debt: the borrowings are short-dated and floating, so book and market value are close.", "link")

r += 1
r = header(wc, r, "3. Capital structure weights")
W_V = r
r = putw(r, "Total capital (D + E)", f"=B{W_E}+B{W_D}", "#,##0.00")
W_WE = r
r = putw(r, "Weight of equity (We)", f"=B{W_E}/B{W_V}", "0.00%")
W_WD = r
r = putw(r, "Weight of debt (Wd)", f"=B{W_D}/B{W_V}", "0.00%",
         "Under 1%. Pidilite is financed almost entirely by equity, so WACC lands within a few basis points of Ke.")

r += 1
r = header(wc, r, "4. Component costs")
W_KE = r
r = putw(r, "Cost of equity (Ke)", f"='Cost of Equity'!B{R_KE}", "0.00%",
         "Linked from the Cost of Equity sheet - CAPM on the regression beta.", "link")
W_KDP = r
r = putw(r, "Cost of debt, pre-tax (Kd)", f"='Cost of Debt'!{lastcol}{R_KDC}", "0.00%",
         "FY2026 total interest / total borrowings, linked from the Cost of Debt sheet.", "link")
W_T = r
r = putw(r, "Effective tax rate", f"='Cost of Debt'!{lastcol}{R_ETR}", "0.00%", None, "link")
W_KDA = r
r = putw(r, "Cost of debt, after tax", f"=B{W_KDP}*(1-B{W_T})", "0.00%", "Interest is tax-deductible, so only the after-tax cost is borne.")

r += 1
r = header(wc, r, "5. WACC")
W_CE = r
r = putw(r, "Equity contribution", f"=B{W_WE}*B{W_KE}", "0.00%")
W_CD = r
r = putw(r, "Debt contribution", f"=B{W_WD}*B{W_KDA}", "0.00%")
W_WACC = r
r = putw(r, "WACC", f"=B{W_CE}+B{W_CD}", "0.00%", "We x Ke + Wd x after-tax Kd.")
wc.cell(row=W_WACC, column=1).font = Font(name=FONT, size=11, bold=True)
wc.cell(row=W_WACC, column=2).font = Font(name=FONT, size=12, bold=True, color="C00000")
wc.cell(row=W_WACC, column=2).fill = key_fill
W_GAP = r
r = putw(r, "WACC less Ke", f"=B{W_WACC}-B{W_KE}", "0.00%",
         "A few basis points. With a sub-1% debt weight the capital structure barely moves the number.")

r += 1
r = header(wc, r, "6. Memo - net cash position")
W_C1 = r
r = putw(r, "Cash and cash equivalents", CASH_EQ, "#,##0.00", "31-Mar-2026 consolidated balance sheet.", "input")
W_C2 = r
r = putw(r, "Cash + short-term investments", CASH_STI, "#,##0.00", "Includes liquid investments held as treasury.", "input")
W_ND = r
r = putw(r, "Net debt (borrowings - cash & investments)", f"=B{W_D}-B{W_C2}", "#,##0.00",
         "NEGATIVE - Pidilite holds far more cash and liquid investments than borrowings, so it is net cash. "
         "A net-debt-weighted WACC is not meaningful here; the gross-debt weighting above is used.")

# Sensitivity: WACC under a target capital structure
r += 1
r = header(wc, r, "7. Sensitivity - WACC at a target debt weight")
wc.cell(row=r, column=1,
        value="Actual gearing is under 1%, so this shows what WACC would be if Pidilite levered up, holding "
              "Ke and after-tax Kd at their current values.").font = small
r += 1
g = r
wc.cell(row=g, column=1, value="Ke \\ Wd").font = bold
wds = [0.00, 0.05, 0.10, 0.20, 0.30]
for j, w in enumerate(wds):
    c = wc.cell(row=g, column=2 + j, value=w); c.font = bold; c.number_format = "0%"
    c.fill = PatternFill("solid", fgColor="D9E2F3"); c.border = box
kes = [0.10, 0.11, 0.12, 0.13, 0.14]
for i, k in enumerate(kes):
    rr = g + 1 + i
    c = wc.cell(row=rr, column=1, value=k); c.font = bold; c.number_format = "0.0%"
    c.fill = PatternFill("solid", fgColor="D9E2F3"); c.border = box
    for j in range(len(wds)):
        col = get_column_letter(2 + j)
        cc = wc.cell(row=rr, column=2 + j, value=f"=(1-{col}${g})*$A{rr}+{col}${g}*$B${W_KDA}")
        cc.number_format = "0.00%"; cc.font = base; cc.border = box
wc.cell(row=g + len(kes) + 2, column=1,
        value="Each cell = (1 - Wd) x Ke + Wd x after-tax Kd, using the after-tax Kd in B%d." % W_KDA).font = small

wc.column_dimensions["A"].width = 38
for col in "BCDEFG":
    wc.column_dimensions[col].width = 15
wc.column_dimensions["H"].width = 100


# ---------------- DCF ----------------
# Historical actuals (INR crore), Pidilite consolidated, FY2023-FY2026.
# Same Yahoo/screener.in source as the Cost of Debt sheet.
H_YRS  = ["FY2023", "FY2024", "FY2025", "FY2026"]
H_REV  = [11751.62, 12337.07, 13092.76, 14551.26]
H_EBIT = [1752.01, 2402.73, 2673.41, 3147.99]     # operating income, i.e. EXCLUDING other income
H_DA   = [269.74, 320.30, 350.45, 388.99]
H_CAP  = [505.35, 558.71, 452.34, 592.77]
H_NWC  = [1459.14, 768.97, 824.87, 836.44]        # current assets less cash & investments,
                                                  # less current liabilities excluding current debt
P_YRS  = ["FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E"]
P_GROW = [0.10, 0.09, 0.08, 0.07, 0.06]           # fading toward the terminal rate
A_MARGIN, A_DA, A_CAPEX, A_NWC, A_TG = 0.2163, 0.0265, 0.0409, 0.0600, 0.0500

dc = wb.create_sheet("DCF")
dc["A1"] = "Cash Flow Projections and DCF Valuation"; dc["A1"].font = title
dc["A2"] = "FCFF  =  EBIT x (1 - tax)  +  D&A  -  Capex  -  Increase in net working capital"
dc["A2"].font = Font(name=FONT, size=10, italic=True)
dc["A3"] = "Blue = hardcoded input   Green = linked from another sheet   Black = formula.  All figures INR crore unless stated."
dc["A3"].font = small

def putd(row, label, value, fmt=None, note=None, kind="formula", indent=0):
    return put(dc, row, label, value, fmt, note, kind, indent, note_col=8)

def band(row, labels, first_col=2):
    for j, t in enumerate(labels):
        c = dc.cell(row=row, column=first_col + j, value=t)
        c.font = hdr_font; c.fill = hdr_fill; c.alignment = Alignment(horizontal="center")
    return row + 1

def series(row, label, values, fmt="#,##0.00", kind="input", indent=0, note=None):
    dc.cell(row=row, column=1, value=("    " * indent) + label).font = base
    for j, v in enumerate(values):
        c = dc.cell(row=row, column=2 + j, value=v)
        c.font = {"input": blue, "formula": bold, "link": green}[kind]
        c.number_format = fmt; c.border = box
        if kind == "input":
            c.fill = key_fill
    if note:
        dc.cell(row=row, column=8, value=note).font = small
    return row + 1

r = 5
r = header(dc, r, "1. Historical performance (reported actuals)")
D_HH = r; r = band(r, H_YRS)
D_HREV = r;  r = series(r, "Revenue", H_REV)
D_HEB  = r;  r = series(r, "EBIT (operating income)", H_EBIT, note="Operating income, which EXCLUDES other income - treasury income is captured separately via net cash.")
D_HEBM = r
dc.cell(row=r, column=1, value="    EBIT margin").font = base
for j in range(4):
    c = get_column_letter(2 + j)
    cc = dc.cell(row=r, column=2 + j, value=f"={c}{D_HEB}/{c}{D_HREV}"); cc.font = bold; cc.number_format = "0.00%"; cc.border = box
r += 1
D_HDA = r;  r = series(r, "Depreciation and amortisation", H_DA)
D_HDAP = r
dc.cell(row=r, column=1, value="    D&A % of revenue").font = base
for j in range(4):
    c = get_column_letter(2 + j)
    cc = dc.cell(row=r, column=2 + j, value=f"={c}{D_HDA}/{c}{D_HREV}"); cc.font = bold; cc.number_format = "0.00%"; cc.border = box
r += 1
D_HCAP = r; r = series(r, "Capital expenditure", H_CAP)
D_HCAPP = r
dc.cell(row=r, column=1, value="    Capex % of revenue").font = base
for j in range(4):
    c = get_column_letter(2 + j)
    cc = dc.cell(row=r, column=2 + j, value=f"={c}{D_HCAP}/{c}{D_HREV}"); cc.font = bold; cc.number_format = "0.00%"; cc.border = box
r += 1
D_HNWC = r; r = series(r, "Net working capital", H_NWC, note="Current assets less cash and investments, less current liabilities excluding current debt.")
D_HNWCP = r
dc.cell(row=r, column=1, value="    NWC % of revenue").font = base
for j in range(4):
    c = get_column_letter(2 + j)
    cc = dc.cell(row=r, column=2 + j, value=f"={c}{D_HNWC}/{c}{D_HREV}"); cc.font = bold; cc.number_format = "0.00%"; cc.border = box
r += 1
D_HG = r
dc.cell(row=r, column=1, value="    Revenue growth").font = base
for j in range(1, 4):
    c, p = get_column_letter(2 + j), get_column_letter(1 + j)
    cc = dc.cell(row=r, column=2 + j, value=f"={c}{D_HREV}/{p}{D_HREV}-1"); cc.font = bold; cc.number_format = "0.00%"; cc.border = box
dc.cell(row=r, column=8, value="FY2024-FY2026 growth of 5.0%, 6.1% and 11.1%; three-year revenue CAGR 7.4%.").font = small
r += 2

r = header(dc, r, "2. Assumptions")
D_MARG = r; r = putd(r, "EBIT margin (held flat)", A_MARGIN, "0.00%",
    "FY2026 actual. Margin rose from 14.9% to 21.6% over four years on softer input costs; holding it flat rather than extrapolating.", "input")
D_DA = r;   r = putd(r, "D&A % of revenue", A_DA, "0.00%", "Three-year average (FY2024-FY2026).", "input")
D_CAP = r;  r = putd(r, "Capex % of revenue", A_CAPEX, "0.00%", "Four-year average. Above D&A, consistent with a growing asset base.", "input")
D_NWC = r;  r = putd(r, "NWC % of incremental revenue", A_NWC, "0.00%",
    "Three-year average NWC intensity. Applied to the CHANGE in revenue, so no step change in the NWC level.", "input")
D_TAX = r;  r = putd(r, "Effective tax rate", f"='Cost of Debt'!{lastcol}{R_ETR}", "0.00%", "Linked from the Cost of Debt sheet.", "link")
D_W = r;    r = putd(r, "WACC (discount rate)", f"=WACC!B{W_WACC}", "0.00%", "Linked from the WACC sheet.", "link")
D_TG = r;   r = putd(r, "Terminal growth rate", A_TG, "0.00%",
    "Perpetual nominal growth. Must stay below WACC, and below the 6.97% risk-free rate - no firm outgrows its economy forever.", "input")
r += 1

r = header(dc, r, "3. Free cash flow projection")
D_PH = r; r = band(r, P_YRS)
D_PG = r; r = series(r, "Revenue growth", P_GROW, "0.00%", note="Fades from 10% toward the 5% terminal rate.")
D_PREV = r
dc.cell(row=r, column=1, value="Revenue").font = base
for j in range(5):
    c = get_column_letter(2 + j)
    prev = f"{get_column_letter(1+len(H_YRS))}{D_HREV}" if j == 0 else f"{get_column_letter(1+j)}{D_PREV}"
    cc = dc.cell(row=r, column=2 + j, value=f"={prev}*(1+{c}{D_PG})"); cc.font = bold; cc.number_format = "#,##0.00"; cc.border = box
dc.cell(row=r, column=8, value="FY2027 grows off the FY2026 actual.").font = small
r += 1
def calc(row, label, tmpl, fmt="#,##0.00", boldit=True, note=None, indent=0):
    dc.cell(row=row, column=1, value=("    " * indent) + label).font = base
    for j in range(5):
        c = get_column_letter(2 + j)
        cc = dc.cell(row=row, column=2 + j, value=tmpl.format(c=c, j=j))
        cc.font = bold if boldit else base; cc.number_format = fmt; cc.border = box
    if note:
        dc.cell(row=row, column=8, value=note).font = small
    return row + 1
D_PEB = r;  r = calc(r, "EBIT", f"={{c}}{D_PREV}*$B${D_MARG}")
D_PTX = r;  r = calc(r, "Less: tax on EBIT", f"=-{{c}}{D_PEB}*$B${D_TAX}")
D_PNO = r;  r = calc(r, "NOPAT", f"={{c}}{D_PEB}+{{c}}{D_PTX}")
D_PDA = r;  r = calc(r, "Add: D&A", f"={{c}}{D_PREV}*$B${D_DA}")
D_PCX = r;  r = calc(r, "Less: capex", f"=-{{c}}{D_PREV}*$B${D_CAP}")
D_PWC = r
dc.cell(row=r, column=1, value="Less: increase in NWC").font = base
for j in range(5):
    c = get_column_letter(2 + j)
    prev = f"{get_column_letter(1+len(H_YRS))}{D_HREV}" if j == 0 else f"{get_column_letter(1+j)}{D_PREV}"
    cc = dc.cell(row=r, column=2 + j, value=f"=-({c}{D_PREV}-{prev})*$B${D_NWC}")
    cc.font = bold; cc.number_format = "#,##0.00"; cc.border = box
dc.cell(row=r, column=8, value="NWC intensity applied to the increase in revenue.").font = small
r += 1
D_FCF = r; r = calc(r, "FREE CASH FLOW TO FIRM (FCFF)", f"={{c}}{D_PNO}+{{c}}{D_PDA}+{{c}}{D_PCX}+{{c}}{D_PWC}")
for j in range(5):
    dc.cell(row=D_FCF, column=2 + j).fill = key_fill
dc.cell(row=D_FCF, column=1).font = Font(name=FONT, size=10, bold=True)
D_DP = r;  r = series(r, "Discount period (years)", [1, 2, 3, 4, 5], "0", note="End-of-period discounting from the 31-Mar-2026 balance sheet date.")
D_DF = r;  r = calc(r, "Discount factor", f"=1/(1+$B${D_W})^{{c}}{D_DP}", "0.0000")
D_PV = r;  r = calc(r, "PV of FCFF", f"={{c}}{D_FCF}*{{c}}{D_DF}")
r += 1

r = header(dc, r, "4. Valuation")
fc, lc = "B", get_column_letter(1 + 5)
D_SPV = r; r = putd(r, "Sum of PV of forecast FCFF", f"=SUM({fc}{D_PV}:{lc}{D_PV})", "#,##0.00")
D_TV = r;  r = putd(r, "Terminal value (Gordon growth)", f"={lc}{D_FCF}*(1+$B${D_TG})/($B${D_W}-$B${D_TG})", "#,##0.00",
    "FY2031 FCFF grown one year, capitalised at WACC less terminal growth.")
D_PTV = r; r = putd(r, "PV of terminal value", f"=B{D_TV}*{lc}{D_DF}", "#,##0.00")
D_EV = r;  r = putd(r, "Enterprise value", f"=B{D_SPV}+B{D_PTV}", "#,##0.00")
D_TVP = r; r = putd(r, "    Terminal value as % of EV", f"=B{D_PTV}/B{D_EV}", "0.00%",
    "Above ~75% means the answer rests mainly on the terminal assumptions rather than the forecast.")
D_ND = r;  r = putd(r, "Less: net debt", f"=WACC!B{W_ND}", "#,##0.00",
    "Negative - Pidilite is net cash, so this ADDS to equity value.", "link")
D_EQ = r;  r = putd(r, "Equity value", f"=B{D_EV}-B{D_ND}", "#,##0.00")
D_SHR = r; r = putd(r, "Shares outstanding (crore)", f"=WACC!B{W_SH}/10000000", "#,##0.0000", None, "link")
D_VPS = r; r = putd(r, "INTRINSIC VALUE PER SHARE (INR)", f"=B{D_EQ}/B{D_SHR}", "#,##0.00")
dc.cell(row=D_VPS, column=1).font = Font(name=FONT, size=11, bold=True)
dc.cell(row=D_VPS, column=2).font = Font(name=FONT, size=12, bold=True, color="C00000")
dc.cell(row=D_VPS, column=2).fill = key_fill
D_MP = r;  r = putd(r, "Current market price (INR)", f"=WACC!B{W_PX}", "#,##0.00", None, "link")
D_UP = r;  r = putd(r, "Upside / (downside)", f"=B{D_VPS}/B{D_MP}-1", "0.00%",
    "Large negative: Pidilite trades on a premium multiple that this set of assumptions does not support. See section 6.")
r += 1

r = header(dc, r, "5. Sensitivity - value per share vs WACC and terminal growth")
gs = r
dc.cell(row=gs, column=1, value="WACC \\ g").font = bold
tgs = [0.04, 0.045, 0.05, 0.055, 0.06]
for j, t in enumerate(tgs):
    c = dc.cell(row=gs, column=2 + j, value=t); c.font = bold; c.number_format = "0.0%"
    c.fill = PatternFill("solid", fgColor="D9E2F3"); c.border = box
waccs = [0.1085, 0.1135, 0.1185, 0.1235, 0.1285]
for i, w in enumerate(waccs):
    rr = gs + 1 + i
    c = dc.cell(row=rr, column=1, value=w); c.font = bold; c.number_format = "0.00%"
    c.fill = PatternFill("solid", fgColor="D9E2F3"); c.border = box
    for j in range(len(tgs)):
        col = get_column_letter(2 + j)
        f = (f"=(SUMPRODUCT($B${D_FCF}:${lc}${D_FCF}/(1+$A{rr})^$B${D_DP}:${lc}${D_DP})"
             f"+${lc}${D_FCF}*(1+{col}${gs})/($A{rr}-{col}${gs})/(1+$A{rr})^${lc}${D_DP}"
             f"-$B${D_ND})/$B${D_SHR}")
        cc = dc.cell(row=rr, column=2 + j, value=f)
        cc.number_format = "#,##0"; cc.font = base; cc.border = box
r = gs + len(waccs) + 1
dc.cell(row=r, column=1, value="Each cell re-runs the whole DCF at that WACC and terminal growth rate.").font = small
r += 2

r = header(dc, r, "6. Reverse DCF - what the market price implies")
D_MEV = r; r = putd(r, "Enterprise value at market price", f"=WACC!B{W_E}+WACC!B{W_ND}", "#,##0.00",
    "Market capitalisation plus net debt (which is negative here).", "link")
D_K = r;   r = putd(r, "Terminal value the market implies", f"=(B{D_MEV}-B{D_SPV})/{lc}{D_DF}", "#,##0.00",
    "Market EV less the PV of the forecast FCFF, un-discounted back to FY2031.")
D_IG = r;  r = putd(r, "Implied perpetual growth rate", f"=(B{D_K}*$B${D_W}-{lc}{D_FCF})/(B{D_K}+{lc}{D_FCF})", "0.00%",
    "Solving Gordon growth backwards. Compare with the 6.97% risk-free rate: a perpetual rate above it implies the "
    "company eventually outgrows the whole economy, so the market is pricing in either much stronger growth, higher "
    "margins, or a lower risk premium than this model assumes.")
r += 1
dc.cell(row=r, column=1, value="Read this section before treating the downside above as a recommendation - it shows which "
                               "assumption the gap actually lives in.").font = small

dc.column_dimensions["A"].width = 36
for col in "BCDEFG":
    dc.column_dimensions[col].width = 14
dc.column_dimensions["H"].width = 100


# ---------------- DCF - 10 year scenario ----------------
# Deliberately identical to the 5-year base case in EVERY per-revenue ratio.
# The only things that change are the length of the explicit window and the
# growth path, so the comparison isolates the effect of the longer runway.
X_YRS  = ["FY%dE" % y for y in range(2027, 2037)]
X_GROW = [0.110, 0.104, 0.098, 0.092, 0.086, 0.080, 0.074, 0.068, 0.062, 0.056]
X_MS, X_ME = 0.2163, 0.2163      # EBIT margin in the first and last forecast year
N10 = len(X_YRS)
XC  = [get_column_letter(2 + j) for j in range(N10)]   # B .. K
XL  = XC[-1]
NOTE10 = 13                                            # notes live in column M

x = wb.create_sheet("DCF 10Y")
x["A1"] = "DCF - 10 Year Forecast Scenario"; x["A1"].font = title
x["A2"] = "Same business assumptions as the 5-year base case. Only the forecast window and the growth path differ."
x["A2"].font = Font(name=FONT, size=10, italic=True)
x["A3"] = "Blue = input   Green = linked from another sheet   Black = formula.  All figures INR crore unless stated."
x["A3"].font = small

def putx(row, label, value, fmt=None, note=None, kind="formula", indent=0):
    return put(x, row, label, value, fmt, note, kind, indent, note_col=NOTE10)

def calcx(row, label, tmpl, fmt="#,##0.00", note=None, boldit=True, indent=0):
    x.cell(row=row, column=1, value=("    " * indent) + label).font = base
    for j in range(N10):
        cc = x.cell(row=row, column=2 + j, value=tmpl.format(c=XC[j], j=j))
        cc.font = bold if boldit else base; cc.number_format = fmt; cc.border = box
    if note:
        x.cell(row=row, column=NOTE10, value=note).font = small
    return row + 1

r = 5
r = header(x, r, "1. What this scenario changes")
r = putx(r, "Forecast window", None, None, "Ten explicit years (FY2027-FY2036) instead of five, then the same terminal value.")
r = putx(r, "Growth path", None, None,
         "Fades linearly from 11.0% to 5.6%, roughly 0.6pp a year, landing beside the 5.0% terminal rate. The 5-year "
         "case had to cut from 10% to 6% in half the time, which forces the fade through faster than a franchise of "
         "this quality plausibly decays.")
r = putx(r, "Everything else", None, None,
         "Margin, D&A, capex, working capital intensity, tax, WACC and terminal growth are all unchanged and, where "
         "possible, linked directly to the base-case DCF sheet - so any difference in value comes from the runway alone.")
r += 1

r = header(x, r, "2. Assumptions")
X_MS_R = r; r = putx(r, "EBIT margin, first forecast year", X_MS, "0.00%", "FY2026 actual, as in the base case.", "input")
X_ME_R = r; r = putx(r, "EBIT margin, final forecast year", X_ME, "0.00%",
    "Defaults EQUAL to the first year, so margin is held flat and this scenario is a like-for-like test of the runway. "
    "Raise it to model operating leverage; the margin is interpolated linearly between the two.", "input")
X_DA = r;   r = putx(r, "D&A % of revenue", f"=DCF!B{D_DA}", "0.00%", None, "link")
X_CAP = r;  r = putx(r, "Capex % of revenue", f"=DCF!B{D_CAP}", "0.00%", None, "link")
X_NWC = r;  r = putx(r, "NWC % of incremental revenue", f"=DCF!B{D_NWC}", "0.00%", None, "link")
X_TAX = r;  r = putx(r, "Effective tax rate", f"=DCF!B{D_TAX}", "0.00%", None, "link")
X_W = r;    r = putx(r, "WACC (discount rate)", f"=DCF!B{D_W}", "0.00%", None, "link")
X_TG = r;   r = putx(r, "Terminal growth rate", f"=DCF!B{D_TG}", "0.00%", None, "link")
r += 1

r = header(x, r, "3. Free cash flow projection - ten years")
X_H = r
for j, t in enumerate(X_YRS):
    c = x.cell(row=r, column=2 + j, value=t)
    c.font = hdr_font; c.fill = hdr_fill; c.alignment = Alignment(horizontal="center")
r += 1
X_G = r
x.cell(row=r, column=1, value="Revenue growth").font = base
for j, g in enumerate(X_GROW):
    c = x.cell(row=r, column=2 + j, value=g)
    c.font = blue; c.number_format = "0.00%"; c.border = box; c.fill = key_fill
x.cell(row=r, column=NOTE10, value="Linear fade of about 0.6pp a year from FY2026's actual 11.1%.").font = small
r += 1
X_REV = r
x.cell(row=r, column=1, value="Revenue").font = base
for j in range(N10):
    prev = f"DCF!{get_column_letter(1+4)}{D_HREV}" if j == 0 else f"{XC[j-1]}{X_REV}"
    c = x.cell(row=r, column=2 + j, value=f"={prev}*(1+{XC[j]}{X_G})")
    c.font = bold; c.number_format = "#,##0.00"; c.border = box
x.cell(row=r, column=NOTE10, value="FY2027 grows off the FY2026 actual on the DCF sheet.").font = small
r += 1
X_M = r
r = calcx(r, "EBIT margin", f"=$B${X_MS_R}+($B${X_ME_R}-$B${X_MS_R})*{{j}}/{N10-1}", "0.00%",
          "Linear interpolation between the two margin inputs. Flat by default.")
X_EB = r;  r = calcx(r, "EBIT", f"={{c}}{X_REV}*{{c}}{X_M}")
X_TX = r;  r = calcx(r, "Less: tax on EBIT", f"=-{{c}}{X_EB}*$B${X_TAX}")
X_NO = r;  r = calcx(r, "NOPAT", f"={{c}}{X_EB}+{{c}}{X_TX}")
X_DAR = r; r = calcx(r, "Add: D&A", f"={{c}}{X_REV}*$B${X_DA}")
X_CX = r;  r = calcx(r, "Less: capex", f"=-{{c}}{X_REV}*$B${X_CAP}")
X_WC = r
x.cell(row=r, column=1, value="Less: increase in NWC").font = base
for j in range(N10):
    prev = f"DCF!{get_column_letter(1+4)}{D_HREV}" if j == 0 else f"{XC[j-1]}{X_REV}"
    c = x.cell(row=r, column=2 + j, value=f"=-({XC[j]}{X_REV}-{prev})*$B${X_NWC}")
    c.font = bold; c.number_format = "#,##0.00"; c.border = box
r += 1
X_F = r; r = calcx(r, "FREE CASH FLOW TO FIRM (FCFF)", f"={{c}}{X_NO}+{{c}}{X_DAR}+{{c}}{X_CX}+{{c}}{X_WC}")
for j in range(N10):
    x.cell(row=X_F, column=2 + j).fill = key_fill
x.cell(row=X_F, column=1).font = Font(name=FONT, size=10, bold=True)
X_DP = r
x.cell(row=r, column=1, value="Discount period (years)").font = base
for j in range(N10):
    c = x.cell(row=r, column=2 + j, value=j + 1)
    c.font = blue; c.number_format = "0"; c.border = box; c.fill = key_fill
r += 1
X_DF = r; r = calcx(r, "Discount factor", f"=1/(1+$B${X_W})^{{c}}{X_DP}", "0.0000")
X_PV = r; r = calcx(r, "PV of FCFF", f"={{c}}{X_F}*{{c}}{X_DF}")
r += 1

r = header(x, r, "4. Valuation")
X_SPV = r; r = putx(r, "Sum of PV of forecast FCFF", f"=SUM(B{X_PV}:{XL}{X_PV})", "#,##0.00")
X_TV = r;  r = putx(r, "Terminal value (Gordon growth)", f"={XL}{X_F}*(1+$B${X_TG})/($B${X_W}-$B${X_TG})", "#,##0.00")
X_PTV = r; r = putx(r, "PV of terminal value", f"=B{X_TV}*{XL}{X_DF}", "#,##0.00")
X_EV = r;  r = putx(r, "Enterprise value", f"=B{X_SPV}+B{X_PTV}", "#,##0.00")
X_TVP = r; r = putx(r, "    Terminal value as % of EV", f"=B{X_PTV}/B{X_EV}", "0.00%",
    "Lower than the 5-year case: a longer explicit window shifts weight out of the terminal assumption and into the forecast.")
X_ND = r;  r = putx(r, "Less: net debt", f"=DCF!B{D_ND}", "#,##0.00", "Negative - net cash, so this adds.", "link")
X_EQ = r;  r = putx(r, "Equity value", f"=B{X_EV}-B{X_ND}", "#,##0.00")
X_SH = r;  r = putx(r, "Shares outstanding (crore)", f"=DCF!B{D_SHR}", "#,##0.0000", None, "link")
X_VPS = r; r = putx(r, "INTRINSIC VALUE PER SHARE (INR)", f"=B{X_EQ}/B{X_SH}", "#,##0.00")
x.cell(row=X_VPS, column=1).font = Font(name=FONT, size=11, bold=True)
x.cell(row=X_VPS, column=2).font = Font(name=FONT, size=12, bold=True, color="C00000")
x.cell(row=X_VPS, column=2).fill = key_fill
r += 1

r = header(x, r, "5. Scenario vs base case")
X_B5 = r; r = putx(r, "5-year base case, value per share", f"=DCF!B{D_VPS}", "#,##0.00", None, "link")
X_B10 = r; r = putx(r, "10-year scenario, value per share", f"=B{X_VPS}", "#,##0.00")
X_DIF = r; r = putx(r, "Uplift from the longer runway", f"=B{X_B10}/B{X_B5}-1", "0.00%",
    "The whole effect of doubling the explicit forecast, with every other assumption held identical.")
X_MP = r; r = putx(r, "Current market price (INR)", f"=DCF!B{D_MP}", "#,##0.00", None, "link")
X_U5 = r; r = putx(r, "    Upside on base case", f"=B{X_B5}/B{X_MP}-1", "0.00%")
X_U10 = r; r = putx(r, "    Upside on this scenario", f"=B{X_B10}/B{X_MP}-1", "0.00%",
    "Still deeply negative. Doubling the runway does not come close to bridging the gap, which is the point of the scenario.")
r += 1

r = header(x, r, "6. Sensitivity - value per share vs WACC and terminal growth")
gs = r
x.cell(row=gs, column=1, value="WACC \\ g").font = bold
tgs = [0.04, 0.045, 0.05, 0.055, 0.06]
for j, t in enumerate(tgs):
    c = x.cell(row=gs, column=2 + j, value=t); c.font = bold; c.number_format = "0.0%"
    c.fill = PatternFill("solid", fgColor="D9E2F3"); c.border = box
waccs = [0.1085, 0.1135, 0.1185, 0.1235, 0.1285]
for i, w in enumerate(waccs):
    rr = gs + 1 + i
    c = x.cell(row=rr, column=1, value=w); c.font = bold; c.number_format = "0.00%"
    c.fill = PatternFill("solid", fgColor="D9E2F3"); c.border = box
    for j in range(len(tgs)):
        col = get_column_letter(2 + j)
        f = (f"=(SUMPRODUCT($B${X_F}:${XL}${X_F}/(1+$A{rr})^$B${X_DP}:${XL}${X_DP})"
             f"+${XL}${X_F}*(1+{col}${gs})/($A{rr}-{col}${gs})/(1+$A{rr})^${XL}${X_DP}"
             f"-$B${X_ND})/$B${X_SH}")
        cc = x.cell(row=rr, column=2 + j, value=f)
        cc.number_format = "#,##0"; cc.font = base; cc.border = box
r = gs + len(waccs) + 1
x.cell(row=r, column=1, value="Each cell re-runs the full ten-year DCF at that WACC and terminal growth rate.").font = small
r += 2

r = header(x, r, "7. Reverse DCF on the ten-year forecast")
X_MEV = r; r = putx(r, "Enterprise value at market price", f"=WACC!B{W_E}+WACC!B{W_ND}", "#,##0.00", None, "link")
X_K = r;   r = putx(r, "Terminal value the market implies", f"=(B{X_MEV}-B{X_SPV})/{XL}{X_DF}", "#,##0.00")
X_IG = r;  r = putx(r, "Implied perpetual growth rate", f"=(B{X_K}*$B${X_W}-{XL}{X_F})/(B{X_K}+{XL}{X_F})", "0.00%",
    "Even after ten years of above-terminal growth, the rate needed to justify today's price stays above the 6.97% "
    "risk-free rate. That is the useful conclusion: the gap is not an artefact of too short a forecast window.")

x.column_dimensions["A"].width = 36
for j in range(N10 + 1):
    x.column_dimensions[get_column_letter(2 + j)].width = 12
x.column_dimensions[get_column_letter(NOTE10)].width = 100

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
    ("", ""),
    ("Cost of equity", ""),
    ("Risk-free rate", "6.97% - India 10-year benchmark G-Sec yield as at 10-Sep-2026, Trading Economics "
                       "(https://tradingeconomics.com/india/government-bond-yield). Cross-checked at 6.96% on 04-Sep-2026. "
                       "A 10Y sovereign yield is the conventional Rf for an INR-denominated valuation."),
    ("Beta", "Linked live to Regression!B4 - the OLS slope computed on this workbook's own daily returns. Not retyped."),
    ("Market return", "^NSEI is a PRICE index and excludes dividends, so a dividend yield of 1.25% (NIFTY 50 long-run average) "
                      "is added to every price CAGR to reach a total market return."),
    ("Why three approaches", "The trailing 5-year window opens close to the Sep-2021 market peak, so the realised NIFTY 50 total "
                             "return lands within a few tenths of a percent of the risk-free rate. The resulting risk premium is "
                             "approximately zero - and on a marginally different start date it goes negative, putting the cost of "
                             "equity below Rf. Either way it is an artefact of the start date, not an economic result, and cannot "
                             "be used on its own. "
                             "Approaches B (10-year realised) and C (forward-looking implied premium) are shown alongside it, "
                             "and the selected-Ke cell defaults to C."),
    ("India ERP", "7.08% - Damodaran implied India equity risk premium, January 2026 vintage "
                  "(https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/ctryprem.html), being the mature-market ERP "
                  "plus a 2.85% India country risk premium at a Baa3 rating. A July-2026 update quotes 7.31%."),
    ("10-year CAGRs", "NIFTY 50 price CAGR 10.49% and Pidilite total-return CAGR 16.46%, both over the 10 years to 11-Sep-2026, "
                      "computed from the full Yahoo Finance price history (a longer window than this workbook's 5-year Data sheet)."),
    ("", ""),
    ("Cost of debt", ""),
    ("Source", "Pidilite Industries Ltd CONSOLIDATED annual financials, FY2023-FY2026 (years ended 31 March). Retrieved from the "
               "Yahoo Finance fundamentals-timeseries API and cross-checked line by line against screener.in "
               "(https://www.screener.in/company/PIDILITIND/consolidated/). Both restate the same statements the company files "
               "with the exchanges; every figure used agreed between the two sources."),
    ("NSE not reachable", "www.nseindia.com returns HTTP 403 to this environment (NSE blocks datacentre traffic), so the figures "
                          "could not be pulled from NSE directly. The two sources above carry the same filed consolidated numbers."),
    ("Borrowings definition", "The 'Total borrowings' line is the borrowings figure as reported, which under Ind AS 116 INCLUDES "
                              "lease liabilities - and finance costs likewise include lease interest, so numerator and denominator "
                              "are consistent. Pidilite is close to debt-free: of FY2026 borrowings of Rs 417.21 crore, only "
                              "Rs 105.91 crore is actual borrowing and the rest is lease liability. Dividing total interest by "
                              "borrowings-excluding-leases would mismatch the two, and is shown on the sheet only as a memo."),
    ("Closing vs average", "The headline follows the requested formula (total interest / total borrowings) on closing balances. "
                           "A cost of debt on AVERAGE borrowings is shown alongside, since interest accrues across the year."),
    ("", ""),
    ("WACC", ""),
    ("Weights", "Market value of equity = 1,017,766,288 shares outstanding at 31-Mar-2026 x the latest unadjusted close "
                "from the Data sheet. Market value of debt is taken at book: the borrowings are short-dated and largely "
                "floating or lease obligations, so book value is a close proxy and is the normal convention."),
    ("Shares outstanding", "1,017,766,288 ordinary shares at 31-Mar-2026 (Yahoo Finance fundamentals; reconciles to the "
                           "Rs 102 crore equity capital reported by screener.in at Re 1 face value)."),
    ("Near-zero gearing", "Debt is under 1% of total capital, so WACC sits within a few basis points of the cost of equity. "
                          "The sensitivity grid on the WACC sheet shows what the number would become at higher target "
                          "debt weights, holding the component costs fixed."),
    ("Net cash", "Cash and equivalents of Rs 232.49 crore and cash plus short-term investments of Rs 4,215.98 crore at "
                 "31-Mar-2026 both exceed total borrowings of Rs 417.21 crore, so Pidilite is net cash. Net debt is "
                 "negative, which makes a net-debt-weighted WACC meaningless; gross debt is used for the weights."),
    ("", ""),
    ("DCF", ""),
    ("Method", "Free cash flow to firm: FCFF = EBIT x (1 - tax) + D&A - capex - increase in net working capital. "
               "Five explicit forecast years (FY2027-FY2031) plus a Gordon-growth terminal value, discounted at the WACC "
               "from the WACC sheet. Enterprise value less net debt gives equity value, divided by shares outstanding."),
    ("Historical base", "FY2023-FY2026 consolidated actuals, same Yahoo Finance / screener.in source as the Cost of Debt sheet. "
                        "EBIT is taken as OPERATING income, which excludes other income - treasury returns are already "
                        "captured by adding net cash, so counting them in EBIT as well would double count them."),
    ("Forecast assumptions", "Revenue growth fades 10% / 9% / 8% / 7% / 6% toward a 5% terminal rate, against FY2026 actual "
                             "growth of 11.1% and a three-year CAGR of 7.4%. EBIT margin held flat at the FY2026 level of "
                             "21.63% rather than extrapolating the 14.9% to 21.6% climb of the last four years. D&A 2.65% and "
                             "capex 4.09% of revenue are historical averages; working capital is charged at 6.00% of the "
                             "INCREASE in revenue. Every one of these is an input cell and can be overwritten."),
    ("Valuation date", "Cash flows are discounted from the 31-Mar-2026 balance sheet date at whole-year end-of-period "
                       "intervals. No stub adjustment is made for the months already elapsed in FY2027, and no mid-year "
                       "convention is applied; both would raise the value modestly."),
    ("Terminal value", "Carries about 74% of enterprise value, which is normal for a five-year forecast but means the answer "
                       "is driven mainly by the terminal growth rate and the WACC. The sensitivity grid shows value per share "
                       "across WACC of 10.85%-12.85% and terminal growth of 4.0%-6.0%."),
    ("The result", "On these assumptions the DCF values Pidilite far below its market price. That is a statement about the "
                   "assumptions, not a recommendation. Pidilite trades on a premium multiple (roughly 65x earnings), and a "
                   "DCF discounting at ~11.9% with 5% perpetual growth cannot reproduce that. The reverse DCF in section 6 "
                   "shows the market is implying a perpetual growth rate of about 10.5% - above the 6.97% risk-free rate, "
                   "which no company can sustain forever. The honest reading is that the gap sits in the growth and margin "
                   "assumptions, and that a five-year fade may simply be too short a runway for this franchise."),
    ("", ""),
    ("DCF - 10 year scenario", ""),
    ("Design", "A controlled test, not a second opinion. Margin, D&A, capex, working capital intensity, tax, WACC and "
               "terminal growth are all linked straight from the 5-year DCF sheet, so they cannot drift apart. The only "
               "differences are the length of the explicit window (ten years, FY2027-FY2036) and the growth path."),
    ("Growth path", "Fades linearly by about 0.6pp a year from 11.0% to 5.6%, arriving beside the 5.0% terminal rate. The "
                    "5-year case compresses the same journey into half the time, which forces a faster decay than a "
                    "franchise of this quality plausibly experiences - that compression is exactly what this scenario relaxes."),
    ("Margin switch", "The sheet interpolates EBIT margin linearly between two input cells, both defaulting to the FY2026 "
                      "level of 21.63%. Left alone the margin is flat, keeping the comparison like for like; raising the "
                      "final-year cell models operating leverage over the longer horizon."),
    ("What it shows", "Value per share rises from about INR 394 to about INR 443, an uplift of roughly 13%, and the terminal "
                      "value's share of enterprise value falls from about 74% to about 56% - a genuinely better-conditioned "
                      "model, since less of the answer rests on a single perpetuity assumption. But the valuation gap to the "
                      "market barely narrows, from about -75% to about -71%, and the reverse DCF still implies a perpetual "
                      "growth rate above the risk-free rate. The conclusion is that the gap is NOT an artefact of too short a "
                      "forecast window: it lives in the growth and margin trajectory, or in the risk premium, not in the "
                      "length of the runway."),
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


# ---------------- Summary (built last, moved to the front) ----------------
sm = wb.create_sheet("Summary")
sm["A1"] = "Pidilite Industries - Beta, Cost of Capital and DCF Valuation"; sm["A1"].font = Font(name=FONT, size=16, bold=True, color="1F3864")
sm["A2"] = "Summary of key outputs. Every figure is linked live from the sheet that computes it - nothing here is retyped."
sm["A2"].font = Font(name=FONT, size=10, italic=True)

def puts(row, label, value, fmt=None, note=None, kind="link", indent=0):
    return put(sm, row, label, value, fmt, note, kind, indent, note_col=3)

def band2(row, text):
    c = sm.cell(row=row, column=1, value=text)
    c.font = Font(name=FONT, size=11, bold=True, color="FFFFFF"); c.fill = hdr_fill
    for col in (2, 3):
        sm.cell(row=row, column=col).fill = hdr_fill
    return row + 1

def big(row):
    sm.cell(row=row, column=2).font = Font(name=FONT, size=12, bold=True, color="C00000")
    sm.cell(row=row, column=2).fill = key_fill

r = 4
r = band2(r, "SCOPE")
r = puts(r, "Company", "Pidilite Industries Ltd (NSE: PIDILITIND)", None,
         "Read as the requested 'Pidilite Inc' - no separately listed entity carries that exact name.", "formula")
r = puts(r, "Market index", "NIFTY 50 (^NSEI)", None, "The x variable in the regression and the market proxy in CAPM.", "formula")
r = puts(r, "Return window - from", f"=Data!A2", "yyyy-mm-dd", "Five years of daily bars, rolled forward each time the workbook is rebuilt.")
r = puts(r, "Return window - to", f"=Data!A{last}", "yyyy-mm-dd")
r = puts(r, "Trading days / return observations", f"=Regression!B{ROW['n']}", "#,##0", "Sessions on which both the index and the stock traded.")
r += 1

r = band2(r, "1. BETA  (Regression sheet)")
S_B = r
r = puts(r, "Beta - slope of Pidilite on NIFTY 50", f"=Regression!B{ROW['slope']}", "0.0000",
         "Below 1.0: Pidilite moves about 0.69% for each 1% move in the index. A defensive, below-market beta."); big(S_B)
r = puts(r, "R-squared", f"=Regression!B{ROW['r2']}", "0.0000",
         "The index explains only about a fifth of Pidilite's daily variance; the rest is stock-specific.")
r = puts(r, "Correlation with NIFTY 50", f"=Regression!B{ROW['corr']}", "0.0000")
r = puts(r, "Annualised volatility - Pidilite", f"=Regression!B{ROW['ay']}", "0.00%")
r = puts(r, "Annualised volatility - NIFTY 50", f"=Regression!B{ROW['ax']}", "0.00%")
r += 1

r = band2(r, "2. COST OF EQUITY  (Cost of Equity sheet)")
r = puts(r, "Risk-free rate - India 10Y G-Sec", f"='Cost of Equity'!B{R_RF}", "0.00%", "Benchmark yield at 10-Sep-2026.")
r = puts(r, "Market risk premium used", f"='Cost of Equity'!B{R_SEL}", "0.00%",
         "Damodaran implied India ERP. A trailing 5-year realised premium is near zero on this window and is not usable - "
         "see approaches A, B and C on the Cost of Equity sheet.")
S_KE = r
r = puts(r, "COST OF EQUITY (Ke)", f"='Cost of Equity'!B{R_KE}", "0.00%", "Ke = Rf + beta x MRP."); big(S_KE)
r += 1

r = band2(r, "3. COST OF DEBT  (Cost of Debt sheet)")
r = puts(r, "Total interest, FY2026 (INR crore)", f"='Cost of Debt'!{lastcol}{R_INT}", "#,##0.00", "Consolidated finance costs.")
r = puts(r, "Total borrowings, FY2026 (INR crore)", f"='Cost of Debt'!{lastcol}{R_BOR}", "#,##0.00",
         "As reported, including Ind AS 116 lease liabilities - of which only about INR 106 crore is actual borrowing.")
r = puts(r, "Cost of debt, pre-tax (Kd)", f"='Cost of Debt'!{lastcol}{R_KDC}", "0.00%", "Total interest / total borrowings.")
r = puts(r, "Effective tax rate", f"='Cost of Debt'!{lastcol}{R_ETR}", "0.00%")
S_KD = r
r = puts(r, "COST OF DEBT, AFTER TAX", f"='Cost of Debt'!{lastcol}{R_ATK}", "0.00%"); big(S_KD)
r += 1

r = band2(r, "4. WACC  (WACC sheet)")
r = puts(r, "Market value of equity (INR crore)", f"=WACC!B{W_E}", "#,##0.00")
r = puts(r, "Total borrowings (INR crore)", f"=WACC!B{W_D}", "#,##0.00")
r = puts(r, "Weight of equity", f"=WACC!B{W_WE}", "0.00%")
r = puts(r, "Weight of debt", f"=WACC!B{W_WD}", "0.00%", "Under 1%, so WACC lands within a basis point of Ke.")
S_W = r
r = puts(r, "WACC", f"=WACC!B{W_WACC}", "0.00%"); big(S_W)
r = puts(r, "Net debt (INR crore)", f"=WACC!B{W_ND}", "#,##0.00",
         "Negative - cash and short-term investments exceed borrowings, so Pidilite is NET CASH.")
r += 1

r = band2(r, "5. DCF VALUATION  (DCF and DCF 10Y sheets)")
r = puts(r, "Enterprise value, 5-year base case (INR crore)", f"=DCF!B{D_EV}", "#,##0.00")
r = puts(r, "    Terminal value as % of EV", f"=DCF!B{D_TVP}", "0.00%", "Most of the answer sits in the terminal assumption.")
S_V5 = r
r = puts(r, "Value per share - 5-year base case (INR)", f"=DCF!B{D_VPS}", "#,##0.00"); big(S_V5)
r = puts(r, "    Terminal value as % of EV, 10-year", f"='DCF 10Y'!B{X_TVP}", "0.00%",
         "Falls sharply against the 5-year case - a better-conditioned model, since less rests on the perpetuity.")
S_V10 = r
r = puts(r, "Value per share - 10-year scenario (INR)", f"='DCF 10Y'!B{X_VPS}", "#,##0.00",
         "The only changes from the base case are the forecast window and the growth path."); big(S_V10)
S_MP = r
r = puts(r, "Current market price (INR)", f"=DCF!B{D_MP}", "#,##0.00")
r = puts(r, "    Implied upside - 5-year base case", f"=DCF!B{D_UP}", "0.00%")
r = puts(r, "    Implied upside - 10-year scenario", f"='DCF 10Y'!B{X_U10}", "0.00%")
r += 1

r = band2(r, "6. REVERSE DCF - WHAT THE MARKET IS PRICING")
r = puts(r, "Implied perpetual growth - 5-year forecast", f"=DCF!B{D_IG}", "0.00%")
r = puts(r, "Implied perpetual growth - 10-year forecast", f"='DCF 10Y'!B{X_IG}", "0.00%")
r = puts(r, "Compare with the risk-free rate", f"='Cost of Equity'!B{R_RF}", "0.00%",
         "Both implied rates sit ABOVE the risk-free rate. No company grows faster than its economy in perpetuity, so the "
         "market is pricing stronger growth, higher margins or a lower risk premium than these models assume.")
r += 1

r = band2(r, "HOW TO READ THIS")
for txt in [
    "The sheets are chained. Change the market risk premium on the Cost of Equity sheet and it flows through Ke, WACC and both DCFs.",
    "Blue cells are hardcoded inputs, green cells are links from another sheet, black cells are formulas. Every input is sourced on the Notes sheet.",
    "The valuation gap is a statement about the assumptions, not a recommendation. Section 6 locates where the gap actually lives.",
    "NSE returns HTTP 403 to automated access, so financials come from Yahoo Finance cross-checked line by line against screener.in; every figure agreed.",
]:
    sm.cell(row=r, column=1, value="-  " + txt).font = base
    r += 1
r += 1

r = band2(r, "SHEET INDEX")
for nm, desc in [
    ("Data", "Five years of daily closes, adjusted closes and daily % returns for both instruments."),
    ("Regression", "OLS of Pidilite returns on NIFTY 50 returns, with the scatter chart and fitted line."),
    ("Cost of Equity", "CAPM, three market risk premium approaches, and a beta x MRP sensitivity grid."),
    ("Cost of Debt", "Interest over borrowings, FY2023-FY2026, closing and average, with the lease split."),
    ("WACC", "Market-value weights, component costs and a target-gearing sensitivity."),
    ("DCF", "Five-year FCFF forecast, terminal value, sensitivity grid and reverse DCF."),
    ("DCF 10Y", "Ten-year scenario, controlled against the base case."),
    ("Notes", "Every source, definition, assumption and caveat."),
]:
    sm.cell(row=r, column=1, value=nm).font = bold
    sm.cell(row=r, column=3, value=desc).font = base
    r += 1

sm.column_dimensions["A"].width = 44
sm.column_dimensions["B"].width = 18
sm.column_dimensions["C"].width = 104
sm.sheet_view.showGridLines = False
wb.move_sheet("Summary", offset=-(len(wb.sheetnames) - 1))

wb.calculation.fullCalcOnLoad = True
wb.save(OUT)
print("saved", OUT)
