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
