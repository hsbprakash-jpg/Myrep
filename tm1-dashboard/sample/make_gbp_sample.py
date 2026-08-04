"""
Generate a synthetic 'Group Business Performance' driller sample that mirrors
the structure of a group P&L / balance-sheet / headcount overview report,
without containing any real figures. Every number is drawn from a seeded RNG
around arbitrary round bases.

Two sheets in GBP_Driller_sample.xlsx:
  'GBP Data'  - flat TM1-style driller extract (one header row, dimension
                columns then period columns) in the exact shape index.html
                loads: drill Ac Type > Ac Line > Prd > Product 1.
  'Overview'  - formatted report replica with native Excel outline
                (expand/collapse) groups on rows and on the monthly columns.

Structure notes
  * Periods: Jan-Jun actuals, Jul-Dec forecast, FY forecast, FY target,
    FY run-rate, Risk/Opp (= run-rate - target), and quarters. FY equals the
    sum of months for flow items and the December value for point-in-time
    items (balance sheet, headcount).
  * 'o/w' (of-which) lines are partial breakdowns; an 'o/w Other (residual)'
    leaf completes each parent so the drill tree reconciles exactly.
  * The Reported and ex-Notables P&L views share one tree via the 'Basis'
    dimension: filter Basis = Underlying for ex-Notables; leave unfiltered
    for Reported (P&L pane grand total = Reported PBT).
  * Costs/ECLs are negative so PBT is a straight sum.
  * Units by pane: P&L $m, Balance Sheet $bn, RWA $bn, FTE counts.
  * 'Total Headcount' appears only on the Overview sheet (in the flat
    extract it would double count FTEs + Contractors).
"""
import random

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

random.seed(7)

# ---------------------------------------------------------------- period axis
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
PERIOD_HEADERS = (
    [f"{m}'26-Act" for m in MONTHS[:6]] +
    [f"{m}'26-FC" for m in MONTHS[6:]] +
    ["FY'26-FC", "FY'26-Tar", "FY'26-RR", "FY'26-RiskOpp",
     "Q1'26-Act", "Q2'26-Act", "Q3'26-FC", "Q4'26-FC"]
)
N = len(PERIOD_HEADERS)  # 20


def flow(base, vol=0.10):
    """Monthly flow series + FY/target/run-rate/RiskOpp + quarters."""
    months = [round(base * random.uniform(1 - vol, 1 + vol)) for _ in range(12)]
    fy = sum(months)
    tar = round(fy * random.uniform(0.95, 1.03))
    rr = round(fy * random.uniform(0.96, 1.06))
    qs = [sum(months[i:i + 3]) for i in (0, 3, 6, 9)]
    return months + [fy, tar, rr, rr - tar] + qs


def stock(base, drift=0.004):
    """Point-in-time random walk; FY = December, target only, no quarters."""
    months, v = [], base
    for _ in range(12):
        v *= 1 + random.gauss(drift, 0.006)
        months.append(round(v))
    tar = round(months[-1] * random.uniform(0.98, 1.02))
    return months + [months[-1], tar] + [None] * 6


def add(*rows):
    out = []
    for vals in zip(*rows):
        nums = [v for v in vals if v is not None]
        out.append(round(sum(nums)) if nums else None)
    return out


# ------------------------------------------------------------ synthetic lines
D = {
    # P&L leaves ($m); arbitrary round bases, no relation to any real report
    "NII":       flow(2000),
    "FoTB":      flow(500),
    "BankOther": flow(-20, vol=0.5),
    "WTB":       flow(800),
    "Wealth":    flow(600),
    "FeesOther": flow(400),
    "RevNtb":    flow(-40, vol=1.5),
    "ECL":       flow(-300, vol=0.6),
    "OpExXVP":   flow(-1800),
    "VP":        flow(-250),
    "OpExNtb":   flow(-80, vol=0.8),
    "IfA":       flow(150, vol=0.3),
    "IfANtb":    [None] * N,
    # Balance sheet ($bn), RWA ($bn), headcount (#): point-in-time
    "Deposits":  stock(1500),
    "Loans":     stock(900),
    "RWA":       stock(700),
    "FTE":       stock(150000, drift=-0.004),
    "Contract":  stock(5000),
}
# parents derived by summation, so every drill level reconciles exactly
D["BankNII"] = add(D["NII"], D["FoTB"], D["BankOther"])
D["Fees"] = add(D["WTB"], D["Wealth"], D["FeesOther"])
D["TotRevXN"] = add(D["BankNII"], D["Fees"])
D["TotRevRep"] = add(D["TotRevXN"], D["RevNtb"])
D["NonNII"] = add(D["TotRevRep"], [-v for v in D["NII"]])
D["OpExXN"] = add(D["OpExXVP"], D["VP"])
D["OpExRep"] = add(D["OpExXN"], D["OpExNtb"])
D["PBTXN"] = add(D["TotRevXN"], D["ECL"], D["OpExXN"], D["IfA"])
D["PBTRep"] = add(D["TotRevRep"], D["ECL"], D["OpExRep"], D["IfA"], D["IfANtb"])
D["TotalHC"] = add(D["FTE"], D["Contract"])
for k, v in D.items():
    assert len(v) == N, f"{k}: {len(v)} values, expected {N}"

# ---------------------------------------------------------- flat extract rows
DIM_HEADERS = ["Unique Ref", "Unique Ref1", "Ac Type", "Seg", "Markets",
               "Cntry name", "Prd", "Prd1", "Ac Line", "Ac Line1", "Entity",
               "Account", "RTN", "CG code", "Product 1", "Basis"]
ENTITY, SEG = "GROUP", "Group"

#        Ac Type, Account (units),                      Ac Line,               Prd,                                  Product 1,                                Basis,       data key
ROWS = [
    ("P&L", "Group P&L @ Current FX ($m)", "Total Revenue",          "Banking NII",                        "o/w NII",                                "Underlying", "NII"),
    ("P&L", "Group P&L @ Current FX ($m)", "Total Revenue",          "Banking NII",                        "o/w FoTB",                               "Underlying", "FoTB"),
    ("P&L", "Group P&L @ Current FX ($m)", "Total Revenue",          "Banking NII",                        "o/w Other (residual)",                   "Underlying", "BankOther"),
    ("P&L", "Group P&L @ Current FX ($m)", "Total Revenue",          "Fees and other operating income",    "o/w WTB",                                "Underlying", "WTB"),
    ("P&L", "Group P&L @ Current FX ($m)", "Total Revenue",          "Fees and other operating income",    "o/w Wealth",                             "Underlying", "Wealth"),
    ("P&L", "Group P&L @ Current FX ($m)", "Total Revenue",          "Fees and other operating income",    "o/w Other (residual)",                   "Underlying", "FeesOther"),
    ("P&L", "Group P&L @ Current FX ($m)", "Total Revenue",          "Revenue Notables",                   "Revenue Notables",                       "Notables",   "RevNtb"),
    ("P&L", "Group P&L @ Current FX ($m)", "ECLs",                   "ECLs",                               "ECLs",                                   "Underlying", "ECL"),
    ("P&L", "Group P&L @ Current FX ($m)", "Operating Expenses",     "OpEx (ex. VP)",                      "OpEx (ex. VP)",                          "Underlying", "OpExXVP"),
    ("P&L", "Group P&L @ Current FX ($m)", "Operating Expenses",     "VP",                                 "VP",                                     "Underlying", "VP"),
    ("P&L", "Group P&L @ Current FX ($m)", "Operating Expenses",     "Opex Notables",                      "Opex Notables",                          "Notables",   "OpExNtb"),
    ("P&L", "Group P&L @ Current FX ($m)", "Income from Associates", "Income from Associates",             "Income from Associates",                 "Underlying", "IfA"),
    ("P&L", "Group P&L @ Current FX ($m)", "Income from Associates", "Income from Associates Notables",    "Income from Associates Notables",        "Notables",   "IfANtb"),
    ("Balance Sheet", "Balance Sheet ($bn)", "Balance Sheet",        "Deposits",                           "Deposits",                               "Underlying", "Deposits"),
    ("Balance Sheet", "Balance Sheet ($bn)", "Balance Sheet",        "Loans and Advances",                 "Loans and Advances",                     "Underlying", "Loans"),
    ("RWA", "RWAs - Current Methodology ($bn)", "RWAs",              "RWAs - Current Methodology",         "RWAs - Current Methodology",             "Underlying", "RWA"),
    ("Metrics", "Headcount (#)", "Headcount",                        "FTE",                                "FTE - Headcount - Total FTEs",           "Underlying", "FTE"),
    ("Metrics", "Headcount (#)", "Headcount",                        "FTE",                                "FTE - Contractor - Total Contractors",   "Underlying", "Contract"),
]

wb = Workbook()
ws = wb.active
ws.title = "GBP Data"
ws.append(DIM_HEADERS + PERIOD_HEADERS)
for ac_type, account, ac_line, prd, product1, basis, key in ROWS:
    uref = f"{ac_type}-{prd}-{ENTITY}-{SEG}"
    dims = [uref, uref, ac_type, SEG, SEG, "Total", prd, prd, ac_line,
            ac_line, ENTITY, account, "RTN00000", "CG0000000", product1, basis]
    ws.append(dims + D[key])
for cell in ws[1]:
    cell.font = Font(name="Arial", bold=True, size=9)
for col in range(1, ws.max_column + 1):
    ws.column_dimensions[get_column_letter(col)].width = 13 if col > len(DIM_HEADERS) else 22
for row in ws.iter_rows(min_row=2):
    for cell in row:
        cell.font = Font(name="Arial", size=9)
ws.freeze_panes = "A2"

# ================================================== formatted Overview sheet
ov = wb.create_sheet("Overview")
RED, GREY, LGREY = "DB0011", "BFBFBF", "F2F2F2"
NUMFMT = '#,##0;(#,##0);"-"'
thin = Side(style="thin", color="D9D9D9")

def style(cell, bold=False, fill=None, color="000000", size=10, center=False):
    cell.font = Font(name="Arial", bold=bold, color=color, size=size)
    if fill:
        cell.fill = PatternFill("solid", fgColor=fill)
    if center:
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

ov["A1"] = "Group Business Performance overview (synthetic sample)"
style(ov["A1"], bold=True, size=16)
ov["A2"] = ("Synthetic demo data generated by make_gbp_sample.py - every figure is "
            "random. Jan-Jun actuals, Jul-Dec forecast, FY'26.")
style(ov["A2"], size=8, color="808080")

BANDS = [("B", "G", "Actuals", GREY, "000000"), ("H", "N", "Forecast", RED, "FFFFFF"),
         ("O", "O", "Target", RED, "FFFFFF"), ("P", "P", "", None, "000000"),
         ("Q", "Q", "", None, "000000"), ("R", "S", "Actuals", GREY, "000000"),
         ("T", "U", "Forecast", RED, "FFFFFF")]
for c1, c2, label, fill, color in BANDS:
    if c1 != c2:
        ov.merge_cells(f"{c1}3:{c2}3")
    ov[f"{c1}3"] = label
    style(ov[f"{c1}3"], bold=True, fill=fill, color=color, center=True)
COLHEADS = (MONTHS + ["FY", "FY", "FY based on Run Rate", "Risk/Opp",
                      "Q1", "Q2", "Q3", "Q4"])
for i, h in enumerate(COLHEADS):
    cell = ov.cell(row=4, column=2 + i, value=h)
    style(cell, bold=True, fill=LGREY, center=True)
    cell.border = Border(bottom=thin)

# (label, outline level, data key, bold)
SECTIONS = [
    ("P&L - Reported ($m) @ Current FX", [
        ("NII", 2, "NII", False),
        ("Non-NII", 2, "NonNII", False),
        ("Total Revenue", 1, "TotRevRep", True),
        ("ECLs", 1, "ECL", False),
        ("Operating Expenses", 1, "OpExRep", False),
        ("Income from Associates", 1, "IfA", False),
        ("PBT", 0, "PBTRep", True),
    ]),
    ("P&L - ex Notables ($m) @ Current FX", [
        ("Banking NII", 2, "BankNII", False),
        ("   o/w NII", 3, "NII", False),
        ("   o/w FoTB", 3, "FoTB", False),
        ("Fees and other operating income", 2, "Fees", False),
        ("   o/w WTB", 3, "WTB", False),
        ("   o/w Wealth", 3, "Wealth", False),
        ("Total Revenue", 1, "TotRevXN", True),
        ("ECLs", 1, "ECL", False),
        ("   OpEx (ex. VP)", 2, "OpExXVP", False),
        ("   VP", 2, "VP", False),
        ("Operating Expenses", 1, "OpExXN", False),
        ("Income from Associates", 1, "IfA", False),
        ("PBT", 1, "PBTXN", True),
        ("   Revenue Notables", 2, "RevNtb", False),
        ("   Opex Notables", 2, "OpExNtb", False),
        ("   Income from Associates Notables", 2, "IfANtb", False),
        ("Reported PBT", 0, "PBTRep", True),
    ]),
    ("Balance Sheet ($bn)", [
        ("Deposits", 1, "Deposits", False),
        ("Loans and Advances", 1, "Loans", False),
        ("RWAs - Current Methodology", 1, "RWA", False),
    ]),
    ("Headcount", [
        ("FTEs", 1, "FTE", False),
        ("Contractors", 1, "Contract", False),
        ("Total Headcount", 0, "TotalHC", True),
    ]),
]

r = 6
for title, lines in SECTIONS:
    cell = ov.cell(row=r, column=1, value=title)
    style(cell, bold=True, fill=RED, color="FFFFFF")
    r += 1
    for label, level, key, bold in lines:
        ov.cell(row=r, column=1, value=label)
        style(ov.cell(row=r, column=1), bold=bold)
        for i, v in enumerate(D[key]):
            c = ov.cell(row=r, column=2 + i, value=v if v is not None else 0)
            c.number_format = NUMFMT
            style(c, bold=bold)
            if bold:
                c.border = Border(top=thin)
        if level:
            ov.row_dimensions[r].outline_level = level
        r += 1
    r += 1  # blank separator

ov.column_dimensions["A"].width = 34
for col in range(2, 2 + N):
    ov.column_dimensions[get_column_letter(col)].width = 9
ov.column_dimensions["P"].width = 12
ov.column_dimensions.group("B", "M", outline_level=1)  # months drill closed to FY view
ov.sheet_properties.outlinePr.summaryBelow = True
ov.freeze_panes = "B5"

out = "GBP_Driller_sample.xlsx"
wb.save(out)
print(f"wrote {out}: '{ws.title}' {ws.max_row - 1} rows x {ws.max_column} cols; "
      f"'Overview' {ov.max_row} rows")
