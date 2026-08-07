"""
Generate a sample driller in the CIB-West-style layout, to test the second
supported format:
  - a Control tab before the data sheet
  - title rows above the table on the data sheet
  - an Actuals / Forecast / Target band row above the column headers
  - a single year of months (Jan-Jun actuals, Jul-Dec forecast), quarters
    without year digits, a Jun YTD column, FY actuals and an FY target column
"""
import random
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

random.seed(11)

DIM_HEADERS = [
    "MICA", "MICA_Level_3", "MICA_Level_2", "MICA_Level_1",
    "Product_code", "Product_Level_3", "Product_Level_2", "Product_Level_1",
    "Segment_code", "CG_Level_2", "CG_Level_1",
    "Function_code", "Function_Level_2", "Function_Level_1",
    "Entity code",
]
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

PERIODS = (
    [(f"{m}-26", "Actuals") for m in MONTHS[:6]] +
    [(f"{m}-26", "Forecast") for m in MONTHS[6:]] +
    [("Q1-26", "Actuals"), ("Q2-26", "Actuals"), ("Q3-26", "Forecast"), ("Q4-26", "Forecast")] +
    [("Jun YTD-26", "Actuals"), ("FY-25", "Actuals"), ("FY-26 Forecast", "Forecast"), ("FY-26 Target", "Target")]
)

MICA_ROWS = [
    # P&L, all rolling into PBT at MICA Level 1
    ("MP10101010000", "NII - Interest Income",   "Revenue",                            "PBT",      +1),
    ("MP10102010000", "NFI - Fee Income",        "Revenue",                            "PBT",      +1),
    ("MP20101000000", "Staff Costs",             "Total Direct Cost",                  "PBT",      -1),
    ("MP20102000000", "Allocated Costs",         "Total Indirect Costs (incl InterCo)","PBT",      -1),
    ("MP30101000000", "ECL Charge",              "Expected credit losses",             "PBT",      -1),
    # balance sheet, separate Level 1 lines
    ("MB40101000000", "Customer Deposits",       "Deposits",                           "Deposits", +1),
    ("MB40201000000", "Customer Loans",          "Loans",                              "Loans",    +1),
]
PRODUCTS = [
    ("UNMAPPDT",   "Unmapped Product",                "Unmapped Product"),
    ("PR04000002", "Wealth Lending",                  "Loans"),
    ("PR04004001", "Other Lending",                   "Loans"),
    ("PR04020000", "Mortgages (Real Estate Secured)", "Loans"),
    ("PR04030000", "Credit Cards",                    "Cards"),
    ("PR04060600", "Structured Products",             "Investments"),
    ("PR05001000", "Current Accounts Total",          "Deposits"),
    ("PR05040000", "Savings and Deposit Accounts",    "Deposits"),
    ("PR04060550", "Other products",                  "Total Other Products"),
    ("PR10090005", "Non Product Allocation of Corp Centre", "Non Product Results Total"),
]
SEG = ("CG1010000", "Retail Banking & Wealth Management", "IWPB")
FUNCS = [("RTN00008", "Premier and Wealth", "Intl Wealth & Premier Banking"),
         ("RTN00012", "Personal Banking",   "Intl Wealth & Premier Banking")]
ENTITY = "4040_1MGT"

wb = Workbook()
ctl = wb.active
ctl.title = "Control"
ctl.append(["CIB West Product Driller — control sheet"])
ctl.append(["Refresh date", "2026-07-15"])

ws = wb.create_sheet("EU&AM")
ws.append(["CIB West Product Driller"])
ws.append(["EU&AM — Revenue ex Notables ($m)"])
ws.append([])
ws.append([""] * len(DIM_HEADERS) + [b for _, b in PERIODS])   # band row
ws.append(DIM_HEADERS + [h for h, _ in PERIODS])               # header row

def cell(v):
    """Render a value, sometimes as bracket-negative text to test the parser."""
    v = round(v, 1)
    if v < 0 and random.random() < 0.25:
        return f"({abs(v):,.1f})"
    return v

for mica, l3, l2, l1, sign in MICA_ROWS:
    for pcode, p3, p2 in PRODUCTS:
        for fcode, f2, f1 in FUNCS:
            if random.random() < 0.2:
                continue
            base = random.uniform(1, 40) if l1=='PBT' else random.uniform(400, 2000)
            dims = [mica, l3, l2, l1, pcode, p3, p2, "Total Product",
                    SEG[0], SEG[1], SEG[2], fcode, f2, f1, ENTITY]
            # 12 monthly values first, then derive the roll-up columns from
            # them so quarters / YTD / FY / target reconcile like a real file
            months = [sign * abs(random.gauss(base, base * 0.35)) for _ in range(12)]
            derived = {
                "Q1-26": sum(months[0:3]),  "Q2-26": sum(months[3:6]),
                "Q3-26": sum(months[6:9]),  "Q4-26": sum(months[9:12]),
                "Jun YTD-26": sum(months[:6]),
                "FY-25": sum(months) * random.uniform(0.85, 1.1),
                "FY-26 Forecast": sum(months),
                "FY-26 Target": sum(months) * random.uniform(0.92, 1.08),
            }
            vals = []
            for i, (h, band) in enumerate(PERIODS):
                vals.append(cell(months[i] if i < 12 else derived[h]))
            ws.append(dims + vals)

out = "IWPB_SG_Driller_mixed.xlsx"
wb.save(out)
print(f"wrote {out}: {ws.max_row-5} data rows, sheets: {wb.sheetnames}")
