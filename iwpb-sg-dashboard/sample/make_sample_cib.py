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
    [("Jun YTD-26", "Actuals"), ("FY-25", "Actuals"), ("FY-26", "Forecast"), ("FY-26 Target", "Target")]
)

MICA_ROWS = [
    ("MP10101010000", "NII - Interest Income",   "NII - Net Interest Income", "Revenue", +1),
    ("MP10101020000", "NII - Interest Expense",  "NII - Net Interest Income", "Revenue", -1),
    ("MP10102010000", "NFI - Fee Income",        "Net Fee Income",            "Revenue", +1),
    ("MP10102020000", "NFI - Fee Expense",       "Net Fee Income",            "Revenue", -1),
    ("MP10103010000", "Trading Income - NII",    "Trading Income",            "Revenue", +1),
    ("MP10103020000", "Trading Income - Income", "Trading Income",            "Revenue", +1),
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

def val(base, sign):
    if random.random() < 0.3:
        return 0
    v = round(random.gauss(base, base * 0.4), 1)
    v = sign * abs(v)
    if v < 0 and random.random() < 0.25:
        return f"({abs(v):,.1f})"
    return v

for mica, l3, l2, l1, sign in MICA_ROWS:
    for pcode, p3, p2 in PRODUCTS:
        for fcode, f2, f1 in FUNCS:
            if random.random() < 0.2:
                continue
            base = random.uniform(1, 40)
            dims = [mica, l3, l2, l1, pcode, p3, p2, "Total Product",
                    SEG[0], SEG[1], SEG[2], fcode, f2, f1, ENTITY]
            vals = []
            for h, band in PERIODS:
                if h.startswith("Q"):
                    scale = 3
                elif "YTD" in h:
                    scale = 6
                elif h.startswith("FY"):
                    scale = 12
                else:
                    scale = 1
                vals.append(val(base * scale, sign))
            ws.append(dims + vals)

out = "IWPB_SG_Driller_CIB_style.xlsx"
wb.save(out)
print(f"wrote {out}: {ws.max_row-5} data rows, sheets: {wb.sheetnames}")
