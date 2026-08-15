"""
Generate a Group driller whose tabs are five business groups side by side —
IWPB and CIB (classic global businesses with countries) and HK, UK and
Corp Centre (entity groups that consolidate like businesses): one tab each,
the same MICA shape on every tab, Corp Centre cost-heavy the way the real
one is. The Group figure must read as the sum of all five.
"""
import random
from openpyxl import Workbook

random.seed(11)

DIM_HEADERS = [
    "MICA", "MICA_Level_4", "MICA_Level_3", "MICA_Level_2", "MICA_Level_1",
    "Product_code", "Product_Level_3", "Product_Level_2", "Product_Level_1",
    "Entity code",
]

MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

PERIODS = (
    [(f"{m}-25", "Actuals") for m in MONTHS] +
    [(f"{m}-26", "Actuals") for m in MONTHS[:6]] +
    [(f"{m}-26", "Forecast") for m in MONTHS[6:]] +
    [("JUN YTD-25", "Actuals"), ("JUN YTD-26", "Actuals"),
     ("JUN YTD-26 Target", "Target"),
     ("FY-25", "Actuals"), ("FY-26 Forecast", "Forecast"), ("FY-26 Target", "Target")]
)

MICA_ROWS = [
    ("MP10101010000", "NII - Interest Income",   "NII - Net Interest Income", "Revenue", +1),
    ("MP10201010000", "NFI - Fee Income",        "Net Fee Income",            "Revenue", +1),
    ("MP10301010000", "Insurance Manufacturing", "Net Insurance Revenue",     "Revenue", +1),
    ("MP20101010000", "Staff Costs",             "Direct Costs",              "Total Direct Cost", -1),
    ("MP30101010000", "ECL Charge",              "ECL",                       "Expected credit losses", -1),
]

PRODUCTS = [
    ("PR04020000", "Mortgages (Real Estate Secured)", "Loans"),
    ("PR05001000", "Current Accounts Total",          "Deposits"),
]

# tab name → revenue scale, cost scale (Corp Centre is a cost centre)
TABS = [
    ("IWPB India",    40, 18),
    ("IWPB France",   25, 12),
    ("CIB Singapore", 55, 24),
    ("HK",            70, 30),
    ("UK",            45, 22),
    ("Corp Centre",    4, 26),
]

wb = Workbook()
first = True
for tab, rev_scale, cost_scale in TABS:
    ws = wb.active if first else wb.create_sheet()
    first = False
    ws.title = tab
    ws.append([""] * len(DIM_HEADERS) + [b for _, b in PERIODS])
    ws.append(DIM_HEADERS + [h for h, _ in PERIODS])
    for mica, l4, l3, l2, sign in MICA_ROWS:
        scale = rev_scale if sign > 0 else cost_scale
        for pcode, p3, p2 in PRODUCTS:
            base = random.uniform(scale * 0.7, scale * 1.3)
            drift = random.uniform(0.92, 1.15)
            py = [round(sign * abs(random.gauss(base, base * .2)), 1) for _ in range(12)]
            cy = [round(sign * abs(random.gauss(base * drift, base * .2)), 1) for _ in range(6)]
            fc = [round(sign * abs(random.gauss(base * drift, base * .2)), 1) for _ in range(6)]
            ytd_py = round(sum(py[:6]), 1)
            ytd_cy = round(sum(cy), 1)
            ytd_tgt = round(ytd_cy * random.uniform(0.94, 1.06), 1)
            fy25 = round(sum(py), 1)
            fy26fc = round(sum(cy) + sum(fc), 1)
            fy26tgt = round(fy26fc * random.uniform(0.93, 1.07), 1)
            dims = [mica, l4, l3, l2, "PBT", pcode, p3, p2, "Total Product", "4040_1MGT"]
            ws.append(dims + py + cy + fc +
                      [ytd_py, ytd_cy, ytd_tgt, fy25, fy26fc, fy26tgt])

out = "IWPB_FiveGroups_Driller.xlsx"
wb.save(out)
print(f"wrote {out}: tabs {[t for t,_,_ in TABS]}")
