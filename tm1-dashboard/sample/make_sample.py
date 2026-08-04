"""
Generate a sample TM1 'GPS Data' extract that mirrors the real hierarchy file's
column structure, so the dashboard can be tested without exposing real data.

Structure (single header row on sheet 'GPS Data'):
  Dimension/attribute columns  -> Unique Ref ... Product 1
  Period (measure) columns      -> Jan'26-Act ... FY'26-Act, '25 grid, YTD Dec'nn
"""
import random
from openpyxl import Workbook

random.seed(42)

DIM_HEADERS = [
    "Unique Ref", "Unique Ref1", "Ac Type", "Seg", "Markets", "Cntry name",
    "Prd", "Prd1", "Ac Line", "Ac Line1", "Entity", "Account",
    "RTN", "CG code", "Product 1",
]

MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def period_headers(yy):
    cols = [f"{m}'{yy}-Act" for m in MONTHS]
    cols += [f"YTD Apr'{yy}-Act"]
    cols += [f"Q{q}'{yy}-Act" for q in range(1, 5)]
    cols += [f"H1'{yy}-Act", f"H2'{yy}-Act", f"FY'{yy}-Act"]
    return cols


TARGET_HEADERS = (
    ["Apr'26-Tar"] +
    [f"Q{q}'26-Tar" for q in range(1, 5)] +
    ["H1'26-Tar", "H2'26-Tar", "FY'26-Tar", "YTD Apr'26-Tar"] +
    [f"FY'{y}-Tar" for y in ("27", "28", "29", "30")]
)
FC_HEADERS = [f"{m}'26-FC" for m in ("Apr", "May", "Jun")]

PERIOD_HEADERS = period_headers("26") + period_headers("25") + \
    ["YTD Dec'19-Act", "YTD Dec'20-Act"] + TARGET_HEADERS + FC_HEADERS

# Semantic hierarchy: Ac Line -> Prd (product rollup) -> Product 1 (leaves)
HIER = {
    "NII": {
        "IBCA-Current Accounts": [
            "PR05010000 - Current Accounts - Other",
            "PR05011002 - Current Accounts - Interest Bearing",
            "PR05012100 - Current Accounts - Non Interest Bearing",
        ],
        "IBCA-Savings": [
            "PR05040100 - Savings Accounts - Other",
            "PR05040103 - Savings Accounts - Interest Bearing - Managed",
            "PR05040105 - Commercial Money Market",
        ],
        "IBCA-TD": [
            "PR05040200 - Time Deposits Other",
            "PR05040220 - Time Deposits 3 - 6 Months",
            "PR05040240 - Time Deposits 12 - 24 Months",
        ],
        "NIBCA-Current Accounts": [
            "PR05020000 - Money Market Call Deposits",
            "PR11020100 - GPS Current Accounts - Other",
        ],
        "NIBCA-Savings": [
            "PR05040300 - Savings Account Liabilities - Other",
            "PR05040400 - Term Deposits Liabilities - Other",
        ],
    },
    "Non-NII": {
        "Fees & Commissions": [
            "PR05070001 - Other Deposits Interest Bearing",
            "PR05080000 - Other Deposits",
        ],
        "Trading Income": [
            "PR05070003 - Global Money",
        ],
    },
}

ENTITY = "1000ALT2_1MGT"
SEG = MARKET = "GB GPS"
AC_TYPE = "P&L"


def prd_group(prd):
    if prd.startswith("IBCA"):
        return "GPS_NII_IBCA", "NII_IBCA"
    if prd.startswith("NIBCA"):
        return "GPS_NII_NIBCA", "NII_NIBCA"
    return "GPS", "NII"


def gen_value(base):
    r = random.random()
    if r < 0.08:
        return None  # truly empty
    if r < 0.12:
        return "-"   # TM1-style blank marker (text)
    v = round(random.gauss(base, base * 0.35), 2)
    if random.random() < 0.08:
        v = -abs(round(v * 0.3, 2))  # occasional negatives
    # some cells come through as formatted text in real extracts
    r2 = random.random()
    if r2 < 0.06:
        return f"({abs(v):,.2f})" if v < 0 else f"{v:,.2f}"
    return v


wb = Workbook()
ws = wb.active
ws.title = "GPS Data"
ws.append(DIM_HEADERS + PERIOD_HEADERS)

for ac_line, prds in HIER.items():
    ac_line1 = ac_line
    for prd, products in prds.items():
        prd1, _ = prd_group(prd)
        for product1 in products:
            base = random.uniform(5, 120)
            uref = f"{AC_TYPE}-{prd}-{ENTITY}-{SEG}"
            uref1 = f"{AC_TYPE}-{prd1}-{ENTITY}-{SEG}"
            account = "MP10101000000 - NII - Net Interest Income"
            dims = [uref, uref1, AC_TYPE, SEG, MARKET, ENTITY,
                    prd, prd1, ac_line, ac_line1, ENTITY, account,
                    "RTN16559", "CG3000000", product1]
            vals = [gen_value(base) for _ in PERIOD_HEADERS]
            ws.append(dims + vals)

# derived subtotal rows, as real extracts contain: marked "Calculate(d)" in a
# dimension column; their values duplicate the leaves and must NOT be summed
for ac_line in HIER:
    dims = [f"{AC_TYPE}-Calculate-{ENTITY}-{SEG}", f"{AC_TYPE}-Calculate-{ENTITY}-{SEG}",
            AC_TYPE, SEG, MARKET, ENTITY,
            "Calculated", "Calculated", ac_line, ac_line, ENTITY,
            "MP10101000000 - NII - Net Interest Income",
            "RTN16559", "CG3000000", f"Calculate - {ac_line} Total"]
    vals = [round(random.uniform(500, 2000), 2) for _ in PERIOD_HEADERS]
    ws.append(dims + vals)

# non-P&L panes: Avg Bal, RWA and Metrics rows share the same layout but a
# different Ac Type — they must never be summed together with P&L
OTHER_CATS = {
    "Avg Bal": [
        ("IBCA-Current Accounts", "PR05010000 - Current Accounts - Other", 60000),
        ("IBCA-Savings", "PR05040100 - Savings Accounts - Other", 45000),
        ("NIBCA-Current Accounts", "PR05020000 - Money Market Call Deposits", 30000),
        ("IBCA-TD", "PR05040200 - Time Deposits Other", 25000),
    ],
    "RWA": [
        ("IBCA-Current Accounts", "PR05010000 - Current Accounts - Other", 8000),
        ("IBCA-TD", "PR05040200 - Time Deposits Other", 5000),
    ],
    "Metrics": [
        ("IBCA-Current Accounts", "PR05010000 - Current Accounts - Other", 2),
        ("IBCA-Savings", "PR05040100 - Savings Accounts - Other", 1),
    ],
}
for cat, rows_ in OTHER_CATS.items():
    for prd, product1, base in rows_:
        prd1, _ = prd_group(prd)
        dims = [f"{cat}-{prd}-{ENTITY}-{SEG}", f"{cat}-{prd1}-{ENTITY}-{SEG}",
                cat, SEG, MARKET, ENTITY,
                prd, prd1, "NII", "NII", ENTITY,
                f"MB20101000000 - {cat}",
                "RTN16559", "CG3000000", product1]
        vals = [gen_value(base) for _ in PERIOD_HEADERS]
        ws.append(dims + vals)

out = "GPS_Driller_sample.xlsx"
wb.save(out)
print(f"wrote {out}: {ws.max_row-1} rows x {ws.max_column} cols "
      f"({len(DIM_HEADERS)} dims + {len(PERIOD_HEADERS)} periods)")
