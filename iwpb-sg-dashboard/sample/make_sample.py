"""
Generate a sample IWPB Singapore driller extract that mirrors the real file's
column structure, so the dashboard can be tested without exposing real data.

Structure (single header row on sheet 'Sheet1'):
  Dimension columns -> MICA ... Entity code   (15 columns)
  Period columns    -> Jan..Dec Actual (prior year), 1Q25..4Q25 Actual,
                       Jan..Jun Actual (current year), Jul FC,
                       JUN YTD Actuals 2026, 1Q26..4Q26 Actual
"""
import random
from openpyxl import Workbook

random.seed(7)

DIM_HEADERS = [
    "MICA", "MICA_Level_3", "MICA_Level_2", "MICA_Level_1",
    "Product_code", "Product_Level_3", "Product_Level_2", "Product_Level_1",
    "Segment_code", "CG_Level_2", "CG_Level_1",
    "Function_code", "Function_Level_2", "Function_Level_1",
    "Entity code",
]

MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

PERIOD_HEADERS = (
    [f"{m} Actual" for m in MONTHS] +                    # prior year (2025)
    [f"{q}Q25 Actual" for q in range(1, 5)] +
    [f"{m} Actual" for m in MONTHS[:6]] +                # current year (2026)
    ["Jul FC", "JUN YTD Actuals 2026"] +
    [f"{q}Q26 Actual" for q in range(1, 5)]
)

# MICA hierarchy: code -> (Level 3, Level 2, Level 1), sign of typical values
MICA_ROWS = [
    ("MP10101010000", "NII - Interest Income",      "NII - Net Interest Income", "Revenue", +1),
    ("MP10101020000", "NII - Interest Expense",     "NII - Net Interest Income", "Revenue", -1),
    ("MP10102010000", "NFI - Fee Income",           "Net Fee Income",            "Revenue", +1),
    ("MP10102020000", "NFI - Fee Expense",          "Net Fee Income",            "Revenue", -1),
    ("MP10103010000", "Trading Income - NII",       "Trading Income",            "Revenue", +1),
    ("MP10103020000", "Trading Income - Income",    "Trading Income",            "Revenue", +1),
    ("MP10104010000", "Other Operating Income",     "Other Income",              "Revenue", +1),
]

# Product hierarchy: code -> (Level 3, Level 2); Level 1 is always Total Product
PRODUCTS = [
    ("UNMAPPDT",   "Unmapped Product",                     "Unmapped Product"),
    ("PR04000002", "Wealth Lending",                       "Loans"),
    ("PR04004001", "Other Lending",                        "Loans"),
    ("PR04020000", "Mortgages (Real Estate Secured)",      "Loans"),
    ("PR04030000", "Credit Cards",                         "Cards"),
    ("PR04060440", "Equity Linked Investments",            "Investments"),
    ("PR04060460", "HSBC Investments - Other",             "Investments"),
    ("PR04060510", "Securities Trading - Equities",        "Investments"),
    ("PR04060520", "Securities Trading - Bonds",           "Investments"),
    ("PR04060530", "Cash FX",                              "Investments"),
    ("PR04060600", "Structured Products",                  "Investments"),
    ("PR04060700", "Collective Investment Schemes",        "Investments"),
    ("PR05001000", "Current Accounts Total",               "Deposits"),
    ("PR05040000", "Savings and Deposit Accounts",         "Deposits"),
    ("PR04060550", "Other products",                       "Total Other Products"),
    ("PR12100000", "Insurance Distribution",               "Insurance"),
    ("PR04060560", "Non Product Results",                  "Non Product Results Total"),
    ("PR10090005", "Non Product Allocation of Corp Centre", "Non Product Results Total"),
    ("PR13010000", "Supplementary Other Non Product",      "Non Product Results Total"),
]

SEGMENT = ("CG1010000", "Retail Banking & Wealth Management", "IWPB")
FUNCTIONS = [
    ("RTN00008", "Premier and Wealth", "Intl Wealth & Premier Banking"),
    ("RTN00012", "Personal Banking",   "Intl Wealth & Premier Banking"),
]
ENTITIES = ["4040_1MGT"]


def gen_value(base, sign):
    r = random.random()
    if r < 0.30:
        return 0
    v = round(random.gauss(base, base * 0.4))
    v = sign * abs(v)
    # real extracts render some negatives as "(n)" text
    if v < 0 and random.random() < 0.25:
        return f"({abs(v):,})"
    return v


wb = Workbook()
ws = wb.active
ws.title = "Sheet1"
ws.append(DIM_HEADERS + PERIOD_HEADERS)

n_month_cols = 12 + 4 + 6 + 2 + 4  # index bookkeeping below relies on this order

for mica, l3, l2, l1, sign in MICA_ROWS:
    for pcode, p3, p2 in PRODUCTS:
        for fcode, f2, f1 in FUNCTIONS:
            for entity in ENTITIES:
                if random.random() < 0.25:
                    continue  # sparse: not every combination exists
                base = random.uniform(2, 60)
                growth = random.uniform(0.85, 1.25)  # CY vs PY drift
                py = [gen_value(base, sign) for _ in range(12)]
                pyq = [gen_value(base * 3, sign) for _ in range(4)]
                cy = [gen_value(base * growth, sign) for _ in range(6)]
                fc = [gen_value(base * growth, sign)]
                ytd = [gen_value(base * growth * 6, sign)]
                cyq = [gen_value(base * growth * 3, sign) for _ in range(2)] + [0, 0]
                dims = [mica, l3, l2, l1, pcode, p3, p2, "Total Product",
                        SEGMENT[0], SEGMENT[1], SEGMENT[2],
                        fcode, f2, f1, entity]
                ws.append(dims + py + pyq + cy + fc + ytd + cyq)

out = "IWPB_SG_Driller_sample.xlsx"
wb.save(out)
print(f"wrote {out}: {ws.max_row-1} rows x {ws.max_column} cols "
      f"({len(DIM_HEADERS)} dims + {len(PERIOD_HEADERS)} periods)")
