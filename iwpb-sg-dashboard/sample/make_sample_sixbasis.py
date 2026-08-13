"""
Generate a sample IWPB Singapore driller that carries every comparison the
pack writes commentary on, so all six bases appear in All Commentaries:

  (Current QTR) Actuals vs PY        prior-year months and quarter columns
  (Month YTD)   Actuals vs Target    a JUN YTD target column
  (Month)       Actuals vs Forecast  Jul-Dec forecast months
  (Month)  FY Forecast vs Prior FC   an FY-26 forecast column
  (Month)  FY Forecast vs Prior Year an FY-25 actuals column
  (Month)  FY Forecast vs Target     an FY-26 target column

The hierarchy is the real IWPB one — the MICA cascade under Revenue, the
product hierarchy under Loans, Deposits, Cards, Investments and Insurance —
so the commentary reads in the pack's own names. Figures are random: the
shape is real, the numbers are not.
"""
import random
from openpyxl import Workbook

random.seed(23)

DIM_HEADERS = [
    "MICA", "MICA_Level_3", "MICA_Level_2", "MICA_Level_1",
    "Product_code", "Product_Level_3", "Product_Level_2", "Product_Level_1",
    "Segment_code", "CG_Level_2", "CG_Level_1",
    "Function_code", "Function_Level_2", "Function_Level_1",
    "Entity code",
]

MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

# a band row above the header names what each column is, the way the real
# extracts do, so a target column is never mistaken for an actual
PERIODS = (
    [(f"{m}-25", "Actuals") for m in MONTHS] +
    [(f"{q}Q25 Actual", "Actuals") for q in range(1, 5)] +
    [(f"{m}-26", "Actuals") for m in MONTHS[:6]] +
    [(f"{m}-26", "Forecast") for m in MONTHS[6:]] +
    [(f"{q}Q26 Actual", "Actuals") for q in range(1, 3)] +
    [(f"{q}Q26 Actual", "Forecast") for q in range(3, 5)] +
    [("JUN YTD-25", "Actuals"), ("JUN YTD-26", "Actuals"),
     ("JUN YTD-26 Target", "Target"),
     ("FY-25", "Actuals"), ("FY-26 Forecast", "Forecast"), ("FY-26 Target", "Target")]
)

MICA_ROWS = [
    ("MP10101010000", "NII - Interest Income",   "NII - Net Interest Income", "Revenue", +1),
    ("MP10101020000", "NII - Interest Expense",  "NII - Net Interest Income", "Revenue", -1),
    ("MP10102010000", "NFI - Fee Income",        "Net Fee Income",            "Revenue", +1),
    ("MP10102020000", "NFI - Fee Expense",       "Net Fee Income",            "Revenue", -1),
    ("MP10103010000", "Trading Income - NII",    "Trading Income",            "Revenue", +1),
    ("MP10103020000", "Trading Income - Income", "Trading Income",            "Revenue", +1),
    ("MP10104010000", "Other Operating Income",  "Other Income",              "Revenue", +1),
]

PRODUCTS = [
    ("PR04000002", "Wealth Lending",                  "Loans"),
    ("PR04020000", "Mortgages (Real Estate Secured)", "Loans"),
    ("PR04030000", "Credit Cards",                    "Cards"),
    ("PR04060530", "Cash FX",                         "Investments"),
    ("PR04060600", "Structured Products",             "Investments"),
    ("PR05001000", "Current Accounts Total",          "Deposits"),
    ("PR05040000", "Savings and Deposit Accounts",    "Deposits"),
    ("PR12100000", "Insurance Distribution",          "Insurance"),
    ("PR04060550", "Other products",                  "Total Other Products"),
]

SEGMENT = ("CG1010000", "Retail Banking & Wealth Management", "IWPB")
FUNCTIONS = [
    ("RTN00008", "Premier and Wealth", "Intl Wealth & Premier Banking"),
    ("RTN00012", "Personal Banking",   "Intl Wealth & Premier Banking"),
]
ENTITY = "4040_1MGT"


def val(base, sign):
    return round(sign * abs(random.gauss(base, base * 0.25)), 1)


wb = Workbook()
ws = wb.active
ws.title = "Sheet1"
ws.append([""] * len(DIM_HEADERS) + [b for _, b in PERIODS])   # band row
ws.append(DIM_HEADERS + [h for h, _ in PERIODS])               # header row

for mica, l3, l2, l1, sign in MICA_ROWS:
    for pcode, p3, p2 in PRODUCTS:
        for fcode, f2, f1 in FUNCTIONS:
            if random.random() < 0.18:
                continue                       # sparse, as a real extract is
            base = random.uniform(4, 55)
            drift = random.uniform(0.88, 1.22)  # this year against last
            py = [val(base, sign) for _ in range(12)]
            pyq = [round(sum(py[i * 3:i * 3 + 3]), 1) for i in range(4)]
            cy = [val(base * drift, sign) for _ in range(6)]
            fc = [val(base * drift, sign) for _ in range(6)]
            cyq = [round(sum(cy[0:3]), 1), round(sum(cy[3:6]), 1),
                   round(sum(fc[0:3]), 1), round(sum(fc[3:6]), 1)]
            # the year-to-date columns are the file's own, close to but not
            # identical with the months, exactly as a real extract's are
            ytd_py = round(sum(py[:6]) * random.uniform(0.99, 1.01), 1)
            ytd_cy = round(sum(cy) * random.uniform(0.99, 1.01), 1)
            ytd_tgt = round(ytd_cy * random.uniform(0.94, 1.06), 1)
            fy25 = round(sum(py) * random.uniform(0.99, 1.01), 1)
            fy26fc = round((sum(cy) + sum(fc)) * random.uniform(0.99, 1.01), 1)
            fy26tgt = round(fy26fc * random.uniform(0.93, 1.07), 1)
            dims = [mica, l3, l2, l1, pcode, p3, p2, "Total Product",
                    SEGMENT[0], SEGMENT[1], SEGMENT[2], fcode, f2, f1, ENTITY]
            ws.append(dims + py + pyq + cy + fc + cyq +
                      [ytd_py, ytd_cy, ytd_tgt, fy25, fy26fc, fy26tgt])

# the month's written commentary travels inside the driller: a Commentary tab
# of bare paragraphs, matched against the driller's own names when read
COMMENTARY = [
    "NII carried the half. Deposit margin held better than the plan assumed as "
    "the pass-through on time deposits was slower than expected.",
    "Net Fee Income was the softer line. Investment activity thinned after the "
    "March rally and brokerage volumes have not come back, though the recurring "
    "wealth fee base is unchanged.",
    "Deposits ended the half broadly flat. The migration from current accounts "
    "into time deposits continued at much the same pace as the first quarter.",
    "Loans grew on mortgage completions that had been sitting in the pipeline "
    "since February.",
    "Cards spend was seasonally strong in June but the balance build is behind "
    "where the plan had it.",
    "Insurance new business held up against a soft market, helped by the "
    "bancassurance push in the second quarter.",
    "Trading Income was flat and is not expected to move materially in the "
    "second half.",
    "The half closed ahead of plan on revenue, with costs the watch item into "
    "the second half.",
]
wc = wb.create_sheet("Commentary")
for t in COMMENTARY:
    wc.append([t])

out = "IWPB_SG_Driller_sixbasis.xlsx"
wb.save(out)
print(f"wrote {out}: {ws.max_row-2} rows x {ws.max_column} cols "
      f"({len(DIM_HEADERS)} dims + {len(PERIODS)} periods)")
