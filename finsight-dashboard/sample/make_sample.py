"""Synthetic IWPB-style TM1 extract for testing the FinSight dashboard.

One tab, headers in row 1: the dimension columns the parser looks for
(MICA / Product / Segment / Function hierarchies plus Country, Region,
Business_Line, Entity code), then PY and CY monthly columns, quarters, YTD,
and the FY actual / forecast / target columns. Figures are US$m.
"""
import random, sys
from openpyxl import Workbook

random.seed(7)
out = sys.argv[1] if len(sys.argv) > 1 else "IWPB_SG_sample.xlsx"
COUNTRIES = [("Singapore","ASEAN",1.0)] if len(sys.argv) < 3 else [("Singapore","ASEAN",1.0),("Hong Kong","North Asia",1.6)]

MON = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
DIMS = ["Country","Region","Business_Line","MICA","MICA_Level_3","MICA_Level_2","MICA_Level_1",
        "Product_code","Product_Level_3","Product_Level_2","Product_Level_1",
        "Segment_code","CG_Level_2","CG_Level_1","Function_code","Function_Level_2","Function_Level_1","Entity code"]

# (L1, L2, L3, base monthly size, sign)
MICA = [
    ("Revenue","Net interest income","NII - Deposits",      38, 1),
    ("Revenue","Net interest income","NII - Loans",         22, 1),
    ("Revenue","Fee income",         "Investment fees",     14, 1),
    ("Revenue","Fee income",         "Insurance fees",       9, 1),
    ("Revenue","Fee income",         "Card fees",            6, 1),
    ("Revenue","Other income",       "FX income",            5, 1),
    ("Costs",  "Staff costs",        "Salaries",           -21, 1),
    ("Costs",  "Staff costs",        "Bonus",               -6, 1),
    ("Costs",  "Non-staff costs",    "Premises",            -5, 1),
    ("Costs",  "Non-staff costs",    "Technology",          -8, 1),
    ("Costs",  "Non-staff costs",    "Marketing",           -3, 1),
    ("ECL",    "Expected credit losses","ECL charge",        -4, 1),
]
PROD = [
    ("Wealth","Investments","Funds","P100"), ("Wealth","Investments","Structured products","P101"),
    ("Wealth","Insurance","Life insurance","P200"),
    ("Personal Banking","Deposits","Savings","P300"), ("Personal Banking","Deposits","Current accounts","P301"),
    ("Personal Banking","Mortgages","Home loans","P400"), ("Personal Banking","Cards","Credit cards","P500"),
]
SEG = [("Customer groups","Premier","S1"), ("Customer groups","Private Bank","S2"), ("Customer groups","Personal","S3")]
FUNC = [("Functions","Frontline","F1"), ("Functions","Operations","F2")]

# period headers
py = [f"{m}-25 Actual" for m in MON]
cy = [f"{m}-26 Actual" for m in MON[:6]] + [f"{m}-26 Forecast" for m in MON[6:]]
qtr = ["1Q25","2Q25","3Q25","4Q25","1Q26","2Q26"]
ytd = ["YTD Jun-25","YTD Jun-26","YTD Jun-26 Target"]
fy = ["FY-25 Actual","FY-26 Forecast","FY-26 Target"]
HEAD = DIMS + py + cy + qtr + ytd + fy

wb = Workbook(); ws = wb.active; ws.title = "Singapore"
ws.append(HEAD)

for (ctry,reg,cm) in COUNTRIES:
  for (l1,l2,l3,base,sgn) in MICA:
    for (pl1,pl2,pl3,pc) in PROD:
        for (cg1,cg2,sc) in SEG:
            fn = random.choice(FUNC)
            # the product mix weights the line so hierarchies look different
            w = {"Wealth":1.3,"Personal Banking":0.8}[pl1] * {"Premier":1.0,"Private Bank":1.6,"Personal":0.5}[cg2]
            if l2=="Net interest income" and pl2 in ("Investments","Insurance"): w*=0.15
            if l2=="Fee income" and pl2 in ("Deposits","Mortgages"): w*=0.2
            b = base*w*cm/6.0
            pyv = [round(b*random.uniform(0.85,1.15),2) for _ in MON]
            cyv = [round(b*1.06*random.uniform(0.85,1.15),2) for _ in MON]
            q = [sum(pyv[0:3]),sum(pyv[3:6]),sum(pyv[6:9]),sum(pyv[9:12]),sum(cyv[0:3]),sum(cyv[3:6])]
            y = [sum(pyv[:6]), sum(cyv[:6]), sum(cyv[:6])*1.04]
            f = [sum(pyv), sum(cyv), sum(cyv)*1.05]
            row = [ctry,reg,"IWPB", f"M{MICA.index((l1,l2,l3,base,sgn))+1:03d}", l3,l2,l1,
                   pc,pl3,pl2,pl1, sc,cg2,cg1, fn[2],fn[1],fn[0], "HBAP-SG"]
            row += pyv + cyv + [round(v,2) for v in q+y+f]
            ws.append(row)

wb.save(out)
print("wrote", out, "rows", ws.max_row-1, "cols", len(HEAD))
