"""
A CIB Asia & MENAT driller in the shape the pack page implies: every printed
line as a leaf, split by country, segment and income type, with the June YTD
figures taken from the pack itself so the page ties line for line.
"""
from openpyxl import Workbook

MON = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']

# line, MICA L1 band, MICA L2, MICA L3, MICA L4, YTD(Jun), Apr, May, Jun, segment
# amounts in US$m, costs and charges negative, exactly as the pack prints them
PL = [
 # ---- revenue, ex notables ----
 ('PBT (ex Notables)','Total Revenue (ex Notables)','MSS','',                    3859,  643, 604, 718, 'Financial Institutions'),
 ('PBT (ex Notables)','Total Revenue (ex Notables)','Credit and Lending','Portfolio Management',
                                                                                   -1,    0,   1,  -1, 'Corporate'),
 ('PBT (ex Notables)','Total Revenue (ex Notables)','Credit and Lending','Corporate Lending',
                                                                                  626,   96, 109,  95, 'Corporate'),
 ('PBT (ex Notables)','Total Revenue (ex Notables)','Global Trade Solutions','',   754,  128, 121, 126, 'Corporate'),
 ('PBT (ex Notables)','Total Revenue (ex Notables)','HIF','',                      144,   16,  16,  41, 'Corporate'),
 ('PBT (ex Notables)','Total Revenue (ex Notables)','Global Payments Solutions','',2293,  378, 387, 388, 'Corporate'),
 ('PBT (ex Notables)','Total Revenue (ex Notables)','CMA Net','',                  344,   74,  40,  45, 'Financial Institutions'),
 ('PBT (ex Notables)','Total Revenue (ex Notables)','Other Revenue','',            728,  130, 114, 139, 'Corporate'),
 # ---- charge ----
 ('PBT (ex Notables)','ECLs (ex Notables)','ECLs (ex Notables)','',              -298,   -6, -26,-117, 'Corporate'),
 # ---- costs ----
 ('PBT (ex Notables)','Total Operating Expenses (ex Notables)','Direct Costs (ex VP)','Direct Costs (ex VP and Litigation)',
                                                                                 -981, -170,-166,-171, 'Corporate'),
 ('PBT (ex Notables)','Total Operating Expenses (ex Notables)','Direct Costs (ex VP)','Litigation',
                                                                                    0,    0,   0,   0, 'Corporate'),
 ('PBT (ex Notables)','Total Operating Expenses (ex Notables)','Variable Pay (VP)','',
                                                                                 -307,  -58, -39, -56, 'Corporate'),
 ('PBT (ex Notables)','Total Operating Expenses (ex Notables)','Indirect Costs','',
                                                                                -2082, -336,-348,-362, 'Corporate'),
 # ---- notables, the bridge to reported ----
 ('Notables','Revenue Notables','Revenue Notables','',                              1,    0,   0,   0, 'Corporate'),
 ('Notables','Opex Notables','Opex Notables','',                                  -27,   -5,  -3,  -9, 'Corporate'),
]

# the Reported revenue split: how much of each revenue line is banking NII
NII = {'MSS':900, 'Portfolio Management':-1, 'Corporate Lending':561,
       'Global Trade Solutions':470, 'HIF':60, 'Global Payments Solutions':2000,
       'CMA Net':60, 'Other Revenue':608}

# balance sheet, closing balances carried in the file's unit (US$m) — the app
# restates them to US$bn on the page; the wider line is the parent
BSB = [
 ('Balance Sheet','Customers and Banks Deposits (PE)','Customer Deposits (PE)','',      287, 265, 282, 287),
 ('Balance Sheet','Customers and Banks Deposits (PE)','Bank Deposits (PE)','',           20,  19,  20,  20),
 ('Balance Sheet','Loans and Advances to Customers and Banks (PE)','Loans and Advances to Customers (PE)','',
                                                                                        200, 193, 194, 200),
 ('Balance Sheet','Loans and Advances to Customers and Banks (PE)','Loans and Advances to Banks (PE)','',
                                                                                         58,  56,  54,  58),
 ('Balance Sheet','Total RWAs','Total RWAs','',                                         233, 228, 227, 233),
]
BS = [(a,b,c,d, *[x*1000 for x in rest]) for a,b,c,d,*rest in BSB]
# key metrics — never totalled, carried for information
KM = [
 ('Key Metrics','Headcount','Headcount','',                                14186, 14240, 14173, 14186),
 ('Key Metrics','Reported RoTE (%)','Reported RoTE (%)','',                 20.9,  21.4,  20.6,  20.9),
 ('Key Metrics','CER ex Notables (%)','CER ex Notables (%)','',             38.5,  38.5,  39.7,  38.0),
 ('Key Metrics','ECLs / Loans and Advances (%)','ECLs / Loans and Advances (%)','',
                                                                            0.31,  0.05,  0.28,  0.19),
]

CTRY = [('Hong Kong','Asia',0.30), ('Singapore','Asia',0.20), ('India','Asia',0.15),
        ('China','Asia',0.15), ('UAE','MENAT',0.12), ('Egypt','MENAT',0.08)]

def split(total):
    """weights that add back to the total exactly"""
    out, run = [], 0.0
    for i,(c,r,w) in enumerate(CTRY):
        v = round(total*w, 3) if i < len(CTRY)-1 else round(total-run, 3)
        run = round(run+v, 3); out.append((c,r,v))
    return out

wb = Workbook(); ws = wb.active; ws.title = 'CIB Asia and MENAT'
HEAD = ['MICA','MICA_Level_4','MICA_Level_3','MICA_Level_2','MICA_Level_1',
        'Income_Type','Product_Level_2','Product_Level_1','Segment_code','CG_Level_2','CG_Level_1',
        'Country','Region','Business_Line','Entity code']
PERIOD = [f'{m}-26' for m in MON] + ['Jun YTD-26','Jun YTD-25','Jun YTD-26 Target',
                                     'FY-26 Forecast','FY-26 Target']
band = ['']*len(HEAD) + ['Actuals']*6 + ['Forecast']*6 + ['Actuals','Actuals','Target','Forecast','Target']
ws.append(band)
ws.append(HEAD + PERIOD)

mi = 0
def emit(l1, l2, l3, l4, ytd, apr, may, jun, seg, income, balance=False, memo=False):
    """one printed line, split across the footprint"""
    global mi
    if l3 == '' or ytd is None: return
    rest = ytd - (apr+may+jun)
    if balance or memo:
        months = [round(ytd*(0.94+0.01*i), 3) for i in range(3)] + [apr,may,jun] + \
                 [round(ytd*(1+0.004*(i+1)), 3) for i in range(6)]
        fy_fc, fy_tgt, py = ytd*1.02, ytd*1.03, ytd*0.95
        ytd_tgt = ytd*1.01
    else:
        months = [round(rest/3, 3)]*3 + [apr,may,jun] + [round(jun*(1+0.01*(i+1)), 3) for i in range(6)]
        fy_fc = round(ytd + sum(months[6:]), 3)
        fy_tgt, py, ytd_tgt = fy_fc*1.03, ytd*0.93, ytd*1.02
    for c, r, w in split(1.0):
        mi += 1
        vals = [round(v*w, 4) for v in months] + \
               [round(ytd*w,4), round(py*w,4), round(ytd_tgt*w,4), round(fy_fc*w,4), round(fy_tgt*w,4)]
        ws.append([f'C{mi:04d}', l4, l3, l2, l1, income,
                   l3 if l1.startswith('PBT') else l2, 'CIB', 'SEG'+c[:2].upper(),
                   seg, 'CIB', c, r, 'CIB', 'CIB'+c[:2].upper()] + vals)

for l1, l2, l3, l4, ytd, apr, may, jun, seg in PL:
    leaf = l4 or l3
    if l2 == 'Total Revenue (ex Notables)':
        nii = NII.get(leaf, 0)
        fee = ytd - nii
        share = (lambda v: 0 if ytd == 0 else v/ytd)
        emit(l1, l2, l3, l4, nii,  round(apr*share(nii),3),  round(may*share(nii),3),  round(jun*share(nii),3),
             seg, 'Banking NII')
        emit(l1, l2, l3, l4, fee,  round(apr*share(fee),3),  round(may*share(fee),3),  round(jun*share(fee),3),
             seg, 'Fee and other income')
    else:
        inc = 'Fee and other income' if l2 == 'Revenue Notables' else ''
        emit(l1, l2, l3, l4, ytd, apr, may, jun, seg, inc)
for l1, l2, l3, l4, ytd, apr, may, jun in BS:
    emit(l1, l2, l3, l4, ytd, apr, may, jun, 'Corporate', '', balance=True)
for l1, l2, l3, l4, ytd, apr, may, jun in KM:
    emit(l1, l2, l3, l4, ytd, apr, may, jun, 'Corporate', '', memo=True)

for col, w in zip('ABCDEFGHIJKLMNO', [8,30,30,38,22,20,26,12,12,22,10,12,10,12,10]):
    ws.column_dimensions[col].width = w
out = 'CIB_AME_Jun26_Driller.xlsx'
wb.save(out)
print('wrote', out, '| rows:', ws.max_row-2)
