"""
Config workbook carrying the reporting pack's own structures:
  AccountHierarchy  — P&L ex Notables cascade, Balance Sheet, Key Metrics
  CountryHierarchy  — Global > region > country, with Level 0 business line
plus every other section pre-filled from the app's defaults.
"""
import csv
from openpyxl import Workbook
from openpyxl.styles import Font

SRC = '/home/user/Myrep/iwpb-sg-dashboard/IWPB_SG_dashboard_config.csv'

HEADERS = {
    'Settings':         ['Key', 'Value'],
    'Calculations':     ['Key', 'Value'],
    'TilePriority':     ['Order', 'Pattern'],
    'Dimensions':       ['Key', 'Label', 'Filter (Y/N)', 'Sim scope (Y/N)', 'Match column'],
    'Views':            ['View', 'Show (Y/N)'],
    'Labels':           ['Key', 'Text'],
    'Drivers':          ['Key', 'Driver note'],
    'Metrics':          ['Key', 'Scope'],
    'Commentary':       ['View', 'Line', 'Text'],
    'ProductHierarchy': ['ID', 'Level 1', 'Level 2', 'Level 3', 'Level 4', 'Level 5', 'Match'],
}

sections = {}
with open(SRC, newline='', encoding='utf-8') as f:
    for row in csv.reader(f):
        if not row or row[0] == 'Sheet':
            continue
        sec, rest = row[0], row[1:]
        if sec in ('CountryHierarchy', 'AccountHierarchy'):
            continue                                  # rebuilt below
        if rest and str(rest[0]).startswith('<'):
            continue
        sections.setdefault(sec, []).append(rest)

# the Financial Summary cascades through the account hierarchy, so its
# Balance Sheet block rolls up into the same two lines the KPI tiles show
for r in sections.get('Settings', []):
    if r and r[0] == 'fsum_levels':
        r[1] = 'accH1,accH2,accH3,cgL2'
        break

# key metrics must never be totalled into the P&L — declare the patterns
for r in sections.get('Settings', []):
    if r and r[0] == 'memo_patterns':
        r[1] = (r'key metric|premier customer|\bcer\b|cost efficiency|\bfte\b|headcount|'
                r'\bnnm\b|\bnnd\b|\bnnia\b|net new money|invested asset|wealth balance')
        break

wb = Workbook()
wb.remove(wb.active)
bold = Font(bold=True)

def sheet(name, hdr, rows, widths):
    ws = wb.create_sheet(name)
    ws.append(hdr)
    for c in ws[1]:
        c.font = bold
    for r in rows:
        ws.append([x if x != '' else None for x in r])
    for i, w in enumerate(widths[:len(hdr)]):
        ws.column_dimensions[chr(65 + i)].width = w
    return ws

# the KPI tiles, scoped against the rebuilt account hierarchy — the pack's
# cost lines each get their own tile, Variable Pay and Indirect Costs included
sections['Metrics'] = [
    ['Revenue',                   'accH2=Revenue'],
    ['Banking NII',               'accH3=Banking NII'],
    ['Fees and Other Income',     'accH3=Fees and Other Income'],
    ['PBT',                       'accH1=PBT ex Notables'],
    ['Direct Cost ex VP, CC, GT', 'accH3=Direct Cost ex VP, CC, GT'],
    ['Variable Pay ex CC',        'accH3=Variable Pay ex CC'],
    ['Indirect Costs',            'accH3=Indirect Costs'],
    ['Operating Expenses',        'accH2=Operating Expenses'],
    ['ECL',                       'accH2=ECLs'],
    ['Deposits',                  'accH2=Deposits'],
    ['Loans and Advances',        'accH2=Loans and Advances'],
]

# the month's written commentary: bare paragraphs, matched against the
# driller's own names when read — the driller itself carries figures only
sections['Commentary'] = [[t] for t in [
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
]]

for sec, hdr in HEADERS.items():
    sheet(sec, hdr, sections.get(sec, []), [26, 52, 30, 30, 30, 30, 40])

# ---- AccountHierarchy: the driller's real MICA vocabulary, mapped onto the
# pack's line names. Level 1/2/3 values are taken from the driller pivot;
# the pack's o/w cuts (Retail, Private Bank, Wealth products) are segment and
# product splits, so they live in the cascade's later levels, not here.
A = []
def a(*levels, match=''):
    A.append([f'A{len(A)+1}', *list(levels) + [''] * (5 - len(levels)), match])

PBT, REV, FOI, OPEX = 'PBT ex Notables', 'Revenue', 'Fees and Other Income', 'Operating Expenses'
a(PBT, match='PBT')                                    # MICA L1
a(PBT, REV, match='Revenue')                           # MICA L2 (subtotal of the two below)
a(PBT, REV, 'Banking NII',
  match='NII - Net Interest Income')                   # MICA L3
a(PBT, REV, FOI,                                       # every other MICA L3 under Revenue
  match='Net Fee Income|Trading Income|Net Insurance Revenue|Operating Income|'
        'Gains less losses arising from derecognition of debt insts measured at FVOCI|'
        'Net inc/exp from other financial assets/financial liabilities DAFVTPL|'
        'Net income/expense from FA mandatorily measured at FVTPL')
a(PBT, 'ECLs',
  match='Expected credit losses|ECL-ICP|ECL-T/P')      # MICA L2 and its L3s
a(PBT, OPEX)                                           # subtotal of the three cost lines
a(PBT, OPEX, 'Direct Cost ex VP, CC, GT',
  match='Total Direct Cost (ex VP)')                   # MICA L3
a(PBT, OPEX, 'Variable Pay ex CC',
  match='Variable Pay')                                # MICA L3
a(PBT, OPEX, 'Indirect Costs',
  match='Total Indirect Costs|Inter Company Expenses|Total Indirect Costs (incl InterCo)')
BS = 'Balance Sheet'
a(BS)
a(BS, 'Deposits',           match='Customer Deposits') # MICA L1
a(BS, 'Loans and Advances', match='Loans and Advances') # MICA L1
KM = 'Key Metrics'
a(KM)
a(KM, 'Premier Customers (K)',           match='Premier Customers')
a(KM, 'CER (%)',                         match='CER|Cost Efficiency Ratio')
a(KM, 'Wealth Balances ($bn)',           match='Wealth Balances')
a(KM, 'Wealth Balances ($bn)', 'Wealth Deposits ($bn)',        match='Wealth Deposits')
a(KM, 'Wealth Balances ($bn)', 'Wealth Invested Assets ($bn)', match='Wealth Invested Assets|Invested Assets')
a(KM, 'NNM ($bn)',                       match='NNM|Net New Money')
a(KM, 'NNM ($bn)', 'NND ($bn)',          match='NND|Net New Deposits')
a(KM, 'NNM ($bn)', 'NNIA ($bn)',         match='NNIA|Net New Invested Assets')
a(KM, 'FTE (#)',                         match='FTE|Headcount')
sheet('AccountHierarchy',
      ['ID', 'Level 1', 'Level 2', 'Level 3', 'Level 4', 'Level 5', 'Match'],
      A, [8, 20, 22, 30, 26, 26, 62])

# ---- CountryHierarchy: Global > region > country, Level 0 business line ----
ASIA = ['Australia', 'Malaysia', 'Taiwan', 'India', 'Indonesia', 'Vietnam', 'Sri Lanka',
        'Philippines', 'Bangladesh', 'Japan', 'AMH', 'HASE', 'China', 'Singapore']
AE   = ['US', 'Mexico', 'France', 'CIIOM', 'Switzerland', 'Bermuda', 'Malta',
        'Germany', 'Luxembourg', 'UK NRFB']
C, n = [], 0
def c(l0, l1, l2, l3):
    global n
    n += 1
    C.append([f'C{n}', l0, l1, l2, l3])
c('IWPB', 'Global', '', '')
for x in ASIA: c('IWPB', 'Global', 'Asia', x)
c('IWPB', 'Global', 'UK', 'UK PB')
for x in ['UAE', 'Egypt', 'Turkiye', 'Qatar']: c('IWPB', 'Global', 'MENAT', x)
for x in AE: c('IWPB', 'Global', 'Americas and Europe', x)
c('IWPB', 'Global', 'Holdings', '')
for x in ASIA: c('CIB', 'Global', 'Asia', x)
for x in ['UAE', 'Egypt', 'Turkiye']: c('CIB', 'Global', 'MENAT', x)
c('CIB', 'Global', 'Americas and Europe', 'US')
c('GROUP', 'GROUP', 'Holdings', '')
sheet('CountryHierarchy', ['ID', 'Level 0', 'Level 1', 'Level 2', 'Level 3'],
      C, [10, 12, 12, 24, 22])

out = 'IWPB_dashboard_config_pack.xlsx'
wb.save(out)
print('wrote', out, '| account rows:', len(A), '| country rows:', len(C))
