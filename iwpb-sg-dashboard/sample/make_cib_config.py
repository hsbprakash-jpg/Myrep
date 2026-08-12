"""
The CIB Asia & MENAT pack page, expressed as a config workbook:
  AccountHierarchy — PBT (Reported) = PBT (ex Notables) + Notables, with the
                     revenue engines, the cost stack, the balance sheet and
                     the key metrics beneath, so every printed line ties
  Metrics          — the KPI tiles the page leads with
  CountryHierarchy — CIB's Asia and MENAT footprint
"""
from openpyxl import Workbook
from openpyxl.styles import Font

wb = Workbook(); wb.remove(wb.active)
bold = Font(bold=True)

def sheet(name, hdr, rows, widths):
    ws = wb.create_sheet(name)
    ws.append(hdr)
    for c in ws[1]: c.font = bold
    for r in rows: ws.append([x if x != '' else None for x in r])
    for i, w in enumerate(widths[:len(hdr)]):
        ws.column_dimensions[chr(65+i)].width = w
    return ws

# ---------------- Settings ----------------
SET = [
 ('app_title',            'CIB Asia & MENAT Driller'),
 ('app_subtitle',         'Management reporting · KPI insights'),
 ('eyebrow',              'HSBC Management Reporting : CIB Asia & MENAT'),
 ('export_prefix',        'CIB_AME'),
 ('logo_mark',            'CIB'),
 ('landing_title',        'Financial Performance'),
 ('fsum_title',           'CIB Asia & MENAT: Financial Summary'),
 ('fsum_section_label',   'P&L'),
 ('fsum_bs_label',        'Balance Sheet'),
 ('fsum_memo_label',      'Key Metrics'),
 ('fsum_levels',          'accH1,accH2,accH3,accH4,accH5,cgL2,country'),
 ('fsum_default_levels',  'accH1,accH2,accH3,accH4'),
 ('fsum_ow',              'N'),
 ('fsum_residual_label',  'Other'),
 ('fsum_parent_row',      'top'),
 ('unit_label',           'US$m'),
 ('bs_unit',              'bn'),
 ('commentary_note',      'ex notables'),
 ('commentary_levels',    'accH1,accH2,accH3'),
 ('tile_line_levels',     'accH1,accH2,accH3,accH4,accH5'),
 ('tile_dims',            'accH1,accH2'),
 # the four ratios and the headcount are never added into a P&L total
 ('memo_patterns',        r'key metric|headcount|\brote\b|\bcer\b|cost efficiency|'
                          r'ecls?\s*/\s*loans|\bbps\b'),
 # the smaller revenue engines start parked; the ✕/↺ on a tile moves them
 ('tile_hidden',          'Credit and Lending;Global Trade Solutions;HIF;CMA Net;'
                          'Other Revenue;Total RWAs'),
 ('calc_ytd_actuals',     'column'),
 ('show_greeting',        'N'),
]
sheet('Settings', ['Key','Value'], SET, [26, 66])

# ---------------- Calculations / Views / Labels / Drivers ----------------
sheet('Calculations', ['Key','Value'], [], [26, 52])
sheet('TilePriority', ['Order','Pattern'], [], [10, 40])
sheet('Dimensions', ['Key','Label','Filter (Y/N)','Sim scope (Y/N)','Match column'],
      [['income','Income type','Y','Y','Income_Type']], [16, 22, 16, 16, 22])
sheet('Views', ['View','Show (Y/N)'],
      [[v,'Y'] for v in ['summary','fsum','custom','builder','query','table','assist','sim']], [16, 14])
sheet('Labels', ['Key','Text'],
      [['nav_summary','KPI summary'], ['nav_fsum','Financial Summary'],
       ['sect_kpis','KPIs · CIB lines']], [22, 30])
sheet('Drivers', ['Key','Driver note'], [], [26, 60])

# ---------------- Metrics: the KPI tiles ----------------
MET = [
 ('Total Revenue (ex Notables)',            'accH3=Total Revenue (ex Notables)'),
 ('MSS',                                    'accH4=MSS'),
 ('Global Payments Solutions',              'accH4=Global Payments Solutions'),
 ('Credit and Lending',                     'accH4=Credit and Lending'),
 ('Global Trade Solutions',                 'accH4=Global Trade Solutions'),
 ('HIF',                                    'accH4=HIF'),
 ('CMA Net',                                'accH4=CMA Net'),
 ('Other Revenue',                          'accH4=Other Revenue'),
 ('Banking NII',                            'income=Banking NII'),
 ('Fee and other income',                   'income=Fee and other income'),
 ('ECLs (ex Notables)',                     'accH3=ECLs (ex Notables)'),
 ('Direct Costs (ex VP)',                   'accH4=Direct Costs (ex VP)'),
 ('Variable Pay (VP)',                      'accH4=Variable Pay (VP)'),
 ('Indirect Costs',                         'accH4=Indirect Costs'),
 ('Total Operating Expenses (ex Notables)', 'accH3=Total Operating Expenses (ex Notables)'),
 ('PBT (ex Notables)',                      'accH2=PBT (ex Notables)'),
 ('PBT (Reported)',                         'accH1=PBT (Reported)'),
 ('Customer Deposits (PE)',                 'accH2=o/w Customer Deposits (PE)'),
 ('Loans and Advances to Customers (PE)',   'accH2=o/w Loans and Advances to Customers (PE)'),
 ('Total RWAs',                             'accH1=Total RWAs'),
]
sheet('Metrics', ['Key','Scope'], MET, [42, 48])

sheet('ProductHierarchy', ['ID','Level 1','Level 2','Level 3','Level 4','Level 5','Match'],
      [], [8, 20, 22, 24, 22, 22, 40])

# ---------------- AccountHierarchy ----------------
# a node with no Match is a subtotal: it claims no rows and equals its children
A = []
def a(*levels, match=''):
    A.append([f'A{len(A)+1}', *list(levels) + ['']*(5-len(levels)), match])

REP, EXN, REV = 'PBT (Reported)', 'PBT (ex Notables)', 'Total Revenue (ex Notables)'
OPX, NOT = 'Total Operating Expenses (ex Notables)', 'Notables'
a(REP)                                                    # = ex notables + notables
a(REP, EXN)                                               # = revenue + ECLs + opex
a(REP, EXN, REV)                                          # = the revenue engines
a(REP, EXN, REV, 'MSS',                       match='MSS')
a(REP, EXN, REV, 'Credit and Lending')                    # printed with its o/w line
a(REP, EXN, REV, 'Credit and Lending', 'o/w Portfolio Management', match='Portfolio Management')
a(REP, EXN, REV, 'Credit and Lending', 'o/w Corporate Lending', match='Corporate Lending')
a(REP, EXN, REV, 'Global Trade Solutions',    match='Global Trade Solutions|GTS')
a(REP, EXN, REV, 'HIF',                       match='HIF')
a(REP, EXN, REV, 'Global Payments Solutions', match='Global Payments Solutions|GPS')
a(REP, EXN, REV, 'CMA Net',                   match='CMA Net')
a(REP, EXN, REV, 'Other Revenue',             match='Other Revenue')
a(REP, EXN, 'ECLs (ex Notables)',             match='ECLs (ex Notables)|ECL (ex Notables)')
a(REP, EXN, OPX)                                          # = direct + VP + indirect
a(REP, EXN, OPX, 'Direct Costs (ex VP)')                  # printed with its two parts
a(REP, EXN, OPX, 'Direct Costs (ex VP)', 'o/w Direct Costs (ex VP and Litigation)',
  match='Direct Costs (ex VP and Litigation)')
a(REP, EXN, OPX, 'Direct Costs (ex VP)', 'o/w Litigation', match='Litigation')
a(REP, EXN, OPX, 'Variable Pay (VP)',         match='Variable Pay (VP)|Variable Pay')
a(REP, EXN, OPX, 'Indirect Costs',            match='Indirect Costs')
a(REP, NOT)                                               # the bridge to reported
a(REP, NOT, 'Revenue Notables',               match='Revenue Notables')
a(REP, NOT, 'Opex Notables',                  match='Opex Notables')
# deposits, loans and RWAs have no common parent — a node above them would
# print a total that adds a deposit book to a risk-weighted asset, so each
# stands at the top of its own branch and the band forms from the statement
a('Customers and Banks Deposits (PE)')
a('Customers and Banks Deposits (PE)', 'o/w Customer Deposits (PE)', match='Customer Deposits (PE)')
a('Customers and Banks Deposits (PE)', 'o/w Bank Deposits (PE)',     match='Bank Deposits (PE)')
a('Loans and Advances to Customers and Banks (PE)')
a('Loans and Advances to Customers and Banks (PE)', 'o/w Loans and Advances to Customers (PE)',
  match='Loans and Advances to Customers (PE)')
a('Loans and Advances to Customers and Banks (PE)', 'o/w Loans and Advances to Banks (PE)',
  match='Loans and Advances to Banks (PE)')
a('Total RWAs',                    match='Total RWAs|RWAs')
# the key metrics are ratios and a count: each stands alone, none is added
a('Headcount',                     match='Headcount|FTE')
a('Reported RoTE (%)',             match='Reported RoTE (%)|RoTE')
a('CER ex Notables (%)',           match='CER ex Notables (%)|CER')
a('ECLs / Loans and Advances (%)', match='ECLs / Loans and Advances (%)')
sheet('AccountHierarchy', ['ID','Level 1','Level 2','Level 3','Level 4','Level 5','Match'],
      A, [8, 18, 22, 38, 30, 34, 46])

# ---------------- CountryHierarchy ----------------
ASIA = ['Hong Kong','Singapore','India','China','Australia','Japan','Taiwan','Malaysia',
        'Indonesia','Vietnam','Bangladesh','Philippines','Sri Lanka','Korea','Thailand']
MENAT = ['UAE','Egypt','Turkiye','Qatar','Saudi Arabia','Bahrain','Kuwait','Oman']
C, n = [], 0
def c(l0, l1, l2, l3):
    global n; n += 1; C.append([f'C{n}', l0, l1, l2, l3])
c('CIB','Global','','')
for x in ASIA:  c('CIB','Global','Asia', x)
for x in MENAT: c('CIB','Global','MENAT', x)
for x in ['US','UK','France','Germany','Mexico']: c('CIB','Global','Americas and Europe', x)
c('GROUP','GROUP','Holdings','')
sheet('CountryHierarchy', ['ID','Level 0','Level 1','Level 2','Level 3'], C, [10, 12, 12, 24, 22])

out = 'CIB_dashboard_config_pack.xlsx'
wb.save(out)
print('wrote', out, '| account rows:', len(A), '| metrics:', len(MET), '| countries:', len(C))
