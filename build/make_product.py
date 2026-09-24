# Generates product/Tidy-Money-Budget-Dashboard.xlsx
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import DataBarRule, CellIsRule, FormulaRule
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.series import SeriesLabel
from openpyxl.workbook.defined_name import DefinedName
import datetime as dt

OUT = "product/Tidy-Money-Budget-Dashboard.xlsx"
YEAR = 2026

# ---------- palette ----------
INK = "1F3D2B"; SAGE = "7FA98B"; SAGE_L = "E4EFE7"; CREAM = "FAF7F2"; MUST = "E3B23C"
GREY = "6B6B6B"; LINE = "D9D4CB"; WHITE = "FFFFFF"; RED_L = "F6DADA"; GREEN_L = "DDEFE0"
FONT = "Arial"

def fill(c): return PatternFill("solid", start_color=c, end_color=c)
thin = Side(style="thin", color=LINE)
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
BOTTOM = Border(bottom=Side(style="thin", color=SAGE))

def f(bold=False, size=10, color="1B1B1B", italic=False):
    return Font(name=FONT, bold=bold, size=size, color=color, italic=italic)

CUR = '"$"#,##0.00;[Red]-"$"#,##0.00;"-"'
CUR0 = '"$"#,##0;[Red]-"$"#,##0;"-"'
PCT = '0%'
DATEF = 'mmm d, yyyy'
MONTHF = 'mmmm yyyy'

wb = Workbook()

def title(ws, text, sub=None, width_cols=8):
    ws["B2"] = text
    ws["B2"].font = Font(name=FONT, bold=True, size=18, color=INK)
    if sub:
        ws["B3"] = sub
        ws["B3"].font = f(size=10, color=GREY, italic=True)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 2

def header_row(ws, row, cols, start_col=2, fill_color=INK, font_color=WHITE):
    for i, h in enumerate(cols):
        c = ws.cell(row=row, column=start_col + i, value=h)
        c.font = Font(name=FONT, bold=True, size=10, color=font_color)
        c.fill = fill(fill_color)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
    ws.row_dimensions[row].height = 22

def box(ws, rng, color=WHITE):
    for row in ws[rng]:
        for c in row:
            c.border = BORDER
            if color: c.fill = fill(color)

def input_hint(cell):
    cell.fill = fill("FFF9E6")  # soft yellow = you type here

# =====================================================================
# SETTINGS
# =====================================================================
st = wb.active
st.title = "Settings"
st.sheet_properties.tabColor = GREY
title(st, "Settings", "Edit these lists once. Every dropdown in the workbook updates automatically.")
st["B5"] = "Budget year"; st["B5"].font = f(bold=True)
st["C5"] = YEAR; st["C5"].font = f(bold=True, size=12, color=INK); input_hint(st["C5"]); st["C5"].border = BORDER
st["D5"] = "← change the year and every month in the workbook follows"; st["D5"].font = f(color=GREY, italic=True)

header_row(st, 7, ["Expense categories"], start_col=2)
header_row(st, 7, ["Income categories"], start_col=4)
header_row(st, 7, ["Accounts"], start_col=6)
header_row(st, 7, ["Months (auto)"], start_col=8)

expense_cats = ["Housing", "Utilities", "Groceries", "Eating out", "Transport", "Insurance", "Health",
                "Subscriptions", "Phone & internet", "Shopping", "Personal care", "Kids & family",
                "Pets", "Gifts & giving", "Fun & hobbies", "Travel", "Education", "Debt payments",
                "Savings transfers", "Other"]
income_cats = ["Salary", "Partner salary", "Side income", "Refunds", "Other income"]
accounts = ["Checking", "Savings", "Credit card", "Cash"]

for i in range(30):
    r = 8 + i
    for col in (2, 4, 6):
        c = st.cell(row=r, column=col); c.border = BORDER; input_hint(c)
    if i < len(expense_cats): st.cell(row=r, column=2, value=expense_cats[i])
    if i < len(income_cats): st.cell(row=r, column=4, value=income_cats[i])
    if i < len(accounts): st.cell(row=r, column=6, value=accounts[i])
for m in range(12):
    c = st.cell(row=8 + m, column=8, value=f"=DATE($C$5,{m+1},1)")
    c.number_format = MONTHF; c.border = BORDER; c.fill = fill(SAGE_L)
for col, w in zip("BCDEFGH", (22, 3, 22, 3, 18, 3, 18)):
    st.column_dimensions[col].width = w
st["B40"] = "Tip: keep category names short. If you rename a category, also update it in Transactions and Budget Plan (Find & Replace)."
st["B40"].font = f(color=GREY, italic=True)

# Named ranges (work in Excel + Google Sheets)
def add_name(name, ref):
    dn = DefinedName(name, attr_text=ref)
    wb.defined_names[name] = dn
add_name("ExpenseCats", "Settings!$B$8:$B$37")
add_name("IncomeCats", "Settings!$D$8:$D$37")
add_name("Accounts", "Settings!$F$8:$F$37")
add_name("Months", "Settings!$H$8:$H$19")

# =====================================================================
# TRANSACTIONS
# =====================================================================
tr = wb.create_sheet("Transactions")
tr.sheet_properties.tabColor = MUST
title(tr, "Transactions", "Log every income and expense here. Yellow cells are yours to fill; the rest is automatic.")
N_TX = 1000
HDR_ROW = 6
header_row(tr, HDR_ROW, ["Date", "Description", "Type", "Category", "Amount", "Account", "Notes", "Month (auto)"])
sample = [
    (dt.date(YEAR, 9, 1), "Paycheck", "Income", "Salary", 3200, "Checking", "example row – delete me"),
    (dt.date(YEAR, 9, 1), "Rent", "Expense", "Housing", 1450, "Checking", "example"),
    (dt.date(YEAR, 9, 3), "Grocery run", "Expense", "Groceries", 128.40, "Credit card", "example"),
    (dt.date(YEAR, 9, 4), "Electric bill", "Expense", "Utilities", 96.10, "Checking", "example"),
    (dt.date(YEAR, 9, 6), "Gas", "Expense", "Transport", 52.00, "Credit card", "example"),
    (dt.date(YEAR, 9, 7), "Streaming", "Expense", "Subscriptions", 15.99, "Credit card", "example"),
    (dt.date(YEAR, 9, 9), "Dinner out", "Expense", "Eating out", 64.30, "Credit card", "example"),
    (dt.date(YEAR, 9, 12), "Freelance invoice", "Income", "Side income", 450, "Checking", "example"),
    (dt.date(YEAR, 9, 14), "Grocery run", "Expense", "Groceries", 142.75, "Credit card", "example"),
    (dt.date(YEAR, 9, 15), "Paycheck", "Income", "Salary", 3200, "Checking", "example"),
    (dt.date(YEAR, 9, 16), "Phone plan", "Expense", "Phone & internet", 70.00, "Checking", "example"),
    (dt.date(YEAR, 9, 18), "Transfer to savings", "Expense", "Savings transfers", 400, "Savings", "example"),
    (dt.date(YEAR, 9, 20), "New shoes", "Expense", "Shopping", 89.99, "Credit card", "example"),
    (dt.date(YEAR, 9, 22), "Coffee", "Expense", "Eating out", 6.50, "Cash", "example"),
    (dt.date(YEAR, 9, 25), "Car insurance", "Expense", "Insurance", 118.00, "Checking", "example"),
    (dt.date(YEAR, 9, 27), "Concert tickets", "Expense", "Fun & hobbies", 75.00, "Credit card", "example"),
]
first = HDR_ROW + 1
last = HDR_ROW + N_TX
for i in range(N_TX):
    r = first + i
    for col in range(2, 10):
        c = tr.cell(row=r, column=col); c.border = BORDER
        if col <= 8: input_hint(c)
    tr.cell(row=r, column=2).number_format = DATEF
    tr.cell(row=r, column=6).number_format = CUR
    mc = tr.cell(row=r, column=9, value=f'=IF(B{r}="","",DATE(YEAR(B{r}),MONTH(B{r}),1))')
    mc.number_format = MONTHF; mc.fill = fill(SAGE_L); mc.font = f(color=GREY)
    if i < len(sample):
        d, desc, typ, cat, amt, acc, note = sample[i]
        for col, v in zip(range(2, 9), (d, desc, typ, cat, amt, acc, note)):
            tr.cell(row=r, column=col, value=v)

dv_type = DataValidation(type="list", formula1='"Income,Expense"', allow_blank=True)
dv_cat = DataValidation(type="list", formula1="=Settings!$B$8:$B$37", allow_blank=True)  # expense list; income cats typed or picked
dv_acc = DataValidation(type="list", formula1="=Settings!$F$8:$F$37", allow_blank=True)
for dv, col in ((dv_type, "D"), (dv_cat, "E"), (dv_acc, "G")):
    dv.error = "Pick a value from the list (or edit the list in Settings)."; dv.errorStyle = "warning"
    tr.add_data_validation(dv); dv.add(f"{col}{first}:{col}{last}")
# Income rows: allow income categories too -> second validation on same column isn't possible; we make the category dropdown
# show BOTH lists by pointing to a combined helper list in Settings col J
st["J7"] = "All categories (auto)"; st["J7"].font = Font(name=FONT, bold=True, color=WHITE); st["J7"].fill = fill(INK); st["J7"].border = BORDER
for i in range(30):
    c = st.cell(row=8 + i, column=10, value=f'=IF(B{8+i}="","",B{8+i})'); c.fill = fill(SAGE_L); c.border = BORDER; c.font = f(color=GREY)
for i in range(30):
    c = st.cell(row=38 + i, column=10, value=f'=IF(D{8+i}="","",D{8+i})'); c.fill = fill(SAGE_L); c.border = BORDER; c.font = f(color=GREY)
st.column_dimensions["J"].width = 20
dv_cat.formula1 = "=Settings!$J$8:$J$67"

for col, w in zip("BCDEFGHI", (13, 30, 11, 20, 13, 14, 26, 15)):
    tr.column_dimensions[col].width = w
tr.freeze_panes = f"B{first}"
tr.auto_filter.ref = f"B{HDR_ROW}:I{last}"
# highlight income rows
tr.conditional_formatting.add(f"B{first}:I{last}", FormulaRule(formula=[f'$D{first}="Income"'], fill=fill(GREEN_L)))

TX = f"Transactions!$"
TX_MONTH = f"Transactions!$I${first}:$I${last}"
TX_TYPE = f"Transactions!$D${first}:$D${last}"
TX_CAT = f"Transactions!$E${first}:$E${last}"
TX_AMT = f"Transactions!$F${first}:$F${last}"

# =====================================================================
# BUDGET PLAN (categories x 12 months, planned amounts)
# =====================================================================
bp = wb.create_sheet("Budget Plan")
bp.sheet_properties.tabColor = SAGE
title(bp, "Budget Plan", "Type a default monthly amount per category. Every month copies the default — overwrite any single month to change it.")
header_row(bp, 6, ["Category", "Default / month"] + [f"=Settings!$H${8+m}" for m in range(12)])
for m in range(12):
    bp.cell(row=6, column=4 + m).number_format = "mmm yy"
bp["B7"] = "INCOME"; bp["B7"].font = f(bold=True, color=INK); bp["B7"].fill = fill(SAGE_L)
for col in range(2, 16): bp.cell(row=7, column=col).fill = fill(SAGE_L)
INC_ROWS = list(range(8, 8 + 10))       # 10 income category slots
for i, r in enumerate(INC_ROWS):
    bp.cell(row=r, column=2, value=f'=IF(Settings!D{8+i}="","",Settings!D{8+i})').font = f(color=GREY)
    d = bp.cell(row=r, column=3); d.number_format = CUR0; input_hint(d); d.border = BORDER
    for m in range(12):
        c = bp.cell(row=r, column=4 + m, value=f"=$C{r}"); c.number_format = CUR0; c.border = BORDER
    bp.cell(row=r, column=2).border = BORDER
r_inc_total = INC_ROWS[-1] + 1
bp.cell(row=r_inc_total, column=2, value="Total planned income").font = f(bold=True)
for col in range(3, 16):
    L = get_column_letter(col)
    c = bp.cell(row=r_inc_total, column=col, value=f"=SUM({L}{INC_ROWS[0]}:{L}{INC_ROWS[-1]})"); c.number_format = CUR0; c.font = f(bold=True); c.border = BORDER; c.fill = fill(SAGE_L)

r0 = r_inc_total + 2
bp.cell(row=r0, column=2, value="EXPENSES").font = f(bold=True, color=INK)
for col in range(2, 16): bp.cell(row=r0, column=col).fill = fill(SAGE_L)
EXP_ROWS = list(range(r0 + 1, r0 + 1 + 30))
for i, r in enumerate(EXP_ROWS):
    bp.cell(row=r, column=2, value=f'=IF(Settings!B{8+i}="","",Settings!B{8+i})').font = f(color=GREY)
    d = bp.cell(row=r, column=3); d.number_format = CUR0; input_hint(d); d.border = BORDER
    for m in range(12):
        c = bp.cell(row=r, column=4 + m, value=f"=$C{r}"); c.number_format = CUR0; c.border = BORDER
    bp.cell(row=r, column=2).border = BORDER
r_exp_total = EXP_ROWS[-1] + 1
bp.cell(row=r_exp_total, column=2, value="Total planned expenses").font = f(bold=True)
for col in range(3, 16):
    L = get_column_letter(col)
    c = bp.cell(row=r_exp_total, column=col, value=f"=SUM({L}{EXP_ROWS[0]}:{L}{EXP_ROWS[-1]})"); c.number_format = CUR0; c.font = f(bold=True); c.border = BORDER; c.fill = fill(SAGE_L)
r_left = r_exp_total + 1
bp.cell(row=r_left, column=2, value="Planned left over").font = f(bold=True, color=INK)
for col in range(3, 16):
    L = get_column_letter(col)
    c = bp.cell(row=r_left, column=col, value=f"={L}{r_inc_total}-{L}{r_exp_total}"); c.number_format = CUR0; c.font = f(bold=True, color=INK); c.border = BORDER
bp.conditional_formatting.add(f"C{r_left}:O{r_left}", CellIsRule(operator="lessThan", formula=["0"], fill=fill(RED_L)))
# sample defaults
defaults_inc = {"Salary": 6400, "Side income": 450}
defaults_exp = {"Housing": 1450, "Utilities": 120, "Groceries": 550, "Eating out": 200, "Transport": 220, "Insurance": 130,
                "Health": 60, "Subscriptions": 45, "Phone & internet": 90, "Shopping": 150, "Personal care": 40,
                "Gifts & giving": 50, "Fun & hobbies": 120, "Travel": 150, "Savings transfers": 800, "Other": 75}
for i, name in enumerate(income_cats):
    if name in defaults_inc: bp.cell(row=INC_ROWS[i], column=3, value=defaults_inc[name])
for i, name in enumerate(expense_cats):
    if name in defaults_exp: bp.cell(row=EXP_ROWS[i], column=3, value=defaults_exp[name])
bp.column_dimensions["B"].width = 22; bp.column_dimensions["C"].width = 15
for m in range(12): bp.column_dimensions[get_column_letter(4 + m)].width = 11
bp.freeze_panes = "D7"
bp.cell(row=r_left + 2, column=2, value="Defaults above are examples – replace them with your own numbers.").font = f(color=GREY, italic=True)

BP_MONTHS = "'Budget Plan'!$D$6:$O$6"

# =====================================================================
# DASHBOARD
# =====================================================================
db = wb.create_sheet("Dashboard", 0)
db.sheet_properties.tabColor = INK
title(db, "Tidy Money — Budget Dashboard")
db["B3"] = "Pick a month →"; db["B3"].font = f(bold=True, size=11, color=INK)
db["C3"] = f"=DATE(Settings!$C$5,9,1)"
db["C3"].number_format = MONTHF; db["C3"].font = f(bold=True, size=12, color=INK); input_hint(db["C3"]); db["C3"].border = BORDER
db["C3"].alignment = Alignment(horizontal="center")
dv_m = DataValidation(type="list", formula1="=Settings!$H$8:$H$19", allow_blank=False)
db.add_data_validation(dv_m); dv_m.add("C3")
db["E3"] = "(dropdown – all tabs marked 'selected month' follow this cell)"; db["E3"].font = f(color=GREY, italic=True)
SEL = "Dashboard!$C$3"

# tiles row 5-8
tiles = [
    ("Income this month", f'=SUMIFS({TX_AMT},{TX_TYPE},"Income",{TX_MONTH},{SEL})', CUR0),
    ("Spent this month",  f'=SUMIFS({TX_AMT},{TX_TYPE},"Expense",{TX_MONTH},{SEL})', CUR0),
    ("Left over",         "=C6-F6", CUR0),
    ("Savings rate",      '=IF(C6=0,0,I6/C6)', PCT),
]
col = 3
for label, formula, fmt in tiles:
    L = get_column_letter(col)
    db.merge_cells(f"{L}5:{get_column_letter(col+1)}5"); db.merge_cells(f"{L}6:{get_column_letter(col+1)}6")
    a = db[f"{L}5"]; a.value = label; a.font = f(bold=True, size=9, color=GREY); a.alignment = Alignment(horizontal="center")
    b = db[f"{L}6"]; b.value = formula; b.number_format = fmt; b.font = Font(name=FONT, bold=True, size=20, color=INK); b.alignment = Alignment(horizontal="center", vertical="center")
    for rr in (5, 6):
        for cc in (col, col + 1):
            db.cell(row=rr, column=cc).fill = fill(WHITE); db.cell(row=rr, column=cc).border = BORDER
    db.row_dimensions[6].height = 34
    col += 3
db.conditional_formatting.add("I6", CellIsRule(operator="lessThan", formula=["0"], font=Font(name=FONT, bold=True, size=20, color="B03A2E")))

# planned vs actual summary
db["C8"] = "Planned expenses"; db["C8"].font = f(bold=True, size=9, color=GREY)
db["C9"] = f"=IFERROR(INDEX('Budget Plan'!$D${r_exp_total}:$O${r_exp_total},MATCH({SEL},{BP_MONTHS},0)),0)"; db["C9"].number_format = CUR0; db["C9"].font = f(bold=True, size=14, color=INK)
db["F8"] = "Actual vs plan"; db["F8"].font = f(bold=True, size=9, color=GREY)
db["F9"] = "=F6-C9"; db["F9"].number_format = '"$"#,##0" over";"$"#,##0" under";"on plan"'; db["F9"].font = f(bold=True, size=14, color=INK)
db.conditional_formatting.add("F9", CellIsRule(operator="greaterThan", formula=["0"], font=Font(name=FONT, bold=True, size=14, color="B03A2E")))
db["I8"] = "Planned income"; db["I8"].font = f(bold=True, size=9, color=GREY)
db["I9"] = f"=IFERROR(INDEX('Budget Plan'!$D${r_inc_total}:$O${r_inc_total},MATCH({SEL},{BP_MONTHS},0)),0)"; db["I9"].number_format = CUR0; db["I9"].font = f(bold=True, size=14, color=INK)
db["L8"] = "Transactions logged"; db["L8"].font = f(bold=True, size=9, color=GREY)
db["L9"] = f"=COUNTIFS({TX_MONTH},{SEL})"; db["L9"].font = f(bold=True, size=14, color=INK)

# category table (selected month)
header_row(db, 11, ["Category", "Planned", "Actual", "Left", "% used"], start_col=3)
CAT_FIRST = 12
for i in range(30):
    r = CAT_FIRST + i
    db.cell(row=r, column=3, value=f'=IF(Settings!B{8+i}="","",Settings!B{8+i})')
    db.cell(row=r, column=4, value=f'=IF($C{r}="","",IFERROR(INDEX(\'Budget Plan\'!$D${EXP_ROWS[i]}:$O${EXP_ROWS[i]},MATCH({SEL},{BP_MONTHS},0)),0))').number_format = CUR0
    db.cell(row=r, column=5, value=f'=IF($C{r}="","",SUMIFS({TX_AMT},{TX_TYPE},"Expense",{TX_CAT},$C{r},{TX_MONTH},{SEL}))').number_format = CUR0
    db.cell(row=r, column=6, value=f'=IF($C{r}="","",D{r}-E{r})').number_format = CUR0
    db.cell(row=r, column=7, value=f'=IF($C{r}="","",IF(D{r}=0,IF(E{r}>0,1,0),E{r}/D{r}))').number_format = PCT
    for cc in range(3, 8):
        db.cell(row=r, column=cc).border = BORDER; db.cell(row=r, column=cc).fill = fill(WHITE)
CAT_LAST = CAT_FIRST + 29
rt = CAT_LAST + 1
db.cell(row=rt, column=3, value="Total").font = f(bold=True)
for cc, L in zip(range(4, 7), "DEF"):
    c = db.cell(row=rt, column=cc, value=f"=SUM({L}{CAT_FIRST}:{L}{CAT_LAST})"); c.number_format = CUR0; c.font = f(bold=True); c.border = BORDER; c.fill = fill(SAGE_L)
c = db.cell(row=rt, column=7, value=f"=IF(D{rt}=0,0,E{rt}/D{rt})"); c.number_format = PCT; c.font = f(bold=True); c.border = BORDER; c.fill = fill(SAGE_L)
db.cell(row=rt, column=3).border = BORDER; db.cell(row=rt, column=3).fill = fill(SAGE_L)
db.conditional_formatting.add(f"G{CAT_FIRST}:G{CAT_LAST}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color=SAGE, showValue=True))
db.conditional_formatting.add(f"F{CAT_FIRST}:F{CAT_LAST}", CellIsRule(operator="lessThan", formula=["0"], fill=fill(RED_L)))

# top 5 categories
header_row(db, 11, ["Top spending", "Amount"], start_col=9)
for k in range(1, 6):
    r = 11 + k
    db.cell(row=r, column=9, value=f'=IFERROR(INDEX($C${CAT_FIRST}:$C${CAT_LAST},MATCH(LARGE($E${CAT_FIRST}:$E${CAT_LAST},{k}),$E${CAT_FIRST}:$E${CAT_LAST},0)),"")')
    c = db.cell(row=r, column=10, value=f'=IFERROR(LARGE($E${CAT_FIRST}:$E${CAT_LAST},{k}),"")'); c.number_format = CUR0
    for cc in (9, 10): db.cell(row=r, column=cc).border = BORDER; db.cell(row=r, column=cc).fill = fill(WHITE)
db["I18"] = "Note: if two categories tie, the first one is shown twice – that's normal for LARGE/MATCH."; db["I18"].font = f(size=8, color=GREY, italic=True)

# chart planned vs actual (top 12 categories rows)
ch = BarChart(); ch.type = "bar"; ch.style = 10; ch.title = "Planned vs actual — selected month"
ch.y_axis.title = None; ch.x_axis.title = None
data = Reference(db, min_col=4, max_col=5, min_row=11, max_row=CAT_FIRST + 15)
cats = Reference(db, min_col=3, min_row=CAT_FIRST, max_row=CAT_FIRST + 15)
ch.add_data(data, titles_from_data=True); ch.set_categories(cats)
ch.height = 11; ch.width = 16
ch.series[0].graphicalProperties.solidFill = LINE
ch.series[1].graphicalProperties.solidFill = SAGE
ch.legend.position = "b"
db.add_chart(ch, "I20")

for colL, w in zip("BCDEFGHIJKLM", (2, 22, 12, 12, 12, 12, 3, 20, 12, 3, 16, 12)):
    db.column_dimensions[colL].width = w
db.freeze_panes = "B4"

# =====================================================================
# MONTHLY BUDGET (detailed, selected month: income + expenses)
# =====================================================================
mb = wb.create_sheet("Monthly Budget", 1)
mb.sheet_properties.tabColor = SAGE
title(mb, "Monthly Budget", "Selected month (change it on the Dashboard):")
mb["E3"] = f"={SEL}"; mb["E3"].number_format = MONTHF; mb["E3"].font = f(bold=True, size=12, color=INK)
header_row(mb, 5, ["Income", "Planned", "Actual", "Difference"])
for i in range(10):
    r = 6 + i
    mb.cell(row=r, column=2, value=f'=IF(Settings!D{8+i}="","",Settings!D{8+i})')
    mb.cell(row=r, column=3, value=f'=IF($B{r}="","",IFERROR(INDEX(\'Budget Plan\'!$D${INC_ROWS[i]}:$O${INC_ROWS[i]},MATCH({SEL},{BP_MONTHS},0)),0))').number_format = CUR0
    mb.cell(row=r, column=4, value=f'=IF($B{r}="","",SUMIFS({TX_AMT},{TX_TYPE},"Income",{TX_CAT},$B{r},{TX_MONTH},{SEL}))').number_format = CUR0
    mb.cell(row=r, column=5, value=f'=IF($B{r}="","",D{r}-C{r})').number_format = CUR0
    for cc in range(2, 6): mb.cell(row=r, column=cc).border = BORDER; mb.cell(row=r, column=cc).fill = fill(WHITE)
mb.cell(row=16, column=2, value="Total income").font = f(bold=True)
for cc, L in zip(range(3, 6), "CDE"):
    c = mb.cell(row=16, column=cc, value=f"=SUM({L}6:{L}15)"); c.number_format = CUR0; c.font = f(bold=True); c.fill = fill(SAGE_L); c.border = BORDER
mb.cell(row=16, column=2).fill = fill(SAGE_L); mb.cell(row=16, column=2).border = BORDER

header_row(mb, 18, ["Expenses", "Planned", "Actual", "Left to spend", "% used"])
for i in range(30):
    r = 19 + i
    mb.cell(row=r, column=2, value=f"=Dashboard!C{CAT_FIRST+i}")
    mb.cell(row=r, column=3, value=f"=Dashboard!D{CAT_FIRST+i}").number_format = CUR0
    mb.cell(row=r, column=4, value=f"=Dashboard!E{CAT_FIRST+i}").number_format = CUR0
    mb.cell(row=r, column=5, value=f"=Dashboard!F{CAT_FIRST+i}").number_format = CUR0
    mb.cell(row=r, column=6, value=f"=Dashboard!G{CAT_FIRST+i}").number_format = PCT
    for cc in range(2, 7): mb.cell(row=r, column=cc).border = BORDER; mb.cell(row=r, column=cc).fill = fill(WHITE)
mb.cell(row=49, column=2, value="Total expenses").font = f(bold=True)
for cc, L in zip(range(3, 6), "CDE"):
    c = mb.cell(row=49, column=cc, value=f"=SUM({L}19:{L}48)"); c.number_format = CUR0; c.font = f(bold=True); c.fill = fill(SAGE_L); c.border = BORDER
c = mb.cell(row=49, column=6, value="=IF(C49=0,0,D49/C49)"); c.number_format = PCT; c.font = f(bold=True); c.fill = fill(SAGE_L); c.border = BORDER
mb.cell(row=49, column=2).fill = fill(SAGE_L); mb.cell(row=49, column=2).border = BORDER
mb["B51"] = "Left over (actual)"; mb["B51"].font = f(bold=True, size=12, color=INK)
mb["D51"] = "=D16-D49"; mb["D51"].number_format = CUR0; mb["D51"].font = Font(name=FONT, bold=True, size=14, color=INK)
mb["B52"] = "Left over (planned)"; mb["B52"].font = f(color=GREY)
mb["D52"] = "=C16-C49"; mb["D52"].number_format = CUR0; mb["D52"].font = f(color=GREY)
mb.conditional_formatting.add("F19:F48", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color=SAGE, showValue=True))
mb.conditional_formatting.add("E19:E48", CellIsRule(operator="lessThan", formula=["0"], fill=fill(RED_L)))
for colL, w in zip("BCDEF", (22, 13, 13, 14, 12)): mb.column_dimensions[colL].width = w
mb.freeze_panes = "B6"

# =====================================================================
# ANNUAL OVERVIEW
# =====================================================================
an = wb.create_sheet("Annual Overview")
an.sheet_properties.tabColor = SAGE
title(an, "Annual Overview", "Actual spending per category, every month of the year. Fully automatic.")
header_row(an, 6, ["Category"] + [f"=Settings!$H${8+m}" for m in range(12)] + ["Year total", "Monthly avg"])
for m in range(12): an.cell(row=6, column=3 + m).number_format = "mmm"
an["B7"] = "INCOME"; an["B7"].font = f(bold=True, color=INK)
for cc in range(2, 17): an.cell(row=7, column=cc).fill = fill(SAGE_L)
for i in range(10):
    r = 8 + i
    an.cell(row=r, column=2, value=f'=IF(Settings!D{8+i}="","",Settings!D{8+i})')
    for m in range(12):
        L = get_column_letter(3 + m)
        c = an.cell(row=r, column=3 + m, value=f'=IF($B{r}="","",SUMIFS({TX_AMT},{TX_TYPE},"Income",{TX_CAT},$B{r},{TX_MONTH},{L}$6))'); c.number_format = CUR0
    an.cell(row=r, column=15, value=f'=IF($B{r}="","",SUM(C{r}:N{r}))').number_format = CUR0
    an.cell(row=r, column=16, value=f'=IF($B{r}="","",IFERROR(O{r}/MAX(1,COUNTIF(C{r}:N{r},">0")),0))').number_format = CUR0
    for cc in range(2, 17): an.cell(row=r, column=cc).border = BORDER
an.cell(row=18, column=2, value="Total income").font = f(bold=True)
for cc in range(3, 17):
    L = get_column_letter(cc)
    c = an.cell(row=18, column=cc, value=f"=SUM({L}8:{L}17)"); c.number_format = CUR0; c.font = f(bold=True); c.fill = fill(SAGE_L); c.border = BORDER
an["B20"] = "EXPENSES"; an["B20"].font = f(bold=True, color=INK)
for cc in range(2, 17): an.cell(row=20, column=cc).fill = fill(SAGE_L)
for i in range(30):
    r = 21 + i
    an.cell(row=r, column=2, value=f'=IF(Settings!B{8+i}="","",Settings!B{8+i})')
    for m in range(12):
        L = get_column_letter(3 + m)
        c = an.cell(row=r, column=3 + m, value=f'=IF($B{r}="","",SUMIFS({TX_AMT},{TX_TYPE},"Expense",{TX_CAT},$B{r},{TX_MONTH},{L}$6))'); c.number_format = CUR0
    an.cell(row=r, column=15, value=f'=IF($B{r}="","",SUM(C{r}:N{r}))').number_format = CUR0
    an.cell(row=r, column=16, value=f'=IF($B{r}="","",IFERROR(O{r}/MAX(1,COUNTIF(C{r}:N{r},">0")),0))').number_format = CUR0
    for cc in range(2, 17): an.cell(row=r, column=cc).border = BORDER
an.cell(row=51, column=2, value="Total expenses").font = f(bold=True)
for cc in range(3, 17):
    L = get_column_letter(cc)
    c = an.cell(row=51, column=cc, value=f"=SUM({L}21:{L}50)"); c.number_format = CUR0; c.font = f(bold=True); c.fill = fill(SAGE_L); c.border = BORDER
an.cell(row=52, column=2, value="Left over").font = f(bold=True, color=INK)
for cc in range(3, 16):
    L = get_column_letter(cc)
    c = an.cell(row=52, column=cc, value=f"={L}18-{L}51"); c.number_format = CUR0; c.font = f(bold=True, color=INK); c.border = BORDER
an.cell(row=53, column=2, value="Savings rate").font = f(color=GREY)
for cc in range(3, 16):
    L = get_column_letter(cc)
    c = an.cell(row=53, column=cc, value=f"=IF({L}18=0,0,{L}52/{L}18)"); c.number_format = PCT; c.font = f(color=GREY); c.border = BORDER
an.conditional_formatting.add("C52:O52", CellIsRule(operator="lessThan", formula=["0"], fill=fill(RED_L)))
an.conditional_formatting.add("C21:N50", FormulaRule(formula=["C21>0"], fill=fill("F3F7F4")))
lc = LineChart(); lc.title = "Income vs expenses by month"; lc.style = 12; lc.height = 8; lc.width = 22
lc.add_data(Reference(an, min_col=2, max_col=14, min_row=18, max_row=18), from_rows=True, titles_from_data=True)
lc.add_data(Reference(an, min_col=2, max_col=14, min_row=51, max_row=51), from_rows=True, titles_from_data=True)
lc.set_categories(Reference(an, min_col=3, max_col=14, min_row=6, max_row=6))
lc.series[0].graphicalProperties.line.solidFill = SAGE; lc.series[0].graphicalProperties.line.width = 28000
lc.series[1].graphicalProperties.line.solidFill = MUST; lc.series[1].graphicalProperties.line.width = 28000
lc.legend.position = "b"
an.add_chart(lc, "B56")
an.column_dimensions["B"].width = 22
for m in range(14): an.column_dimensions[get_column_letter(3 + m)].width = 10
an.freeze_panes = "C7"

# =====================================================================
# SAVINGS GOALS
# =====================================================================
sg = wb.create_sheet("Savings Goals")
sg.sheet_properties.tabColor = MUST
title(sg, "Savings Goals & Sinking Funds", "One row per goal. Update 'Saved so far' whenever you move money. Progress and monthly need are automatic.")
sg["B5"] = "Today's date"; sg["B5"].font = f(bold=True); sg["C5"] = "=TODAY()"; sg["C5"].number_format = DATEF; sg["C5"].font = f(color=GREY)
header_row(sg, 7, ["Goal", "Target amount", "Saved so far", "Remaining", "Progress", "Target date", "Months left", "Save per month", "Status"])
goals = [("Emergency fund", 6000, 2150, dt.date(YEAR + 1, 6, 30)), ("Vacation", 1800, 400, dt.date(YEAR + 1, 3, 1)),
         ("New laptop", 1400, 1400, dt.date(YEAR, 12, 1)), ("Car repairs fund", 1000, 250, dt.date(YEAR + 1, 1, 31)),
         ("Holiday gifts", 600, 120, dt.date(YEAR, 12, 10))]
for i in range(20):
    r = 8 + i
    for cc in range(2, 11):
        c = sg.cell(row=r, column=cc); c.border = BORDER
        if cc in (2, 3, 4, 7): input_hint(c)
    sg.cell(row=r, column=3).number_format = CUR0; sg.cell(row=r, column=4).number_format = CUR0
    sg.cell(row=r, column=5, value=f'=IF(B{r}="","",MAX(0,C{r}-D{r}))').number_format = CUR0
    sg.cell(row=r, column=6, value=f'=IF(OR(B{r}="",C{r}=0),"",MIN(1,D{r}/C{r}))').number_format = PCT
    sg.cell(row=r, column=7).number_format = DATEF
    sg.cell(row=r, column=8, value=f'=IF(OR(B{r}="",G{r}=""),"",MAX(0,(YEAR(G{r})-YEAR($C$5))*12+MONTH(G{r})-MONTH($C$5)))')
    sg.cell(row=r, column=9, value=f'=IF(OR(B{r}="",G{r}=""),"",IF(E{r}=0,0,IF(H{r}=0,E{r},E{r}/H{r})))').number_format = CUR0
    sg.cell(row=r, column=10, value=f'=IF(B{r}="","",IF(E{r}=0,"Done ✓",IF(G{r}="","No date",IF(H{r}=0,"Due now","On track"))))')
    if i < len(goals):
        g, t, s, d = goals[i]
        sg.cell(row=r, column=2, value=g); sg.cell(row=r, column=3, value=t); sg.cell(row=r, column=4, value=s); sg.cell(row=r, column=7, value=d)
sg.conditional_formatting.add("F8:F27", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color=MUST, showValue=True))
sg.conditional_formatting.add("J8:J27", CellIsRule(operator="equal", formula=['"Done ✓"'], fill=fill(GREEN_L)))
sg.cell(row=29, column=2, value="Totals").font = f(bold=True)
for cc, L in zip((3, 4, 5, 9), "CDEI"):
    c = sg.cell(row=29, column=cc, value=f"=SUM({L}8:{L}27)"); c.number_format = CUR0; c.font = f(bold=True); c.fill = fill(SAGE_L); c.border = BORDER
sg.cell(row=31, column=2, value="Example goals shown – replace with yours. 'Save per month' = remaining ÷ months left until the target date.").font = f(color=GREY, italic=True)
for colL, w in zip("BCDEFGHIJ", (24, 14, 14, 13, 14, 14, 11, 14, 12)): sg.column_dimensions[colL].width = w

# =====================================================================
# DEBT PAYOFF
# =====================================================================
dp = wb.create_sheet("Debt Payoff")
dp.sheet_properties.tabColor = "B03A2E"
title(dp, "Debt Payoff Planner", "List each debt. Choose Snowball (smallest balance first) or Avalanche (highest rate first). Estimates assume fixed payments and no new charges.")
dp["B5"] = "Method"; dp["B5"].font = f(bold=True)
dp["C5"] = "Avalanche"; input_hint(dp["C5"]); dp["C5"].border = BORDER; dp["C5"].font = f(bold=True, color=INK)
dvm = DataValidation(type="list", formula1='"Snowball,Avalanche"', allow_blank=False); dp.add_data_validation(dvm); dvm.add("C5")
dp["E5"] = "Extra you can pay per month"; dp["E5"].font = f(bold=True)
dp["G5"] = 150; dp["G5"].number_format = CUR0; input_hint(dp["G5"]); dp["G5"].border = BORDER
header_row(dp, 7, ["Debt", "Balance", "APR %", "Minimum payment", "Payoff order", "Months at minimum", "Interest at minimum (est.)", "Monthly interest now", "Note"])
debts = [("Credit card A", 3200, 0.2499, 95), ("Car loan", 8900, 0.069, 265), ("Student loan", 12400, 0.055, 140), ("Store card", 640, 0.2899, 30)]
for i in range(15):
    r = 8 + i
    for cc in range(2, 11):
        c = dp.cell(row=r, column=cc); c.border = BORDER
        if cc in (2, 3, 4, 5, 10): input_hint(c)
    dp.cell(row=r, column=3).number_format = CUR0; dp.cell(row=r, column=4).number_format = '0.00%'; dp.cell(row=r, column=5).number_format = CUR0
    # order: snowball = rank by balance asc; avalanche = rank by APR desc (ties broken by row)
    dp.cell(row=r, column=6, value=(f'=IF(B{r}="","",IF($C$5="Snowball",'
                                     f'COUNTIFS($C$8:$C$22,"<"&C{r},$B$8:$B$22,"<>")+COUNTIFS($C$8:$C$22,C{r},$B$8:$B$22,"<>")-COUNTIFS($C$8:$C{r},C{r},$B$8:$B{r},"<>")+1,'
                                     f'COUNTIFS($D$8:$D$22,">"&D{r},$B$8:$B$22,"<>")+COUNTIFS($D$8:$D$22,D{r},$B$8:$B$22,"<>")-COUNTIFS($D$8:$D{r},D{r},$B$8:$B{r},"<>")+1))'))
    dp.cell(row=r, column=7, value=f'=IF(B{r}="","",IF(E{r}<=C{r}*D{r}/12,"Never – raise payment",IFERROR(ROUNDUP(NPER(D{r}/12,-E{r},C{r}),0),ROUNDUP(C{r}/MAX(E{r},1),0))))')
    dp.cell(row=r, column=8, value=f'=IF(OR(B{r}="",NOT(ISNUMBER(G{r}))),"",G{r}*E{r}-C{r})').number_format = CUR0
    dp.cell(row=r, column=9, value=f'=IF(B{r}="","",C{r}*D{r}/12)').number_format = CUR
    if i < len(debts):
        n, b, a, m = debts[i]
        dp.cell(row=r, column=2, value=n); dp.cell(row=r, column=3, value=b); dp.cell(row=r, column=4, value=a); dp.cell(row=r, column=5, value=m)
dp.conditional_formatting.add("F8:F22", CellIsRule(operator="equal", formula=["1"], fill=fill(MUST)))
dp.cell(row=24, column=2, value="Totals").font = f(bold=True)
for cc, L in zip((3, 5, 9), "CEI"):
    c = dp.cell(row=24, column=cc, value=f"=SUM({L}8:{L}22)"); c.number_format = CUR0 if L != "I" else CUR; c.font = f(bold=True); c.fill = fill(SAGE_L); c.border = BORDER
dp["B26"] = "How to use the plan"; dp["B26"].font = f(bold=True, size=11, color=INK)
dp["B27"] = "1. Pay the minimum on every debt."; dp["B28"] = "2. Send the 'extra' amount (cell G5) to the debt with Payoff order = 1 (highlighted)."
dp["B29"] = "3. When it's paid off, add its old minimum to the extra and move to order 2. Repeat."
dp["B30"] = "Snowball = quick wins (smallest balance first). Avalanche = least interest overall (highest APR first). Both work — pick the one you'll stick with."
dp["B31"] = "Estimates only. They assume fixed payments, no new charges, and simple monthly compounding. Your lender's statement is the source of truth."
for r in range(27, 32): dp.cell(row=r, column=2).font = f(color=GREY)
for colL, w in zip("BCDEFGHIJ", (20, 13, 10, 16, 12, 18, 22, 18, 24)): dp.column_dimensions[colL].width = w

# =====================================================================
# NET WORTH
# =====================================================================
nw = wb.create_sheet("Net Worth")
nw.sheet_properties.tabColor = INK
title(nw, "Net Worth Tracker", "Once a month, type your balances. Net worth = what you own − what you owe.")
header_row(nw, 6, ["Month", "Cash & checking", "Savings", "Investments / retirement", "Home / car value", "Other assets", "TOTAL ASSETS",
                   "Credit cards", "Loans", "Mortgage", "Other debts", "TOTAL DEBTS", "NET WORTH", "Change"])
for i in range(24):
    r = 7 + i
    mcell = nw.cell(row=r, column=2, value=f"=DATE(Settings!$C$5,{i+1},1)" if i == 0 else f"=DATE(YEAR(B{r-1}),MONTH(B{r-1})+1,1)")
    mcell.number_format = "mmm yyyy"; mcell.font = f(color=GREY)
    for cc in range(2, 16):
        c = nw.cell(row=r, column=cc); c.border = BORDER
        if cc in (3, 4, 5, 6, 7, 9, 10, 11, 12): input_hint(c); c.number_format = CUR0
    nw.cell(row=r, column=8, value=f"=SUM(C{r}:G{r})").number_format = CUR0
    nw.cell(row=r, column=13, value=f"=SUM(I{r}:L{r})").number_format = CUR0
    nw.cell(row=r, column=14, value=f'=IF(AND(H{r}=0,M{r}=0),"",H{r}-M{r})').number_format = CUR0
    nw.cell(row=r, column=15, value=("" if i == 0 else f'=IF(OR(N{r}="",N{r-1}=""),"",N{r}-N{r-1})')).number_format = CUR0
    for cc in (8, 13, 14): nw.cell(row=r, column=cc).font = f(bold=True); nw.cell(row=r, column=cc).fill = fill(SAGE_L)
    nw.cell(row=r, column=14).font = f(bold=True, color=INK)
# sample 3 months
samples_nw = [(2400, 6100, 15200, 0, 0, 3200, 21300, 0, 0), (2650, 6500, 15650, 0, 0, 3050, 21035, 0, 0), (2100, 6900, 15900, 0, 0, 2900, 20770, 0, 0)]
for i, row in enumerate(samples_nw):
    for cc, v in zip((3, 4, 5, 6, 7, 9, 10, 11, 12), row):
        nw.cell(row=7 + i, column=cc, value=v)
nw.conditional_formatting.add("O7:O30", CellIsRule(operator="lessThan", formula=["0"], font=Font(name=FONT, color="B03A2E")))
nw.conditional_formatting.add("O7:O30", CellIsRule(operator="greaterThan", formula=["0"], font=Font(name=FONT, color="2E7D4F")))
nwc = LineChart(); nwc.title = "Net worth over time"; nwc.style = 12; nwc.height = 8; nwc.width = 20
nwc.add_data(Reference(nw, min_col=14, min_row=6, max_row=30), titles_from_data=True)
nwc.set_categories(Reference(nw, min_col=2, min_row=7, max_row=30))
nwc.series[0].graphicalProperties.line.solidFill = INK; nwc.series[0].graphicalProperties.line.width = 28000
nwc.legend = None
nw.add_chart(nwc, "B33")
nw.cell(row=31, column=2, value="First 3 months are examples – overwrite them.").font = f(color=GREY, italic=True)
nw.column_dimensions["B"].width = 12
for cc in range(3, 16): nw.column_dimensions[get_column_letter(cc)].width = 14
nw.freeze_panes = "C7"

# =====================================================================
# START HERE
# =====================================================================
sh = wb.create_sheet("Start Here", 0)
sh.sheet_properties.tabColor = MUST
sh.sheet_view.showGridLines = False
sh.column_dimensions["A"].width = 2; sh.column_dimensions["B"].width = 4; sh.column_dimensions["C"].width = 100
sh["B2"] = "Tidy Money — Budget Dashboard"; sh["B2"].font = Font(name=FONT, bold=True, size=22, color=INK)
sh["B3"] = "Thank you for your purchase. This page gets you set up in about 10 minutes."; sh["B3"].font = f(size=11, color=GREY, italic=True)
steps = [
    ("USING GOOGLE SHEETS?", None),
    ("1", "Upload this file to Google Drive, open it, then choose File → Save as Google Sheets. (Or in Sheets: File → Import → Upload → 'Replace spreadsheet'.) Every formula, dropdown and chart is compatible."),
    ("USING EXCEL?", None),
    ("1", "Just open the file. If Excel shows 'Protected View', click Enable Editing. Works on Mac & Windows, Excel 2016 or newer and Microsoft 365. Apple Numbers opens it too (charts may look slightly different)."),
    ("SET-UP (ONCE)", None),
    ("2", "Settings tab → set your budget year and edit the category and account lists to match your life. Keep names short."),
    ("3", "Budget Plan tab → type a default monthly amount for each income and expense category. Any single month can be overwritten later."),
    ("4", "Transactions tab → delete the example rows (they say 'example' in Notes) and start logging. Date, description, Income/Expense, category, amount. That's it."),
    ("EVERY WEEK (5 MINUTES)", None),
    ("5", "Add your transactions. Open the Dashboard, pick the month in the yellow dropdown, and see what's left in each category."),
    ("6", "Savings Goals → update 'Saved so far'. Debt Payoff → update balances. Net Worth → once a month, type your balances."),
    ("GOOD TO KNOW", None),
    ("•", "Yellow cells are for you to type in. White and green cells are formulas — leave them alone (if you break one, just undo)."),
    ("•", "Savings rate = (income − expenses) ÷ income for the selected month. 'Savings transfers' counts as an expense so money you move to savings isn't double-counted."),
    ("•", "Couples: share the Google Sheet with your partner; you can both log from your phones with the Google Sheets app."),
    ("•", "New year? Save a copy, change the year in Settings, clear the Transactions rows, and keep your plan."),
    ("•", "Need help or a refund within 7 days? Reply to your Whop receipt email."),
]
r = 5
for a, b in steps:
    if b is None:
        sh.cell(row=r, column=2, value=a).font = Font(name=FONT, bold=True, size=11, color=SAGE)
        r += 1; continue
    sh.cell(row=r, column=2, value=a).font = f(bold=True, color=INK)
    c = sh.cell(row=r, column=3, value=b); c.font = f(size=10); c.alignment = Alignment(wrap_text=True, vertical="top")
    sh.row_dimensions[r].height = 32 if len(b) > 95 else 18
    r += 1
sh.cell(row=r + 1, column=3, value="© Tidy Money. For personal use. Please don't resell or redistribute this file.").font = f(size=8, color=GREY, italic=True)

# print setup
for ws in wb.worksheets:
    ws.page_setup.orientation = "landscape"; ws.page_setup.fitToWidth = 1; ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToHeight = 0

wb.active = 0
wb.save(OUT)
print("saved", OUT)
