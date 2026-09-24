from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, ListFlowable, ListItem, PageBreak
from reportlab.lib.enums import TA_LEFT

INK = colors.HexColor("#1F3D2B"); SAGE = colors.HexColor("#7FA98B"); MUST = colors.HexColor("#E3B23C"); GREY = colors.HexColor("#6B6B6B"); CREAM = colors.HexColor("#FAF7F2")
H1 = ParagraphStyle("h1", fontName="Helvetica-Bold", fontSize=24, leading=30, textColor=INK, spaceAfter=6)
H2 = ParagraphStyle("h2", fontName="Helvetica-Bold", fontSize=14, leading=18, textColor=INK, spaceBefore=14, spaceAfter=6)
P = ParagraphStyle("p", fontName="Helvetica", fontSize=10.5, leading=15, textColor=colors.HexColor("#1B1B1B"))
SM = ParagraphStyle("sm", parent=P, fontSize=9, leading=12, textColor=GREY)
TIP = ParagraphStyle("tip", parent=P, backColor=CREAM, borderPadding=8, leftIndent=0)

def bullets(items):
    return ListFlowable([ListItem(Paragraph(i, P), leftIndent=12) for i in items], bulletType="bullet", start="•", leftIndent=14, bulletFontSize=9)

def numbered(items):
    return ListFlowable([ListItem(Paragraph(i, P), leftIndent=14) for i in items], bulletType="1", leftIndent=16)

def footer(canvas, doc):
    canvas.saveState(); canvas.setFont("Helvetica", 8); canvas.setFillColor(GREY)
    canvas.drawString(inch, 0.6 * inch, "Tidy Money — Budget Dashboard · Quick Start Guide")
    canvas.drawRightString(letter[0] - inch, 0.6 * inch, f"Page {doc.page}")
    canvas.setStrokeColor(SAGE); canvas.setLineWidth(2); canvas.line(inch, letter[1] - 0.55 * inch, letter[0] - inch, letter[1] - 0.55 * inch)
    canvas.restoreState()

doc = SimpleDocTemplate("product/Tidy-Money-Quick-Start-Guide.pdf", pagesize=letter, leftMargin=inch, rightMargin=inch, topMargin=0.9 * inch, bottomMargin=0.9 * inch,
                        title="Tidy Money — Quick Start Guide", author="Tidy Money")
s = []
s += [Paragraph("Tidy Money", ParagraphStyle("k", parent=SM, textColor=SAGE, fontName="Helvetica-Bold", fontSize=11)),
      Paragraph("Budget Dashboard — Quick Start Guide", H1),
      Paragraph("Thanks for your purchase. Ten minutes from now you'll have a working budget that updates itself every time you log a transaction. This guide covers set-up, the weekly routine, and answers to the questions people ask most.", P), Spacer(1, 10)]

s += [Paragraph("1. Open the file", H2)]
t = Table([
    [Paragraph("<b>Google Sheets</b> (recommended for phones and sharing)", P), Paragraph("<b>Microsoft Excel</b>", P)],
    [numbered(["Go to drive.google.com and upload <b>Tidy-Money-Budget-Dashboard.xlsx</b> (New → File upload).",
               "Double-click it to open, then choose <b>File → Save as Google Sheets</b>.",
               "Use the new Google Sheets version from now on. The .xlsx is your backup."]),
     numbered(["Double-click the .xlsx file.",
               "If you see <b>Protected View</b>, click <b>Enable Editing</b>.",
               "Save it somewhere you'll find it (Documents or OneDrive)."])],
], colWidths=[3.25 * inch, 3.25 * inch])
t.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("BACKGROUND", (0, 0), (-1, 0), CREAM), ("BOX", (0, 0), (-1, -1), 0.5, SAGE), ("INNERGRID", (0, 0), (-1, -1), 0.5, SAGE), ("PADDING", (0, 0), (-1, -1), 8)]))
s += [t, Spacer(1, 6), Paragraph("Requirements: Google Sheets (any account) or Excel 2016+/Microsoft 365 on Mac or Windows. Apple Numbers also opens the file; charts may look slightly different there.", SM)]

s += [Paragraph("2. Set it up once (about 10 minutes)", H2), numbered([
    "<b>Settings tab</b> — set your budget year. Edit the expense categories, income categories and accounts so they match your life. Keep names short. Every dropdown in the workbook reads these lists.",
    "<b>Budget Plan tab</b> — type a default monthly amount for each income and expense category (the yellow 'Default / month' column). Each month copies the default automatically. If one month is different (say December gifts), just overwrite that month's cell.",
    "<b>Transactions tab</b> — delete the example rows (they say 'example' in the Notes column). Start logging: date, description, Income or Expense, category, amount, account.",
    "<b>Dashboard tab</b> — pick a month in the yellow dropdown. Income, spending, what's left, savings rate, the category table and the chart all update by themselves.",
])]

s += [Paragraph("3. Your weekly routine (about 5 minutes)", H2), bullets([
    "Add the week's transactions. Bank app open on one side, sheet on the other. Round numbers are fine.",
    "Open the Dashboard. Check 'Left to spend' per category — that's your guardrail for the rest of the month.",
    "Moving money to savings? Log it as an expense under 'Savings transfers' and update <b>Savings Goals</b> → 'Saved so far'.",
    "Once a month: <b>Net Worth</b> tab, type your account balances. <b>Debt Payoff</b> tab, update balances.",
])]

s += [Paragraph("4. What each tab does", H2)]
tabs = [
    ("Start Here", "The short version of this guide, inside the workbook."),
    ("Dashboard", "Month selector, four headline numbers, planned vs actual per category, top 5 spending, chart."),
    ("Monthly Budget", "The detailed view of the selected month: income and expenses, planned vs actual, left to spend, % used."),
    ("Transactions", "Your log. 1,000 rows ready with dropdowns and an automatic Month column. Add more rows by copying the last one."),
    ("Budget Plan", "Planned amounts for every category × 12 months. Set defaults once; override any month."),
    ("Annual Overview", "Actual spending per category for every month of the year, totals, monthly averages and an income-vs-expenses trend chart."),
    ("Savings Goals", "Sinking funds and goals: target, saved, remaining, progress bar, and how much to put aside per month to hit the date."),
    ("Debt Payoff", "List your debts, pick Snowball or Avalanche, and see the payoff order, months to payoff at minimum and estimated interest."),
    ("Net Worth", "24 months of assets and debts with automatic totals, net worth, month-over-month change and a chart."),
    ("Settings", "Categories, accounts and the budget year. Change here, everything follows."),
]
tt = Table([[Paragraph(f"<b>{a}</b>", P), Paragraph(b, P)] for a, b in tabs], colWidths=[1.5 * inch, 5.0 * inch])
tt.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("ROWBACKGROUNDS", (0, 0), (-1, -1), [colors.white, CREAM]), ("LINEBELOW", (0, 0), (-1, -1), 0.25, SAGE), ("PADDING", (0, 0), (-1, -1), 6)]))
s += [tt]

s += [Paragraph("5. Good to know", H2), bullets([
    "<b>Yellow cells are yours.</b> White and green cells hold formulas — leave them alone. If you break one, press Undo.",
    "<b>Savings rate</b> = (income − expenses) ÷ income for the selected month. Because 'Savings transfers' is an expense category, money you move to savings is counted once, not twice.",
    "<b>Couples and families:</b> in Google Sheets, click Share and add your partner. Both of you can log from the free Google Sheets phone app.",
    "<b>Renaming a category</b> after you've logged transactions? Use Find &amp; Replace (Ctrl/Cmd+H) across the whole workbook so old rows follow.",
    "<b>New year:</b> save a copy, change the year in Settings, clear the Transactions rows (keep the header), and your plan and categories carry over.",
    "<b>Debt estimates</b> assume fixed payments, no new charges and simple monthly compounding. They're a planning aid — your lender's statement is the source of truth.",
])]

s += [Paragraph("6. Help &amp; refunds", H2),
      Paragraph("Stuck, or found something that doesn't work the way this guide says? Reply to your Whop receipt email and describe what you see. Not for you? Reply within 7 days of purchase and you'll get a refund — no questions asked.", P),
      Spacer(1, 16), Paragraph("© Tidy Money. Licensed for personal use by the purchaser. Please don't resell, share or redistribute the files.", SM)]

doc.build(s, onFirstPage=footer, onLaterPages=footer)
print("ok")
