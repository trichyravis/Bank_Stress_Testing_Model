
"""
╔══════════════════════════════════════════════════════════════════╗
║   THE MOUNTAIN PATH - WORLD OF FINANCE                          ║
║   Bank Stress Testing Lab — Prof. V. Ravichandran               ║
╚══════════════════════════════════════════════════════════════════╝
"""

import streamlit as st
import io
import base64
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
from scipy import stats
import warnings
warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Bank Stress Testing Lab | Mountain Path",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ─────────────────────────────────────────────────────────────────
# MOUNTAIN PATH DESIGN SYSTEM
# ─────────────────────────────────────────────────────────────────

def hex_to_rgba(hex_color, alpha=0.08):
    """Convert hex colour string to rgba() -- safe for all Plotly versions."""
    h = hex_color.lstrip("#")
    if len(h) == 6:
        r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    else:
        r, g, b = 173, 216, 230
    return f"rgba({r},{g},{b},{alpha:.2f})"


# ─────────────────────────────────────────────────────────────────
# EXCEL TEMPLATE GENERATOR
# ─────────────────────────────────────────────────────────────────
def create_excel_template():
    """Build and return the Bank Stress Test input template as bytes."""
    wb = Workbook()

    C_DB = "00003366"; C_MB = "00004d80"; C_GOLD = "00FFD700"
    C_LB = "00ADD8E6"; C_BG = "001A2A4A"; C_INP = "00EBF5FB"
    C_FML = "00F0FFF0"; C_WH = "00FFFFFF"
    thin   = Side(style="thin",   color="004d80")
    gold_s = Side(style="medium", color="FFD700")

    def _bdr():  return Border(left=thin, right=thin, top=thin, bottom=thin)
    def _fill(c): return PatternFill("solid", fgColor=c)
    def _ctr():  return Alignment(horizontal="center", vertical="center", wrap_text=True)
    def _lft():  return Alignment(horizontal="left",   vertical="center", wrap_text=True)
    def _rgt():  return Alignment(horizontal="right",  vertical="center")

    def _sheet_header(ws, title):
        ws.row_dimensions[1].height = 32
        ws.merge_cells("B1:F1")
        c = ws["B1"]; c.value = title
        c.font = Font(name="Calibri", size=13, bold=True, color=C_GOLD)
        c.fill = _fill(C_DB); c.alignment = _lft()
        for col, hdr in enumerate(["Field", "Unit", "► Enter Value Here ◄", "Notes / RBI Minimum", "Source"], start=2):
            cc = ws.cell(row=2, column=col, value=hdr)
            cc.font = Font(name="Calibri", size=10, bold=True, color=C_WH)
            cc.fill = _fill(C_MB); cc.alignment = _ctr()
            cc.border = Border(left=thin, right=thin, top=thin,
                               bottom=Side(style="medium", color="FFD700"))

    def _section(ws, row, title):
        ws.row_dimensions[row].height = 20
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=6)
        c = ws.cell(row=row, column=2, value=title)
        c.font = Font(name="Calibri", size=10, bold=True, color=C_GOLD)
        c.fill = _fill(C_DB)
        c.border = Border(left=gold_s, right=thin, top=thin, bottom=thin)
        c.alignment = _lft()

    def _inp(ws, row, label, unit, value, note=""):
        ws.row_dimensions[row].height = 17
        lc = ws.cell(row=row, column=2, value=label)
        lc.font = Font(name="Calibri", size=10, color="00D0DFF0"); lc.fill = _fill(C_BG)
        lc.alignment = _lft(); lc.border = _bdr()
        uc = ws.cell(row=row, column=3, value=unit)
        uc.font = Font(name="Calibri", size=9, color="00ADD8E6"); uc.fill = _fill(C_BG)
        uc.alignment = _ctr(); uc.border = _bdr()
        vc = ws.cell(row=row, column=4, value=value)
        vc.font = Font(name="Calibri", size=10, bold=True, color="00003366")
        vc.fill = _fill(C_INP)
        vc.border = Border(left=gold_s, right=thin, top=thin, bottom=thin)
        vc.alignment = _rgt()
        if isinstance(value, (int, float)):
            vc.number_format = "#,##0.00"
        nc = ws.cell(row=row, column=5, value=note)
        nc.font = Font(name="Calibri", size=8, italic=True, color="00555555")
        nc.fill = _fill("00F5F5F5"); nc.alignment = _lft(); nc.border = _bdr()
        ws.cell(row=row, column=6).fill = _fill("00F5F5F5")
        ws.cell(row=row, column=6).border = _bdr()

    def _calc(ws, row, label, unit, formula, note=""):
        """Auto-calculated (formula) row — light green, locked visually."""
        ws.row_dimensions[row].height = 17
        lc = ws.cell(row=row, column=2, value=label)
        lc.font = Font(name="Calibri", size=10, bold=True, color=C_WH)
        lc.fill = _fill(C_MB); lc.alignment = _lft(); lc.border = _bdr()
        uc = ws.cell(row=row, column=3, value=unit)
        uc.font = Font(name="Calibri", size=9, color="00ADD8E6")
        uc.fill = _fill(C_MB); uc.alignment = _ctr(); uc.border = _bdr()
        fc = ws.cell(row=row, column=4, value=formula)
        fc.font = Font(name="Calibri", size=10, italic=True, color="00004400")
        fc.fill = _fill(C_FML)
        fc.border = Border(left=gold_s, right=thin, top=thin, bottom=thin)
        fc.alignment = _rgt(); fc.number_format = "#,##0.00"
        nc = ws.cell(row=row, column=5, value="🔒 Auto-calculated — do not edit" if not note else note)
        nc.font = Font(name="Calibri", size=8, italic=True, color="00228800")
        nc.fill = _fill(C_FML); nc.alignment = _lft(); nc.border = _bdr()
        ws.cell(row=row, column=6).fill = _fill(C_FML)
        ws.cell(row=row, column=6).border = _bdr()

    def _col_widths(ws, widths):
        for col_letter, w in widths.items():
            ws.column_dimensions[col_letter].width = w

    # ════════════════════════════════════════════════════════════
    # SHEET 0: INSTRUCTIONS
    # ════════════════════════════════════════════════════════════
    ws0 = wb.active; ws0.title = "📋 Instructions"
    ws0.sheet_properties.tabColor = "FFD700"
    _col_widths(ws0, {"A":3,"B":55,"C":30})
    ws0.row_dimensions[1].height = 36
    ws0.merge_cells("B1:C1")
    t = ws0["B1"]; t.value = "🏦  THE MOUNTAIN PATH — Bank Stress Test Input Template"
    t.font = Font(name="Calibri", size=14, bold=True, color=C_GOLD)
    t.fill = _fill(C_DB); t.alignment = _lft()
    lines = [
        (2, "Prof. V. Ravichandran  |  28+ Yrs Corporate Finance & Banking  |  10+ Yrs Academic Excellence",
            False, "00ADD8E6", C_DB, 10),
        (4, "HOW TO USE THIS TEMPLATE", True, C_GOLD, C_DB, 11),
        (5, "1.  Fill ONLY the light-blue ► Enter Value Here ◄ cells (column D on each sheet).", False, "00003366","00EBF5FB",10),
        (6, "2.  Do NOT change row order, field names, or sheet names — the parser reads fixed positions.", False,"00003366","00EBF5FB",10),
        (7, "3.  All monetary values in ₹ CRORE unless the unit column says otherwise.", False,"00003366","00EBF5FB",10),
        (8, "4.  Percentages: enter as plain numbers — e.g. type 7.5 for 7.5%, NOT 0.075.", False,"00003366","00EBF5FB",10),
        (9, "5.  Light-green cells are AUTO-CALCULATED from your inputs — do not edit them.", False,"00003366","00EBF5FB",10),
        (10,"6.  After filling all sheets, save the file and upload to the Stress Testing Lab.", False,"00003366","00EBF5FB",10),
        (12,"SHEET GUIDE", True, C_GOLD, C_DB, 11),
        (13,"📊 BS_Assets        →  Balance Sheet: Assets (loans, investments, HQLA, NPA)", False,"00D0DFF0",C_BG,10),
        (14,"📊 BS_Liabilities   →  Balance Sheet: Liabilities & Capital (deposits, CET1, RWA)", False,"00D0DFF0",C_BG,10),
        (15,"📊 Income_Stmt      →  Annual P&L — NII, fees, provisions, PAT (auto-calculates PPOP/PBT/PAT)", False,"00D0DFF0",C_BG,10),
        (16,"📊 Asset_Quality    →  NPA ratios, provision coverage, sector concentrations (₹ Crore)", False,"00D0DFF0",C_BG,10),
        (17,"📊 Capital_Ratios   →  CET1, CRAR, LCR, NSFR, NIM, ROE, ROA, Cost-to-Income", False,"00D0DFF0",C_BG,10),
        (18,"📊 Duration_Risk    →  AFS/HTM duration, repricing gap, equity beta, net FX position", False,"00D0DFF0",C_BG,10),
        (20,"COLOUR CODING", True, C_GOLD, C_DB, 11),
        (21,"🔵 Light Blue   =  USER INPUT — enter your bank data here (column D)", False,"00003366","00EBF5FB",10),
        (22,"🟢 Light Green  =  AUTO-CALCULATED formulas — do not edit these", False,"00004400","00F0FFF0",10),
        (23,"🔵 Dark Blue    =  Labels, section headers — do not edit", False,"00D0DFF0",C_BG,10),
        (25,"LINKED CALCULATIONS IN Income_Stmt SHEET", True, C_GOLD, C_DB, 11),
        (26,"Total Operating Income  =  NII + Fee + Trading + Other Income  (auto)", False,"00003366","00EBF5FB",10),
        (27,"PPOP  =  Total Income − Operating Costs  (auto)", False,"00003366","00EBF5FB",10),
        (28,"PBT   =  PPOP − Loan Loss Provisions − Other Provisions  (auto)", False,"00003366","00EBF5FB",10),
        (29,"Tax   =  PBT × 25%  (edit the rate in the sheet if your effective rate differs)", False,"00003366","00EBF5FB",10),
        (30,"PAT   =  PBT − Tax  (auto)", False,"00003366","00EBF5FB",10),
        (32,"CROSS-SHEET AUTO-CALCULATIONS (Capital_Ratios sheet)", True, C_GOLD, C_DB, 11),
        (33,"CET1 Ratio       =  BS_Liabilities CET1 ÷ RWA × 100  (auto)", False,"00003366","00EBF5FB",10),
        (34,"Total CRAR       =  Total Capital ÷ RWA × 100  (auto)", False,"00003366","00EBF5FB",10),
        (35,"CASA Ratio       =  BS_Liabilities CASA ÷ Total Deposits × 100  (auto)", False,"00003366","00EBF5FB",10),
        (36,"NIM              =  Income_Stmt NII ÷ BS_Assets Total Assets × 100  (auto)", False,"00003366","00EBF5FB",10),
        (37,"ROA              =  Income_Stmt PAT ÷ BS_Assets Total Assets × 100  (auto)", False,"00003366","00EBF5FB",10),
        (38,"Gross NPA Ratio  =  BS_Assets Gross NPA ÷ Gross Loans × 100  (auto)", False,"00003366","00EBF5FB",10),
        (39,"PCR              =  BS_Assets Provisions ÷ Gross NPA × 100  (auto)", False,"00003366","00EBF5FB",10),
        (40,"Net Repricing Gap = RSA − RSL  (auto)", False,"00003366","00EBF5FB",10),
    ]
    for row, txt, bold, fc, bg, sz in lines:
        ws0.row_dimensions[row].height = 18
        ws0.merge_cells(f"B{row}:C{row}")
        c = ws0[f"B{row}"]; c.value = txt
        c.font = Font(name="Calibri", size=sz, bold=bold, color=fc)
        c.fill = _fill(bg); c.alignment = _lft()

    # ════════════════════════════════════════════════════════════
    # SHEET 1: BS_ASSETS   (col D = user input)
    # Row map (fixed — parser reads these exact rows):
    #  4=Bank Name, 5=Type, 6=FY, 7=Rating
    #  10=Total Assets, 11=Gross Loans, 12=Retail, 13=Corp, 14=MSME, 15=Agri, 16=Other
    #  17=Gross NPA, 18=Net NPA, 19=Provisions
    #  22=HTM, 23=AFS, 24=HFT, 25=Equity Portfolio
    #  28=Cash/RBI, 29=HQLA, 30=Fixed Assets, 31=Other Assets
    # ════════════════════════════════════════════════════════════
    ws1 = wb.create_sheet("📊 BS_Assets")
    ws1.sheet_properties.tabColor = "003366"
    _col_widths(ws1, {"A":3,"B":36,"C":12,"D":18,"E":40,"F":22})
    _sheet_header(ws1, "BALANCE SHEET — ASSETS  (₹ Crore)")
    r = 3
    _section(ws1, r, "BANK IDENTIFICATION"); r+=1      # r=4
    for lbl,unit,val,note in [
        ("Bank Name",       "Text",  "Enter Bank Name",       "Official registered name"),
        ("Bank Type",       "Text",  "Private / PSU / SFB",   "Scheduled Commercial Bank type"),
        ("Financial Year",  "Text",  "FY2024-25",             "Period of financial data"),
        ("Credit Rating",   "Text",  "AA- / AA / A+",         "CRISIL / ICRA / CARE rating"),
    ]:
        _inp(ws1, r, lbl, unit, val, note); r+=1       # rows 4-7; r=8
    r+=1                                                # blank row 8, section at 9
    _section(ws1, r, "CREDIT PORTFOLIO"); r+=1         # section=9, r=10
    for lbl,unit,val,note in [
        ("Total Assets",            "₹ Cr", 185000, "Sum of all balance sheet assets"),
        ("Gross Loans & Advances",  "₹ Cr", 108000, "Total gross loan book before provisions"),
        ("  ► Retail Loans",        "₹ Cr", 38000,  "Home, personal, vehicle, gold loans"),
        ("  ► Corporate Loans",     "₹ Cr", 45000,  "Large corporate and mid-market"),
        ("  ► MSME Loans",          "₹ Cr", 18000,  "Micro, small and medium enterprises"),
        ("  ► Agriculture Loans",   "₹ Cr", 7000,   "Priority sector agricultural credit"),
        ("  ► Other Loans",         "₹ Cr", 0,      "Other loan segments"),
        ("Gross NPA",               "₹ Cr", 7560,   "Total gross non-performing assets"),
        ("Net NPA",                 "₹ Cr", 3780,   "Gross NPA minus provisions held"),
        ("Provisions Held",         "₹ Cr", 3780,   "Total loan loss provisions on balance sheet"),
    ]:
        _inp(ws1, r, lbl, unit, val, note); r+=1      # rows 10-19; r=20
    r+=1                                               # blank 20, section at 21
    _section(ws1, r, "INVESTMENTS"); r+=1              # section=21, r=22
    for lbl,unit,val,note in [
        ("HTM Investments",   "₹ Cr", 28000, "Held-to-maturity (G-secs, SDLs)"),
        ("AFS Investments",   "₹ Cr", 14000, "Available-for-sale — mark-to-market risk"),
        ("HFT Investments",   "₹ Cr", 0,     "Held-for-trading"),
        ("Equity Portfolio",  "₹ Cr", 3200,  "Listed equity shares and ETFs"),
    ]:
        _inp(ws1, r, lbl, unit, val, note); r+=1      # rows 22-25; r=26
    r+=1                                               # blank 26, section at 27
    _section(ws1, r, "OTHER ASSETS"); r+=1             # section=27, r=28
    for lbl,unit,val,note in [
        ("Cash & Balances with RBI", "₹ Cr", 12000, "CRR balances + vault cash"),
        ("HQLA / Liquid Assets",     "₹ Cr", 22000, "High quality liquid assets for LCR computation"),
        ("Fixed Assets",             "₹ Cr", 1800,  "Premises, equipment, computer systems"),
        ("Other Assets",             "₹ Cr", 8000,  "Deferred tax asset, intangibles, sundry"),
    ]:
        _inp(ws1, r, lbl, unit, val, note); r+=1      # rows 28-31

    # ════════════════════════════════════════════════════════════
    # SHEET 2: BS_LIABILITIES
    # Row map:
    #  4=Total Deposits, 5=CASA, 6=Term
    #  9=Wholesale, 10=Sub-debt, 11=Other Liab
    #  14=Share Capital, 15=Reserves, 16=CET1, 17=AT1, 18=Tier2
    #  19=Total Capital, 20=RWA, 21=Net Worth
    # ════════════════════════════════════════════════════════════
    ws2 = wb.create_sheet("📊 BS_Liabilities")
    ws2.sheet_properties.tabColor = "003366"
    _col_widths(ws2, {"A":3,"B":36,"C":12,"D":18,"E":40,"F":22})
    _sheet_header(ws2, "BALANCE SHEET — LIABILITIES & CAPITAL  (₹ Crore)")
    r = 3
    _section(ws2, r, "DEPOSITS"); r+=1                # section=3, r=4
    for lbl,unit,val,note in [
        ("Total Deposits",      "₹ Cr", 148000, "All customer deposits"),
        ("  ► CASA Deposits",   "₹ Cr", 59200,  "Current + savings accounts (low-cost)"),
        ("  ► Term Deposits",   "₹ Cr", 88800,  "Fixed deposits by retail & corporate"),
    ]:
        _inp(ws2, r, lbl, unit, val, note); r+=1     # rows 4-6; r=7
    r+=1                                              # blank 7, section at 8
    _section(ws2, r, "BORROWINGS"); r+=1              # section=8, r=9
    for lbl,unit,val,note in [
        ("Wholesale / Market Funding", "₹ Cr", 18500, "CPs, NCDs, PSL bonds, interbank"),
        ("Subordinated Debt",          "₹ Cr", 3200,  "Lower Tier 2 / sub-debt bonds"),
        ("Other Liabilities",          "₹ Cr", 1000,  "Provisions payable, deferred tax"),
    ]:
        _inp(ws2, r, lbl, unit, val, note); r+=1     # rows 9-11; r=12
    r+=1                                              # blank 12, section at 13
    _section(ws2, r, "CAPITAL & RESERVES"); r+=1      # section=13, r=14
    for lbl,unit,val,note in [
        ("Share Capital",              "₹ Cr", 1200,  "Paid-up equity share capital"),
        ("Reserves & Surplus",         "₹ Cr", 12300, "Retained earnings + statutory reserves"),
        ("CET1 Capital",               "₹ Cr", 13500, "Common Equity Tier 1"),
        ("Additional Tier 1 (AT1)",    "₹ Cr", 1500,  "Perpetual bonds / AT1 instruments"),
        ("Tier 2 Capital",             "₹ Cr", 3200,  "Sub-debt + eligible general provisions"),
    ]:
        _inp(ws2, r, lbl, unit, val, note); r+=1     # rows 14-18; r=19
    # Auto-calculated capital totals
    _calc(ws2, r, "Total Capital (Regulatory)", "₹ Cr",
          "=D16+D17+D18", "CET1 + AT1 + Tier 2"); r+=1   # row 19
    _inp(ws2,  r, "Risk-Weighted Assets (RWA)", "₹ Cr", 141000,
         "Credit + market + operational RWA"); r+=1        # row 20
    _inp(ws2,  r, "Net Worth / Equity",         "₹ Cr", 15300,
         "Share capital + all reserves"); r+=1             # row 21

    # ════════════════════════════════════════════════════════════
    # SHEET 3: INCOME STATEMENT
    # Row map:
    #  4=NII, 5=Fee, 6=Trading, 7=Other Income
    #  8=Total Income (CALC)
    #  11=Opex, 12=Loan Provisions, 13=Other Provisions
    #  16=PPOP (CALC), 17=PBT (CALC), 18=Tax (CALC), 19=PAT (CALC)
    # ════════════════════════════════════════════════════════════
    ws3 = wb.create_sheet("📊 Income_Stmt")
    ws3.sheet_properties.tabColor = "FFD700"
    _col_widths(ws3, {"A":3,"B":38,"C":12,"D":18,"E":40,"F":22})
    _sheet_header(ws3, "INCOME STATEMENT — ANNUAL  (₹ Crore)")
    r = 3
    _section(ws3, r, "INCOME"); r+=1                  # section=3, r=4
    for lbl,unit,val,note in [
        ("Net Interest Income (NII)",   "₹ Cr", 7200, "Interest income minus interest expense"),
        ("Fee & Commission Income",     "₹ Cr", 2100, "Transaction fees, trade finance, WM"),
        ("Trading & MTM Income",        "₹ Cr", 680,  "Treasury P&L, bond gains/losses"),
        ("Other Income",                "₹ Cr", 420,  "FX income, recoveries, miscellaneous"),
    ]:
        _inp(ws3, r, lbl, unit, val, note); r+=1     # rows 4-7; r=8
    _calc(ws3, r, "Total Operating Income", "₹ Cr",
          "=SUM(D4:D7)", "NII + Fee + Trading + Other"); r+=1   # row 8; r=9
    r+=1                                              # blank row 9, section at 10
    _section(ws3, r, "EXPENSES"); r+=1                # section=10, r=11
    for lbl,unit,val,note in [
        ("Operating / Staff Costs", "₹ Cr", 4800, "Employee costs + admin + depreciation"),
        ("Loan Loss Provisions",    "₹ Cr", 2800, "Provisions for NPAs and standard assets"),
        ("Other Provisions",        "₹ Cr", 0,    "Investment depreciation, contingency prov."),
    ]:
        _inp(ws3, r, lbl, unit, val, note); r+=1     # rows 11-13; r=14
    r+=1                                              # blank 14, section at 15
    _section(ws3, r, "PROFIT SUMMARY — AUTO-CALCULATED"); r+=1  # section=15, r=16
    _calc(ws3, r, "Pre-Provision Operating Profit (PPOP)", "₹ Cr",
          "=D8-D11", "Total Income − Operating Costs"); r+=1         # row 16
    _calc(ws3, r, "Profit Before Tax (PBT)", "₹ Cr",
          "=D16-D12-D13", "PPOP − Loan Provisions − Other Prov."); r+=1  # row 17
    _calc(ws3, r, "Income Tax (est. 25%)", "₹ Cr",
          "=D17*0.25", "Effective tax ≈25% — edit multiplier if needed"); r+=1  # row 18
    _calc(ws3, r, "Profit After Tax (PAT)", "₹ Cr",
          "=D17-D18", "PBT − Income Tax"); r+=1                      # row 19

    # ════════════════════════════════════════════════════════════
    # SHEET 4: ASSET QUALITY
    # Row map:
    #  4=Gross NPA%, 5=Net NPA%, 6=SMA2%, 7=Restructured%, 8=PCR%, 9=CD Ratio%
    #  (rows 4-9 — user enters; auto versions added below as calc rows)
    #  12=Real Estate, 13=Infra, 14=NBFC, 15=Power, 16=Textile, 17=Gems
    #  18=Iron&Steel, 19=Aviation, 20=Other
    # ════════════════════════════════════════════════════════════
    ws4 = wb.create_sheet("📊 Asset_Quality")
    ws4.sheet_properties.tabColor = "DC3545"
    _col_widths(ws4, {"A":3,"B":36,"C":12,"D":18,"E":40,"F":22})
    _sheet_header(ws4, "ASSET QUALITY & SECTOR EXPOSURE")
    r = 3
    _section(ws4, r, "NPA & ASSET QUALITY RATIOS — enter OR leave 0 to auto-calculate"); r+=1  # r=4
    for lbl,unit,val,note in [
        ("Gross NPA Ratio",              "%", 7.0,  "If 0, auto-calculated from BS_Assets"),
        ("Net NPA Ratio",                "%", 3.5,  "If 0, auto-calculated from BS_Assets"),
        ("SMA-2 Ratio",                  "%", 3.2,  "SMA-2 / Gross Loans × 100"),
        ("Restructured Assets Ratio",    "%", 1.8,  "Restructured book / Gross Loans × 100"),
        ("Provision Coverage Ratio (PCR)","%",50.0, "If 0, auto-calculated from BS_Assets"),
        ("Credit-Deposit Ratio",         "%", 72.9, "If 0, auto-calculated from BS sheets"),
    ]:
        _inp(ws4, r, lbl, unit, val, note); r+=1    # rows 4-9; r=10
    r+=1                                             # blank 10, section at 11
    _section(ws4, r, "SECTOR CONCENTRATION (₹ Crore) — loan outstanding to each sector"); r+=1  # r=12
    for lbl,unit,val,note in [
        ("Real Estate & Construction", "₹ Cr", 12000, "Developer loans + project finance"),
        ("Infrastructure",             "₹ Cr", 18000, "Roads, power, ports, telecom"),
        ("NBFC & HFC",                 "₹ Cr", 9500,  "Loans to NBFCs and housing finance cos"),
        ("Power Sector",               "₹ Cr", 8800,  "Thermal, hydro, renewable energy"),
        ("Textile",                    "₹ Cr", 5500,  "Spinning, weaving, garments"),
        ("Gems & Jewellery",           "₹ Cr", 3200,  "Diamond, gold jewellery"),
        ("Iron & Steel",               "₹ Cr", 0,     "Steel, alloys, metals"),
        ("Aviation",                   "₹ Cr", 0,     "Airlines, airports, MRO"),
        ("Other Sectors",              "₹ Cr", 0,     "All remaining sectors"),
    ]:
        _inp(ws4, r, lbl, unit, val, note); r+=1    # rows 12-20

    # ════════════════════════════════════════════════════════════
    # SHEET 5: CAPITAL RATIOS  (cross-sheet formula auto-calcs)
    # Row map:
    #  4=CET1%, 5=Tier1%, 6=CRAR%, 7=Leverage%
    #  10=LCR%, 11=NSFR%, 12=CASA%
    #  15=NIM%, 16=ROE%, 17=ROA%, 18=Cost-Income%
    # ════════════════════════════════════════════════════════════
    ws5 = wb.create_sheet("📊 Capital_Ratios")
    ws5.sheet_properties.tabColor = "28A745"
    _col_widths(ws5, {"A":3,"B":36,"C":12,"D":18,"E":40,"F":22})
    _sheet_header(ws5, "CAPITAL ADEQUACY & LIQUIDITY RATIOS")
    r = 3
    _section(ws5, r, "CAPITAL ADEQUACY — enter values OR leave 0 to auto-calculate"); r+=1  # r=4
    # CET1 — user can enter; also show cross-sheet auto-calc
    _inp(ws5, r, "CET1 Ratio",    "%", 9.57,  "RBI min 8.5% incl. conservation buffer"); r+=1   # row 4
    _inp(ws5, r, "Tier 1 Ratio",  "%", 10.64, "RBI minimum 9.5%"); r+=1                          # row 5
    _inp(ws5, r, "Total CRAR",    "%", 12.9,  "RBI minimum 11.5% incl. CCB"); r+=1               # row 6
    _inp(ws5, r, "Leverage Ratio","%", 8.1,   "RBI minimum 3.5% (Tier 1 / Total Exposure)"); r+=1  # row 7
    r+=1                                             # blank 8, section at 9
    _section(ws5, r, "AUTO-CALCULATED FROM BALANCE SHEET"); r+=1  # section=9, r=10
    _calc(ws5, r, "CET1 Ratio (from BS)", "%",
          "=IF('📊 BS_Liabilities'!D20>0,'📊 BS_Liabilities'!D16/'📊 BS_Liabilities'!D20*100,0)",
          "= BS CET1 ÷ RWA × 100"); r+=1            # row 10
    _calc(ws5, r, "Total CRAR (from BS)", "%",
          "=IF('📊 BS_Liabilities'!D20>0,('📊 BS_Liabilities'!D16+'📊 BS_Liabilities'!D17+'📊 BS_Liabilities'!D18)/'📊 BS_Liabilities'!D20*100,0)",
          "= (CET1+AT1+Tier2) ÷ RWA × 100"); r+=1  # row 11
    _calc(ws5, r, "CASA Ratio (from BS)",  "%",
          "=IF('📊 BS_Liabilities'!D4>0,'📊 BS_Liabilities'!D5/'📊 BS_Liabilities'!D4*100,0)",
          "= CASA Deposits ÷ Total Deposits × 100"); r+=1  # row 12
    _calc(ws5, r, "NIM (from BS+P&L)",     "%",
          "=IF('📊 BS_Assets'!D10>0,'📊 Income_Stmt'!D4/'📊 BS_Assets'!D10*100,0)",
          "= NII ÷ Total Assets × 100"); r+=1       # row 13
    _calc(ws5, r, "ROA (from BS+P&L)",     "%",
          "=IF('📊 BS_Assets'!D10>0,'📊 Income_Stmt'!D19/'📊 BS_Assets'!D10*100,0)",
          "= PAT ÷ Total Assets × 100"); r+=1       # row 14
    _calc(ws5, r, "Gross NPA Ratio (from BS)", "%",
          "=IF('📊 BS_Assets'!D11>0,'📊 BS_Assets'!D17/'📊 BS_Assets'!D11*100,0)",
          "= Gross NPA ÷ Gross Loans × 100"); r+=1  # row 15
    _calc(ws5, r, "PCR (from BS)",         "%",
          "=IF('📊 BS_Assets'!D17>0,'📊 BS_Assets'!D19/'📊 BS_Assets'!D17*100,0)",
          "= Provisions ÷ Gross NPA × 100"); r+=1   # row 16
    r+=1                                             # blank, section
    _section(ws5, r, "LIQUIDITY & PROFITABILITY — enter your reported values"); r+=1  # r=18
    _inp(ws5, r, "LCR (Liquidity Coverage Ratio)", "%", 142.0, "RBI minimum 100%"); r+=1   # row 18
    _inp(ws5, r, "NSFR (Net Stable Funding Ratio)","%", 118.0, "RBI minimum 100%"); r+=1   # row 19
    _inp(ws5, r, "Return on Equity (ROE)",          "%", 13.7,  "PAT ÷ Average equity × 100"); r+=1  # row 20
    _inp(ws5, r, "Cost-to-Income Ratio",            "%", 46.2,  "Operating costs ÷ Total income × 100"); r+=1  # row 21

    # ════════════════════════════════════════════════════════════
    # SHEET 6: DURATION RISK
    # Row map:
    #  4=AFS Duration, 5=HTM Duration, 6=Liab Duration, 7=Duration Gap
    #  10=RSA, 11=RSL, 12=Net Repricing Gap (CALC), 13=Equity Beta, 14=FX Position
    # ════════════════════════════════════════════════════════════
    ws6 = wb.create_sheet("📊 Duration_Risk")
    ws6.sheet_properties.tabColor = "ADD8E6"
    _col_widths(ws6, {"A":3,"B":36,"C":12,"D":18,"E":40,"F":22})
    _sheet_header(ws6, "DURATION, MARKET RISK & REPRICING GAPS")
    r = 3
    _section(ws6, r, "INVESTMENT PORTFOLIO — DURATION RISK"); r+=1   # r=4
    for lbl,unit,val,note in [
        ("Avg. Modified Duration — AFS",          "Years", 4.8, "Key driver of MTM loss under rate shock"),
        ("Avg. Modified Duration — HTM",          "Years", 6.2, "No MTM impact — economic risk only"),
        ("Avg. Modified Duration — Liabilities",  "Years", 2.1, "Weighted avg. duration of deposits + borrowings"),
    ]:
        _inp(ws6, r, lbl, unit, val, note); r+=1    # rows 4-6; r=7
    _calc(ws6, r, "Duration Gap (Assets − Liabilities)", "Years",
          "=D4-D6", "Positive = rate rise hurts equity value"); r+=1  # row 7; r=8
    r+=1                                             # blank 8, section at 9
    _section(ws6, r, "INTEREST RATE REPRICING GAPS"); r+=1   # section=9, r=10
    _inp(ws6, r, "Rate-Sensitive Assets (RSA)",        "₹ Cr", 95000, "Assets repricing within 1 year"); r+=1   # row 10
    _inp(ws6, r, "Rate-Sensitive Liabilities (RSL)",   "₹ Cr", 76500, "Liabilities repricing within 1 year"); r+=1  # row 11
    _calc(ws6, r, "Net Repricing Gap (RSA − RSL)",     "₹ Cr",
          "=D10-D11", "Positive = asset-sensitive bank (NII rises when rates rise)"); r+=1  # row 12
    _inp(ws6, r, "Equity Portfolio Beta",              "β",    0.95,  "Weighted avg beta vs Nifty 50"); r+=1    # row 13
    _inp(ws6, r, "Net Open FX Position",               "₹ Cr", 800,   "Net forex exposure (+ = long USD)"); r+=1  # row 14

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()

def parse_uploaded_excel(file_bytes):
    """Parse uploaded Excel template → bank data dict. Returns (bank_dict, errors).
    Row numbers here MUST match the template layout in create_excel_template().
    """
    errors = []
    try:
        wb = load_workbook(filename=io.BytesIO(file_bytes), data_only=True)
    except Exception as e:
        return None, [f"Cannot open file: {e}"]

    required_sheets = ["📊 BS_Assets", "📊 BS_Liabilities", "📊 Income_Stmt",
                       "📊 Asset_Quality", "📊 Capital_Ratios", "📊 Duration_Risk"]
    missing = [s for s in required_sheets if s not in wb.sheetnames]
    if missing:
        return None, [f"Missing sheets: {', '.join(missing)}. Please use the official template."]

    def _val(ws, row, col=4, fallback=0.0):
        v = ws.cell(row=row, column=col).value
        if v is None: return fallback
        try:   return float(v)
        except: return fallback

    def _str(ws, row, col=4, fallback=""):
        v = ws.cell(row=row, column=col).value
        return str(v).strip() if v else fallback

    wa = wb["📊 BS_Assets"]
    wl = wb["📊 BS_Liabilities"]
    wi = wb["📊 Income_Stmt"]
    wq = wb["📊 Asset_Quality"]
    wc = wb["📊 Capital_Ratios"]
    wd = wb["📊 Duration_Risk"]

    # ── BS_Assets (verified row map) ───────────────────────────
    # Identity: 4=Name, 5=Type, 6=FY, 7=Rating
    # Credit:   10=TotalAssets, 11=GrossLoans, 12=Retail, 13=Corp, 14=MSME,
    #           15=Agri, 16=Other, 17=GrossNPA, 18=NetNPA, 19=Provisions
    # Invest:   22=HTM, 23=AFS, 24=HFT, 25=Equity
    # Other:    28=Cash, 29=HQLA, 30=Fixed, 31=Other
    # ────────────────────────────────────────────────────────────
    # ── BS_Liabilities (verified row map) ──────────────────────
    # Deposits: 4=TotalDep, 5=CASA, 6=Term
    # Borrow:   9=Wholesale, 10=SubDebt, 11=OtherLiab
    # Capital:  14=ShareCap, 15=Reserves, 16=CET1, 17=AT1, 18=Tier2
    #           19=TotalCap(CALC), 20=RWA, 21=NetWorth
    # ────────────────────────────────────────────────────────────
    # ── Income_Stmt (verified row map) ─────────────────────────
    # Income: 4=NII, 5=Fee, 6=Trading, 7=OtherIncome
    #         8=TotalIncome(CALC)
    # Expense:11=Opex, 12=LoanProv, 13=OtherProv
    # Profit: 16=PPOP(CALC), 17=PBT(CALC), 18=Tax(CALC), 19=PAT(CALC)
    # ────────────────────────────────────────────────────────────
    # ── Asset_Quality (verified row map) ───────────────────────
    # Ratios: 4=GrossNPA%, 5=NetNPA%, 6=SMA2%, 7=Restruc%, 8=PCR%, 9=CD%
    # Sector: 12=RealEstate, 13=Infra, 14=NBFC, 15=Power, 16=Textile,
    #         17=Gems, 18=Steel, 19=Aviation, 20=Other
    # ────────────────────────────────────────────────────────────
    # ── Capital_Ratios (verified row map) ──────────────────────
    # Entered: 4=CET1%, 5=Tier1%, 6=CRAR%, 7=Leverage%
    # Calc:   10=CET1%(BS), 11=CRAR%(BS), 12=CASA%(BS), 13=NIM%(BS), 
    #         14=ROA%(BS), 15=GrossNPA%(BS), 16=PCR%(BS)
    # Liquidity/Prof: 18=LCR, 19=NSFR, 20=ROE, 21=CostIncome
    # ────────────────────────────────────────────────────────────
    # ── Duration_Risk (verified row map) ───────────────────────
    # Duration: 4=AFS_dur, 5=HTM_dur, 6=Liab_dur, 7=DurGap(CALC)
    # Repricing:10=RSA, 11=RSL, 12=Gap(CALC), 13=Beta, 14=FX

    bank = {
        # Identity
        "name":   _str(wa, 4) or "Uploaded Bank",
        "type":   _str(wa, 5) or "Scheduled Commercial Bank",
        "rating": _str(wa, 7) or "N/A",

        # ── Balance Sheet: Assets ──
        "total_assets":       _val(wa, 10),
        "gross_loans":        _val(wa, 11),
        "retail_loans":       _val(wa, 12),
        "corporate_loans":    _val(wa, 13),
        "msme_loans":         _val(wa, 14),
        "agri_loans":         _val(wa, 15),
        "gross_npa":          _val(wa, 17),
        "net_npa":            _val(wa, 18),
        "provisions":         _val(wa, 19),
        "investments_htm":    _val(wa, 22),
        "investments_afs":    _val(wa, 23),
        "equity_portfolio":   _val(wa, 25),
        "cash_hqla":          _val(wa, 29),
        "fixed_assets":       _val(wa, 30),
        "other_assets":       _val(wa, 31),

        # ── Balance Sheet: Liabilities ──
        "total_deposits":     _val(wl, 4),
        "casa_deposits":      _val(wl, 5),
        "term_deposits":      _val(wl, 6),
        "wholesale_funding":  _val(wl, 9),
        "sub_debt":           _val(wl, 10),
        "cet1_capital":       _val(wl, 16),
        "tier1_capital":      _val(wl, 16) + _val(wl, 17),   # CET1 + AT1
        "tier2_capital":      _val(wl, 18),
        "total_capital":      _val(wl, 16) + _val(wl, 17) + _val(wl, 18),
        "rwa":                _val(wl, 20),
        "equity_capital":     _val(wl, 21),

        # ── Income Statement ──
        "net_interest_income": _val(wi, 4),
        "fee_income":          _val(wi, 5),
        "trading_income":      _val(wi, 6),
        "other_income":        _val(wi, 7),
        "operating_costs":     _val(wi, 11),
        "provisions_charge":   _val(wi, 12),
        "pre_provision_profit":_val(wi, 16),   # PPOP calc cell
        "pbt":                 _val(wi, 17),
        "tax":                 _val(wi, 18),
        "pat":                 _val(wi, 19),

        # ── Asset Quality ──
        "gross_npa_ratio":     _val(wq, 4),
        "net_npa_ratio":       _val(wq, 5),
        "slma_ratio":          _val(wq, 6),
        "restructured_ratio":  _val(wq, 7),
        "pcr":                 _val(wq, 8),
        "cd_ratio":            _val(wq, 9),
        "sector_real_estate":  _val(wq, 12),
        "sector_infrastructure":_val(wq, 13),
        "sector_nbfc":         _val(wq, 14),
        "sector_power":        _val(wq, 15),
        "sector_textile":      _val(wq, 16),
        "sector_gems":         _val(wq, 17),

        # ── Capital Ratios (prefer user-entered; fallback to BS calc) ──
        "cet1_ratio":    _val(wc, 4) or _val(wc, 10),
        "crar":          _val(wc, 6) or _val(wc, 11),
        "casa_ratio":    _val(wc, 12) if _val(wc, 12) else 0,  # from calc row
        "nim":           _val(wc, 13) if _val(wc, 13) else 0,
        "roa":           _val(wc, 14) if _val(wc, 14) else 0,
        "lcr":           _val(wc, 18),
        "nsfr":          _val(wc, 19),
        "roe":           _val(wc, 20),
        "leverage_ratio":_val(wc, 7),

        # ── Duration & Market Risk ──
        "avg_duration_assets": _val(wd, 4),
        "avg_duration_liabs":  _val(wd, 6),
        "repricing_gap":       _val(wd, 12),   # calc row RSA-RSL
    }

    # ── Derive any still-zero ratios from raw balance sheet data ──
    gl = bank["gross_loans"]; gn = bank["gross_npa"]
    td = bank["total_deposits"]; ta = bank["total_assets"]
    pr = bank["provisions"]; ec = bank["equity_capital"]

    if bank["gross_npa_ratio"]  == 0 and gl > 0:
        bank["gross_npa_ratio"] = round(gn / gl * 100, 2)
    if bank["net_npa_ratio"]    == 0 and gl > 0:
        bank["net_npa_ratio"]   = round(bank["net_npa"] / gl * 100, 2)
    if bank["pcr"]              == 0 and gn > 0:
        bank["pcr"]             = round(pr / gn * 100, 1)
    if bank["casa_ratio"]       == 0 and td > 0:
        bank["casa_ratio"]      = round(bank["casa_deposits"] / td * 100, 1)
    if bank["cd_ratio"]         == 0 and td > 0:
        bank["cd_ratio"]        = round(gl / td * 100, 1)
    if bank["cet1_ratio"]       == 0 and bank["rwa"] > 0:
        bank["cet1_ratio"]      = round(bank["cet1_capital"] / bank["rwa"] * 100, 2)
    if bank["crar"]             == 0 and bank["rwa"] > 0:
        bank["crar"]            = round(bank["total_capital"]  / bank["rwa"] * 100, 2)
    if bank["nim"]              == 0 and ta > 0:
        bank["nim"]             = round(bank["net_interest_income"] / ta * 100, 2)
    # ROA/ROE derived after PAT is finalised below
    pass
    # Always re-derive P&L from raw inputs since openpyxl data_only
    # cannot evaluate Excel formulas — formula cells return None
    total_income = (bank["net_interest_income"] + bank["fee_income"] +
                    bank["trading_income"]        + bank["other_income"])
    ppop_derived = total_income - bank["operating_costs"]
    total_prov   = bank["provisions_charge"] + _val(wi, 13)   # LoanProv + OtherProv
    pbt_derived  = ppop_derived - total_prov
    tax_derived  = pbt_derived  * 0.25
    pat_derived  = pbt_derived  - tax_derived

    if bank["pre_provision_profit"] == 0 or bank["pre_provision_profit"] is None:
        bank["pre_provision_profit"] = ppop_derived
    if bank["pbt"]  == 0: bank["pbt"]  = pbt_derived
    if bank["tax"]  == 0: bank["tax"]  = tax_derived
    if bank["pat"]  == 0: bank["pat"]  = pat_derived

    if bank["repricing_gap"] == 0:
        bank["repricing_gap"] = _val(wd, 10) - _val(wd, 11)  # RSA - RSL

    # Re-derive ROA/ROE using the now-finalised PAT
    if ta > 0:
        bank["roa"] = round(bank["pat"] / ta * 100, 2)
    if ec > 0:
        bank["roe"] = round(bank["pat"] / ec * 100, 1)

    # ── Validation ──────────────────────────────────────────────
    critical = {
        "Total Assets":    bank["total_assets"],
        "Gross Loans":     bank["gross_loans"],
        "RWA":             bank["rwa"],
        "NII":             bank["net_interest_income"],
        "Total Deposits":  bank["total_deposits"],
    }
    for field, val in critical.items():
        if val <= 0:
            errors.append(f"❌ {field} is zero or missing in the template — please fill it in")

    return bank, errors
