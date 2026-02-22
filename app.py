
"""
╔══════════════════════════════════════════════════════════════════╗
║   THE MOUNTAIN PATH - WORLD OF FINANCE                          ║
║   Bank Stress Testing Lab — Prof. V. Ravichandran               ║
╚══════════════════════════════════════════════════════════════════╝
"""

import streamlit as st
import io
import base64
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
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    wb = Workbook()

    C_DB = "00003366"; C_MB = "00004d80"; C_GOLD = "00FFD700"
    C_LB = "00ADD8E6"; C_BG = "001A2A4A"; C_INP = "00EBF5FB"
    C_FML = "00F0FFF0"; C_WH = "00FFFFFF"
    thin  = Side(style="thin",   color="004d80")
    gold_s = Side(style="medium", color="FFD700")

    def _bdr(): return Border(left=thin, right=thin, top=thin, bottom=thin)
    def _fill(c): return PatternFill("solid", fgColor=c)
    def _ctr(): return Alignment(horizontal="center", vertical="center", wrap_text=True)
    def _lft(): return Alignment(horizontal="left",   vertical="center", wrap_text=True)
    def _rgt(): return Alignment(horizontal="right",  vertical="center")

    def _sheet_header(ws, title):
        ws.row_dimensions[1].height = 32
        ws.merge_cells("B1:F1")
        c = ws["B1"]; c.value = title
        c.font = Font(name="Calibri", size=13, bold=True, color=C_GOLD)
        c.fill = _fill(C_DB); c.alignment = _lft()
        for col, hdr in enumerate(["Field", "Unit", "► Enter Value Here ◄", "RBI Min / Notes", "Source"], start=2):
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

    def _inp(ws, row, label, unit, value, note="", src=""):
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
        sc = ws.cell(row=row, column=6, value=src)
        sc.font = Font(name="Calibri", size=8, color="00888888")
        sc.fill = _fill("00F5F5F5"); sc.alignment = _lft(); sc.border = _bdr()

    def _col_widths(ws, widths):
        for col_letter, w in widths.items():
            ws.column_dimensions[col_letter].width = w

    # ── SHEET 0: INSTRUCTIONS ───────────────────────────────────
    ws0 = wb.active; ws0.title = "📋 Instructions"
    ws0.sheet_properties.tabColor = "FFD700"
    _col_widths(ws0, {"A":3,"B":52,"C":30})
    ws0.row_dimensions[1].height = 36
    ws0.merge_cells("B1:C1")
    t = ws0["B1"]; t.value = "🏦  THE MOUNTAIN PATH — Bank Stress Test Input Template"
    t.font = Font(name="Calibri", size=14, bold=True, color=C_GOLD)
    t.fill = _fill(C_DB); t.alignment = _lft()
    lines = [
        (2,"Prof. V. Ravichandran  |  28+ Yrs Corporate Finance & Banking",False,"00ADD8E6",C_DB,10),
        (4,"HOW TO USE THIS TEMPLATE",True,C_GOLD,C_DB,11),
        (5,"1.  Fill ONLY the light-blue ► Enter Value Here ◄ cells.",False,"00003366","00EBF5FB",10),
        (6,"2.  Do NOT change row order, field names, or sheet names.",False,"00003366","00EBF5FB",10),
        (7,"3.  All monetary values in ₹ CRORE unless unit column says otherwise.",False,"00003366","00EBF5FB",10),
        (8,"4.  Percentages: enter as numbers — e.g. type 7.5 for 7.5%, NOT 0.075.",False,"00003366","00EBF5FB",10),
        (9,"5.  After filling, save and upload to the Stress Testing Lab app.",False,"00003366","00EBF5FB",10),
        (11,"SHEET GUIDE",True,C_GOLD,C_DB,11),
        (12,"📊 BS_Assets       →  Balance Sheet: Assets",False,"00D0DFF0",C_BG,10),
        (13,"📊 BS_Liabilities  →  Balance Sheet: Liabilities & Capital",False,"00D0DFF0",C_BG,10),
        (14,"📊 Income_Stmt     →  Annual P&L (NII, fees, provisions, PAT)",False,"00D0DFF0",C_BG,10),
        (15,"📊 Asset_Quality   →  NPA ratios and sector concentrations",False,"00D0DFF0",C_BG,10),
        (16,"📊 Capital_Ratios  →  CET1, CRAR, LCR, NSFR, profitability",False,"00D0DFF0",C_BG,10),
        (17,"📊 Duration_Risk   →  AFS duration, repricing gap, FX exposure",False,"00D0DFF0",C_BG,10),
        (19,"COLOUR CODING",True,C_GOLD,C_DB,11),
        (20,"🔵 Light Blue  =  USER INPUT — enter your bank data here",False,"00003366","00EBF5FB",10),
        (21,"🟢 Light Green =  Auto-calculated formulas — do not edit",False,"00005500","00F0FFF0",10),
        (22,"🔵 Dark Blue   =  Labels and headers — do not edit",False,"00D0DFF0",C_BG,10),
    ]
    for row,txt,bold,fc,bg,sz in lines:
        ws0.row_dimensions[row].height = 18
        ws0.merge_cells(f"B{row}:C{row}")
        c = ws0[f"B{row}"]; c.value = txt
        c.font = Font(name="Calibri", size=sz, bold=bold, color=fc)
        c.fill = _fill(bg); c.alignment = _lft()

    # ── SHEET 1: BS ASSETS ──────────────────────────────────────
    ws1 = wb.create_sheet("📊 BS_Assets")
    ws1.sheet_properties.tabColor = "003366"
    _col_widths(ws1, {"A":3,"B":36,"C":12,"D":18,"E":36,"F":22})
    _sheet_header(ws1, "BALANCE SHEET — ASSETS  (₹ Crore)")
    r = 3
    _section(ws1, r, "BANK IDENTIFICATION"); r+=1
    for lbl,unit,val,note in [
        ("Bank Name","Text","Enter Bank Name","Official registered name"),
        ("Bank Type","Text","Private / PSU / SFB","Scheduled Commercial Bank"),
        ("Financial Year","Text","FY2024-25","Period of data"),
        ("Credit Rating","Text","AA- / AA / A+","CRISIL / ICRA / CARE"),
    ]:
        _inp(ws1, r, lbl, unit, val, note); r+=1
    r+=1
    _section(ws1, r, "CREDIT PORTFOLIO"); r+=1
    for lbl,unit,val,note in [
        ("Total Assets","₹ Cr",185000,"Sum of all assets on balance sheet"),
        ("Gross Loans & Advances","₹ Cr",108000,"Total gross loan book before provisions"),
        ("  ► Retail Loans","₹ Cr",38000,"Home, personal, vehicle, gold loans"),
        ("  ► Corporate Loans","₹ Cr",45000,"Large corporate and mid-market"),
        ("  ► MSME Loans","₹ Cr",18000,"Micro, small and medium enterprises"),
        ("  ► Agriculture Loans","₹ Cr",7000,"Priority sector agri credit"),
        ("  ► Other Loans","₹ Cr",0,"Other loan segments"),
        ("Gross NPA","₹ Cr",7560,"Total gross non-performing assets"),
        ("Net NPA","₹ Cr",3780,"Gross NPA minus provisions"),
        ("Provisions Held","₹ Cr",3780,"Total loan loss provisions"),
    ]:
        _inp(ws1, r, lbl, unit, val, note); r+=1
    r+=1
    _section(ws1, r, "INVESTMENTS"); r+=1
    for lbl,unit,val,note in [
        ("HTM Investments","₹ Cr",28000,"Held-to-maturity (G-secs, SDLs)"),
        ("AFS Investments","₹ Cr",14000,"Available-for-sale (MTM risk)"),
        ("HFT Investments","₹ Cr",0,"Held-for-trading"),
        ("Equity Portfolio","₹ Cr",3200,"Equity shares and ETFs"),
    ]:
        _inp(ws1, r, lbl, unit, val, note); r+=1
    r+=1
    _section(ws1, r, "OTHER ASSETS"); r+=1
    for lbl,unit,val,note in [
        ("Cash & Balances with RBI","₹ Cr",12000,"CRR + vault cash"),
        ("HQLA / Liquid Assets","₹ Cr",22000,"High quality liquid assets (LCR)"),
        ("Fixed Assets","₹ Cr",1800,"Premises, equipment, software"),
        ("Other Assets","₹ Cr",8000,"Deferred tax, intangibles, other"),
    ]:
        _inp(ws1, r, lbl, unit, val, note); r+=1

    # ── SHEET 2: BS LIABILITIES ─────────────────────────────────
    ws2 = wb.create_sheet("📊 BS_Liabilities")
    ws2.sheet_properties.tabColor = "003366"
    _col_widths(ws2, {"A":3,"B":36,"C":12,"D":18,"E":36,"F":22})
    _sheet_header(ws2, "BALANCE SHEET — LIABILITIES & CAPITAL  (₹ Crore)")
    r = 3
    _section(ws2, r, "DEPOSITS"); r+=1
    for lbl,unit,val,note in [
        ("Total Deposits","₹ Cr",148000,"All customer deposits"),
        ("  ► CASA Deposits","₹ Cr",59200,"Current + savings accounts"),
        ("  ► Term Deposits","₹ Cr",88800,"Fixed deposits"),
    ]:
        _inp(ws2, r, lbl, unit, val, note); r+=1
    r+=1
    _section(ws2, r, "BORROWINGS"); r+=1
    for lbl,unit,val,note in [
        ("Wholesale / Market Funding","₹ Cr",18500,"CPs, NCDs, interbank borrowings"),
        ("Subordinated Debt","₹ Cr",3200,"Lower Tier 2 bonds"),
        ("Other Liabilities","₹ Cr",1000,"Provisions payable, deferred tax"),
    ]:
        _inp(ws2, r, lbl, unit, val, note); r+=1
    r+=1
    _section(ws2, r, "CAPITAL & RESERVES"); r+=1
    for lbl,unit,val,note in [
        ("Share Capital","₹ Cr",1200,"Paid-up equity share capital"),
        ("Reserves & Surplus","₹ Cr",12300,"Retained earnings + statutory reserves"),
        ("CET1 Capital","₹ Cr",13500,"Common Equity Tier 1"),
        ("Additional Tier 1 (AT1)","₹ Cr",1500,"Perpetual bonds / AT1 instruments"),
        ("Tier 2 Capital","₹ Cr",3200,"Sub-debt + general provisions"),
        ("Total Capital (Regulatory)","₹ Cr",18200,"CET1 + AT1 + Tier 2"),
        ("Risk-Weighted Assets (RWA)","₹ Cr",141000,"Credit + market + operational RWA"),
        ("Net Worth / Equity","₹ Cr",15300,"Share capital + all reserves"),
    ]:
        _inp(ws2, r, lbl, unit, val, note); r+=1

    # ── SHEET 3: INCOME STATEMENT ───────────────────────────────
    ws3 = wb.create_sheet("📊 Income_Stmt")
    ws3.sheet_properties.tabColor = "FFD700"
    _col_widths(ws3, {"A":3,"B":36,"C":12,"D":18,"E":36,"F":22})
    _sheet_header(ws3, "INCOME STATEMENT — ANNUAL  (₹ Crore)")
    r = 3
    _section(ws3, r, "INCOME"); r+=1
    inc_start = r
    for lbl,unit,val,note in [
        ("Net Interest Income (NII)","₹ Cr",7200,"Interest income − interest expense"),
        ("Fee & Commission Income","₹ Cr",2100,"Transaction fees, trade finance, WM fees"),
        ("Trading & MTM Income","₹ Cr",680,"Treasury P&L, bond gains/losses"),
        ("Other Income","₹ Cr",420,"FX income, recoveries, misc"),
    ]:
        _inp(ws3, r, lbl, unit, val, note); r+=1
    # Total income formula row
    ws3.row_dimensions[r].height = 18
    tc = ws3.cell(row=r, column=2, value="Total Operating Income")
    tc.font = Font(name="Calibri", size=10, bold=True, color=C_WH)
    tc.fill = _fill(C_MB); tc.alignment = _lft(); tc.border = _bdr()
    ws3.cell(row=r, column=3, value="₹ Cr").fill = _fill(C_MB); ws3.cell(row=r, column=3).border = _bdr()
    fc = ws3.cell(row=r, column=4,
                  value=f"=SUM(D{inc_start}:D{inc_start+3})")
    fc.font = Font(name="Calibri", size=10, italic=True, color="00005500")
    fc.fill = _fill(C_FML); fc.border = Border(left=gold_s, right=thin, top=thin, bottom=thin)
    fc.alignment = _rgt(); fc.number_format = "#,##0.00"
    ws3.cell(row=r, column=5, value="Auto-calculated").font = Font(name="Calibri", size=8, italic=True, color="00888888")
    ws3.cell(row=r, column=5).fill = _fill(C_FML); ws3.cell(row=r, column=5).border = _bdr()
    total_inc_row = r; r+=2

    _section(ws3, r, "EXPENSES"); r+=1
    exp_start = r
    for lbl,unit,val,note in [
        ("Operating / Staff Costs","₹ Cr",4800,"Employee costs + admin + depreciation"),
        ("Loan Loss Provisions","₹ Cr",2800,"Provisions for NPAs and standard assets"),
        ("Other Provisions","₹ Cr",0,"Investment depreciation, contingency"),
    ]:
        _inp(ws3, r, lbl, unit, val, note); r+=1
    r+=1

    _section(ws3, r, "PROFIT SUMMARY (AUTO-CALCULATED)"); r+=1
    ppop_row = r
    for lbl, fml in [
        ("Pre-Provision Operating Profit (PPOP)", f"=D{total_inc_row}-D{exp_start}"),
        ("Profit Before Tax (PBT)",               f"=D{ppop_row}-D{exp_start+1}-D{exp_start+2}"),
        ("Income Tax (est.)",                     f"=D{ppop_row+1}*0.25"),
        ("Profit After Tax (PAT)",                f"=D{ppop_row+1}-D{ppop_row+2}"),
    ]:
        ws3.row_dimensions[r].height = 18
        lc = ws3.cell(row=r, column=2, value=lbl)
        lc.font = Font(name="Calibri", size=10, bold=True, color=C_WH)
        lc.fill = _fill(C_MB); lc.alignment = _lft(); lc.border = _bdr()
        ws3.cell(row=r, column=3).fill = _fill(C_MB); ws3.cell(row=r, column=3).border = _bdr()
        fc2 = ws3.cell(row=r, column=4, value=fml)
        fc2.font = Font(name="Calibri", size=10, italic=True, color="00005500")
        fc2.fill = _fill(C_FML)
        fc2.border = Border(left=gold_s, right=thin, top=thin, bottom=thin)
        fc2.alignment = _rgt(); fc2.number_format = "#,##0.00"
        ws3.cell(row=r, column=5, value="Auto-calculated").font = Font(name="Calibri", size=8, italic=True, color="00888888")
        ws3.cell(row=r, column=5).fill = _fill(C_FML); ws3.cell(row=r, column=5).border = _bdr()
        r+=1

    # ── SHEET 4: ASSET QUALITY ──────────────────────────────────
    ws4 = wb.create_sheet("📊 Asset_Quality")
    ws4.sheet_properties.tabColor = "DC3545"
    _col_widths(ws4, {"A":3,"B":36,"C":12,"D":18,"E":36,"F":22})
    _sheet_header(ws4, "ASSET QUALITY & SECTOR EXPOSURE")
    r = 3
    _section(ws4, r, "NPA & ASSET QUALITY RATIOS"); r+=1
    for lbl,unit,val,note in [
        ("Gross NPA Ratio","%",7.0,"Gross NPA / Gross Loans × 100"),
        ("Net NPA Ratio","%",3.5,"Net NPA / Net Advances × 100"),
        ("SMA-2 Ratio","%",3.2,"SMA-2 / Gross Loans × 100"),
        ("Restructured Assets Ratio","%",1.8,"Restructured / Gross Loans × 100"),
        ("Provision Coverage Ratio (PCR)","%",50.0,"Provisions / Gross NPA × 100"),
        ("Credit-Deposit Ratio","%",72.9,"Gross Loans / Total Deposits × 100"),
    ]:
        _inp(ws4, r, lbl, unit, val, note); r+=1
    r+=1
    _section(ws4, r, "SECTOR CONCENTRATION (₹ Crore)"); r+=1
    for lbl,unit,val,note in [
        ("Real Estate & Construction","₹ Cr",12000,"Developer loans + project finance"),
        ("Infrastructure","₹ Cr",18000,"Roads, power, ports, telecom"),
        ("NBFC & HFC","₹ Cr",9500,"Loans to NBFCs and housing finance"),
        ("Power Sector","₹ Cr",8800,"Thermal, hydro, renewable energy"),
        ("Textile","₹ Cr",5500,"Spinning, weaving, garments"),
        ("Gems & Jewellery","₹ Cr",3200,"Diamond, gold jewellery"),
        ("Iron & Steel","₹ Cr",0,"Steel, alloys, metals"),
        ("Aviation","₹ Cr",0,"Airlines, airports, MRO"),
        ("Other Sectors","₹ Cr",0,"All remaining sectors"),
    ]:
        _inp(ws4, r, lbl, unit, val, note); r+=1

    # ── SHEET 5: CAPITAL RATIOS ─────────────────────────────────
    ws5 = wb.create_sheet("📊 Capital_Ratios")
    ws5.sheet_properties.tabColor = "28A745"
    _col_widths(ws5, {"A":3,"B":36,"C":12,"D":18,"E":36,"F":22})
    _sheet_header(ws5, "CAPITAL ADEQUACY & LIQUIDITY RATIOS")
    r = 3
    _section(ws5, r, "CAPITAL ADEQUACY"); r+=1
    for lbl,unit,val,note in [
        ("CET1 Ratio","%",9.57,"RBI minimum 8.5% (incl. conservation buffer)"),
        ("Tier 1 Ratio","%",10.64,"RBI minimum 9.5%"),
        ("Total CRAR","%",12.9,"RBI minimum 11.5% (incl. CCB)"),
        ("Leverage Ratio","%",8.1,"RBI minimum 3.5%"),
    ]:
        _inp(ws5, r, lbl, unit, val, note); r+=1
    r+=1
    _section(ws5, r, "LIQUIDITY RATIOS"); r+=1
    for lbl,unit,val,note in [
        ("LCR (Liquidity Coverage Ratio)","%",142.0,"RBI minimum 100%"),
        ("NSFR (Net Stable Funding Ratio)","%",118.0,"RBI minimum 100%"),
        ("CASA Ratio","%",40.0,"CASA / Total Deposits × 100"),
    ]:
        _inp(ws5, r, lbl, unit, val, note); r+=1
    r+=1
    _section(ws5, r, "PROFITABILITY RATIOS"); r+=1
    for lbl,unit,val,note in [
        ("Net Interest Margin (NIM)","%",3.9,"NII / Average earning assets"),
        ("Return on Equity (ROE)","%",13.7,"PAT / Average equity"),
        ("Return on Assets (ROA)","%",1.1,"PAT / Average total assets"),
        ("Cost-to-Income Ratio","%",46.2,"Operating costs / Total income"),
    ]:
        _inp(ws5, r, lbl, unit, val, note); r+=1

    # ── SHEET 6: DURATION RISK ──────────────────────────────────
    ws6 = wb.create_sheet("📊 Duration_Risk")
    ws6.sheet_properties.tabColor = "ADD8E6"
    _col_widths(ws6, {"A":3,"B":36,"C":12,"D":18,"E":36,"F":22})
    _sheet_header(ws6, "DURATION, MARKET RISK & REPRICING GAPS")
    r = 3
    _section(ws6, r, "INVESTMENT PORTFOLIO — DURATION RISK"); r+=1
    for lbl,unit,val,note in [
        ("Avg. Modified Duration — AFS","Years",4.8,"Key driver of MTM loss under rate shock"),
        ("Avg. Modified Duration — HTM","Years",6.2,"No MTM impact — economic risk only"),
        ("Avg. Modified Duration — Liabilities","Years",2.1,"Weighted avg. duration of deposits + borrowings"),
        ("Duration Gap (Assets − Liabilities)","Years",2.7,"Positive = rate rise hurts equity value"),
    ]:
        _inp(ws6, r, lbl, unit, val, note); r+=1
    r+=1
    _section(ws6, r, "INTEREST RATE REPRICING GAPS (₹ Crore)"); r+=1
    for lbl,unit,val,note in [
        ("Rate-Sensitive Assets (RSA)","₹ Cr",95000,"Assets repricing within 1 year"),
        ("Rate-Sensitive Liabilities (RSL)","₹ Cr",76500,"Liabilities repricing within 1 year"),
        ("Net Repricing Gap (RSA − RSL)","₹ Cr",18500,"Positive = asset-sensitive bank"),
        ("Equity Portfolio Beta","β",0.95,"Weighted avg beta vs Nifty 50"),
        ("Net Open FX Position","₹ Cr",800,"Net forex exposure (+ = long USD)"),
    ]:
        _inp(ws6, r, lbl, unit, val, note); r+=1

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ─────────────────────────────────────────────────────────────────
# EXCEL PARSER — reads uploaded file → bank dict
# ─────────────────────────────────────────────────────────────────
def parse_uploaded_excel(file_bytes):
    """Parse uploaded Excel template → bank data dict. Returns (bank_dict, errors)."""
    from openpyxl import load_workbook
    errors = []
    warnings = []

    try:
        wb = load_workbook(filename=io.BytesIO(file_bytes), data_only=True)
    except Exception as e:
        return None, [f"Cannot open file: {e}"]

    required_sheets = ["📊 BS_Assets", "📊 BS_Liabilities", "📊 Income_Stmt",
                       "📊 Asset_Quality", "📊 Capital_Ratios", "📊 Duration_Risk"]
    missing = [s for s in required_sheets if s not in wb.sheetnames]
    if missing:
        return None, [f"Missing sheets: {', '.join(missing)}. Please use the official template."]

    def _val(ws, row, col=4, fallback=0):
        v = ws.cell(row=row, column=col).value
        if v is None: return fallback
        try: return float(v)
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

    bank = {
        # Identity
        "name":   _str(wa, 4) or "Uploaded Bank",
        "type":   _str(wa, 5) or "Scheduled Commercial Bank",
        "rating": _str(wa, 7) or "N/A",
        # ── Assets ──
        "total_assets":       _val(wa, 10),
        "gross_loans":        _val(wa, 11),
        "retail_loans":       _val(wa, 12),
        "corporate_loans":    _val(wa, 13),
        "msme_loans":         _val(wa, 14),
        "agri_loans":         _val(wa, 15),
        "gross_npa":          _val(wa, 18),
        "net_npa":            _val(wa, 19),
        "provisions":         _val(wa, 20),
        "investments_htm":    _val(wa, 23),
        "investments_afs":    _val(wa, 24),
        "equity_portfolio":   _val(wa, 26),
        "cash_hqla":          _val(wa, 29),
        "fixed_assets":       _val(wa, 31),
        "other_assets":       _val(wa, 32),
        # ── Liabilities ──
        "total_deposits":     _val(wl, 4),
        "casa_deposits":      _val(wl, 5),
        "term_deposits":      _val(wl, 6),
        "wholesale_funding":  _val(wl, 9),
        "sub_debt":           _val(wl, 10),
        "equity_capital":     _val(wl, 19),
        "cet1_capital":       _val(wl, 14),
        "tier1_capital":      _val(wl, 16),
        "tier2_capital":      _val(wl, 15),
        "total_capital":      _val(wl, 16),
        "rwa":                _val(wl, 17),
        # ── Income ──
        "net_interest_income": _val(wi, 4),
        "fee_income":          _val(wi, 5),
        "trading_income":      _val(wi, 6),
        "other_income":        _val(wi, 7),
        "operating_costs":     _val(wi, 12),
        "provisions_charge":   _val(wi, 13),
        "pre_provision_profit":_val(wi, 9),
        "pat":                 _val(wi, 21),
        "pbt":                 _val(wi, 19),
        "tax":                 _val(wi, 20),
        # ── Asset quality ──
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
        # ── Capital ratios ──
        "cet1_ratio":          _val(wc, 4),
        "crar":                _val(wc, 6),
        "leverage_ratio":      _val(wc, 7),
        "lcr":                 _val(wc, 11),
        "nsfr":                _val(wc, 12),
        "casa_ratio":          _val(wc, 13),
        "nim":                 _val(wc, 17),
        "roe":                 _val(wc, 18),
        "roa":                 _val(wc, 19),
        # ── Duration ──
        "avg_duration_assets": _val(wd, 4),
        "avg_duration_liabs":  _val(wd, 6),
        "repricing_gap":       _val(wd, 11),
    }

    # ── Derive missing fields ──
    if bank["gross_loans"] > 0:
        bank["gross_npa_ratio"] = bank["gross_npa_ratio"] or round(bank["gross_npa"] / bank["gross_loans"] * 100, 2)
        bank["net_npa_ratio"]   = bank["net_npa_ratio"]   or round(bank["net_npa"]   / bank["gross_loans"] * 100, 2)
        bank["pcr"]             = bank["pcr"]             or round(bank["provisions"]/ bank["gross_npa"]   * 100, 2) if bank["gross_npa"] else 50.0
    if bank["total_deposits"] > 0:
        bank["casa_ratio"] = bank["casa_ratio"] or round(bank["casa_deposits"] / bank["total_deposits"] * 100, 1)
        bank["cd_ratio"]   = bank["cd_ratio"]   or round(bank["gross_loans"]   / bank["total_deposits"] * 100, 1)
    if bank["rwa"] > 0:
        bank["cet1_ratio"] = bank["cet1_ratio"] or round(bank["cet1_capital"] / bank["rwa"] * 100, 2)
        bank["crar"]       = bank["crar"]       or round(bank["total_capital"] / bank["rwa"] * 100, 2)
    if bank["total_assets"] > 0:
        bank["nim"] = bank["nim"] or round(bank["net_interest_income"] / bank["total_assets"] * 100, 2)
        bank["roa"] = bank["roa"] or round(bank["pat"] / bank["total_assets"] * 100, 2)
    if bank["equity_capital"] > 0:
        bank["roe"] = bank["roe"] or round(bank["pat"] / bank["equity_capital"] * 100, 1)
    bank["pre_provision_profit"] = bank["pre_provision_profit"] or (
        bank["net_interest_income"] + bank["fee_income"] +
        bank["trading_income"] + bank["other_income"] - bank["operating_costs"]
    )

    # Validation checks
    critical = {
        "Total Assets": bank["total_assets"],
        "Gross Loans":  bank["gross_loans"],
        "RWA":          bank["rwa"],
        "NII":          bank["net_interest_income"],
        "Total Deposits": bank["total_deposits"],
    }
    for field, val in critical.items():
        if val <= 0:
            errors.append(f"❌ {field} is zero or missing — please fill in the template")

    return bank, errors


COLORS = {
    "darkblue":   "#003366",
    "midblue":    "#004d80",
    "cardBg":     "#112240",
    "bgDark":     "#0a1628",
    "bgGradient": "linear-gradient(135deg,#1a2332,#243447,#2a3f5f)",
    "gold":       "#FFD700",
    "lightblue":  "#ADD8E6",
    "text":       "#e6f1ff",
    "muted":      "#8892b0",
    "green":      "#28a745",
    "red":        "#dc3545",
    "orange":     "#fd7e14",
    "yellow":     "#ffc107",
}

PLOTLY_TEMPLATE = dict(
    paper_bgcolor="#0a1628",
    plot_bgcolor="#112240",
    font=dict(color="#e6f1ff", family="Source Sans Pro"),
    xaxis=dict(gridcolor="#1e3a5f", linecolor="#003366"),
    yaxis=dict(gridcolor="#1e3a5f", linecolor="#003366"),
)

def apply_css():
    st.markdown(f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@400;700;900&family=Source+Sans+Pro:wght@300;400;600&display=swap');

    html, body, [class*="css"] {{
        font-family: 'Source Sans Pro', sans-serif;
        color: {COLORS['text']};
    }}
    .stApp {{
        background: {COLORS['bgGradient']};
        min-height: 100vh;
    }}
    /* Sidebar */
    [data-testid="stSidebar"] {{
        background: linear-gradient(180deg, {COLORS['bgDark']}, {COLORS['darkblue']}) !important;
        border-right: 2px solid {COLORS['gold']};
    }}
    [data-testid="stSidebar"] * {{ color: {COLORS['text']} !important; }}
    [data-testid="stSidebar"] .stSelectbox label,
    [data-testid="stSidebar"] .stSlider label,
    [data-testid="stSidebar"] .stNumberInput label {{ color: {COLORS['gold']} !important; font-weight: 600; }}

    /* Cards */
    .metric-card {{
        background: {COLORS['cardBg']};
        border: 1px solid {COLORS['midblue']};
        border-top: 3px solid {COLORS['gold']};
        border-radius: 8px;
        padding: 16px 20px;
        margin: 6px 0;
    }}
    .metric-label {{
        color: {COLORS['muted']};
        font-size: 0.78rem;
        text-transform: uppercase;
        letter-spacing: 1px;
        font-weight: 600;
    }}
    .metric-value {{
        color: {COLORS['gold']};
        font-size: 1.8rem;
        font-weight: 700;
        font-family: 'Playfair Display', serif;
    }}
    .metric-delta {{
        font-size: 0.82rem;
        font-weight: 600;
    }}
    .delta-pos {{ color: {COLORS['green']}; }}
    .delta-neg {{ color: {COLORS['red']}; }}
    .delta-warn {{ color: {COLORS['orange']}; }}

    /* Section headers */
    .section-header {{
        font-family: 'Playfair Display', serif;
        font-size: 1.4rem;
        font-weight: 700;
        color: {COLORS['gold']};
        border-bottom: 2px solid {COLORS['midblue']};
        padding-bottom: 8px;
        margin: 18px 0 14px 0;
    }}
    .hero-title {{
        font-family: 'Playfair Display', serif;
        font-size: 2.2rem;
        font-weight: 900;
        color: {COLORS['gold']};
        letter-spacing: 1px;
    }}
    .hero-subtitle {{
        color: {COLORS['lightblue']};
        font-size: 1.0rem;
        font-weight: 300;
    }}
    .brand-tag {{
        color: {COLORS['muted']};
        font-size: 0.75rem;
        letter-spacing: 2px;
        text-transform: uppercase;
    }}

    /* Risk badges */
    .badge-green  {{ background:#0d3320; color:{COLORS['green']}; border:1px solid {COLORS['green']}; border-radius:4px; padding:2px 10px; font-size:0.78rem; font-weight:700; }}
    .badge-yellow {{ background:#332900; color:{COLORS['yellow']}; border:1px solid {COLORS['yellow']}; border-radius:4px; padding:2px 10px; font-size:0.78rem; font-weight:700; }}
    .badge-orange {{ background:#332000; color:{COLORS['orange']}; border:1px solid {COLORS['orange']}; border-radius:4px; padding:2px 10px; font-size:0.78rem; font-weight:700; }}
    .badge-red    {{ background:#330d10; color:{COLORS['red']}; border:1px solid {COLORS['red']}; border-radius:4px; padding:2px 10px; font-size:0.78rem; font-weight:700; }}

    /* Tables */
    .styled-table {{ width:100%; border-collapse:collapse; font-size:0.85rem; }}
    .styled-table th {{ background:{COLORS['darkblue']}; color:{COLORS['gold']}; padding:8px 12px; text-align:left; font-weight:700; letter-spacing:0.5px; }}
    .styled-table td {{ padding:7px 12px; border-bottom:1px solid #1e3a5f; color:{COLORS['text']}; }}
    .styled-table tr:hover td {{ background:#243447; }}

    /* Scenario pill */
    .scenario-pill {{
        display:inline-block;
        background:{COLORS['midblue']};
        color:{COLORS['gold']};
        border-radius:20px;
        padding:3px 14px;
        font-size:0.78rem;
        font-weight:700;
        letter-spacing:0.5px;
        margin:2px;
    }}
    /* Divider */
    .gold-divider {{ border:none; border-top:1px solid {COLORS['gold']}44; margin:16px 0; }}

    /* Override Streamlit defaults */
    .stTabs [data-baseweb="tab-list"] {{ background:{COLORS['cardBg']}; border-radius:8px; padding:4px; border:1px solid {COLORS['midblue']}; }}
    .stTabs [data-baseweb="tab"] {{ color:{COLORS['muted']} !important; font-weight:600; font-family:'Source Sans Pro',sans-serif; padding:6px 14px; }}
    .stTabs [data-baseweb="tab"]:hover {{ color:{COLORS['text']} !important; }}
    .stTabs [aria-selected="true"] {{ background:{COLORS['darkblue']} !important; color:{COLORS['gold']} !important; border-radius:6px; font-weight:700 !important; }}
    .stTabs [data-baseweb="tab-panel"] {{ padding:16px 0; background:transparent; }}
    [data-testid="stTabsNavContainer"] {{ margin-bottom:8px; }}

    /* ── Expander contrast fix ── */
    [data-testid="stExpander"] {{
        background:{COLORS['cardBg']} !important;
        border:1px solid {COLORS['midblue']} !important;
        border-radius:8px !important;
    }}
    [data-testid="stExpander"] summary {{
        background:{COLORS['darkblue']} !important;
        border-radius:7px !important;
        padding:8px 14px !important;
    }}
    [data-testid="stExpander"] summary p,
    [data-testid="stExpander"] summary span,
    [data-testid="stExpander"] summary div {{
        color:#ffffff !important;
        font-weight:700 !important;
        font-size:0.9rem !important;
    }}
    [data-testid="stExpander"] summary svg {{
        fill:{COLORS['gold']} !important;
    }}
    [data-testid="stExpander"] summary:hover {{
        background:{COLORS['midblue']} !important;
    }}
    [data-testid="stExpanderDetails"] {{
        background:{COLORS['cardBg']} !important;
        border-top:1px solid {COLORS['midblue']} !important;
    }}
    div[data-testid="stMetricValue"] {{ color:{COLORS['gold']}; font-family:'Playfair Display',serif; }}

    /* ── Selectbox: selected value text ── */
    [data-testid="stSidebar"] [data-baseweb="select"] [data-testid="stMarkdownContainer"] p,
    [data-testid="stSidebar"] [data-baseweb="select"] span,
    [data-testid="stSidebar"] [data-baseweb="select"] div,
    [data-testid="stSidebar"] .stSelectbox div[data-baseweb="select"] > div {{
        color: {COLORS['text']} !important;
        font-weight: 600 !important;
    }}
    /* Selectbox container box */
    [data-testid="stSidebar"] [data-baseweb="select"] > div:first-child {{
        background-color: {COLORS['cardBg']} !important;
        border: 1px solid {COLORS['midblue']} !important;
        border-radius: 6px !important;
    }}
    [data-testid="stSidebar"] [data-baseweb="select"] > div:first-child:hover {{
        border-color: {COLORS['gold']} !important;
    }}
    /* ── Dropdown popup list ── */
    [data-baseweb="popover"] ul,
    [data-baseweb="menu"] ul,
    [role="listbox"] {{
        background-color: {COLORS['cardBg']} !important;
        border: 1px solid {COLORS['midblue']} !important;
    }}
    /* Each option item */
    [data-baseweb="menu"] li,
    [role="option"] {{
        background-color: {COLORS['cardBg']} !important;
        color: {COLORS['text']} !important;
        font-weight: 500 !important;
    }}
    [data-baseweb="menu"] li:hover,
    [role="option"]:hover {{
        background-color: {COLORS['darkblue']} !important;
        color: {COLORS['gold']} !important;
    }}
    /* Selected/highlighted option */
    [aria-selected="true"][role="option"],
    [data-baseweb="menu"] li[aria-selected="true"] {{
        background-color: {COLORS['darkblue']} !important;
        color: {COLORS['gold']} !important;
        font-weight: 700 !important;
    }}
    /* Dropdown chevron icon */
    [data-testid="stSidebar"] [data-baseweb="select"] svg {{
        fill: {COLORS['gold']} !important;
    }}
    </style>
    """, unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────
# SYNTHETIC BANK DATA
# ─────────────────────────────────────────────────────────────────
def generate_bank_data():
    """Generate a realistic synthetic Indian mid-size bank."""
    np.random.seed(42)
    bank = {
        "name": "Mountain Path Bank Ltd.",
        "type": "Private Sector Scheduled Commercial Bank",
        "rating": "AA-",
        # ── Balance Sheet (₹ Crore) ──
        "total_assets": 185000,
        "gross_loans": 108000,
        "retail_loans": 38000,
        "corporate_loans": 45000,
        "msme_loans": 18000,
        "agri_loans": 7000,
        "gross_npa": 7560,       # 7.0% gross NPA
        "net_npa": 3780,
        "provisions": 3780,
        "investments_htm": 28000,
        "investments_afs": 14000,
        "equity_portfolio": 3200,
        "cash_hqla": 22000,
        "fixed_assets": 1800,
        "other_assets": 8000,
        # ── Liabilities ──
        "total_deposits": 148000,
        "casa_deposits": 59200,   # 40% CASA
        "term_deposits": 88800,
        "wholesale_funding": 18500,
        "sub_debt": 3200,
        "equity_capital": 15300,
        # ── Capital ──
        "cet1_capital": 13500,
        "tier1_capital": 15000,
        "tier2_capital": 3200,
        "total_capital": 18200,
        "rwa": 141000,
        # ── Income (Annual, ₹ Crore) ──
        "net_interest_income": 7200,
        "fee_income": 2100,
        "trading_income": 680,
        "other_income": 420,
        "operating_costs": 4800,
        "pre_provision_profit": 5600,
        "provisions_charge": 2800,
        "pbt": 2800,
        "tax": 700,
        "pat": 2100,
        # ── Key Ratios ──
        "nim": 3.9,
        "roe": 13.7,
        "roa": 1.1,
        "casa_ratio": 40.0,
        "pcr": 50.0,
        "lcr": 142.0,
        "nsfr": 118.0,
        "crar": 12.9,
        "cet1_ratio": 9.57,
        "leverage_ratio": 8.1,
        "cd_ratio": 72.9,
        # ── Asset Quality ──
        "gross_npa_ratio": 7.0,
        "net_npa_ratio": 3.5,
        "slma_ratio": 3.2,
        "restructured_ratio": 1.8,
        # ── Sector Exposure ──
        "sector_real_estate": 12000,
        "sector_infrastructure": 18000,
        "sector_textile": 5500,
        "sector_gems": 3200,
        "sector_nbfc": 9500,
        "sector_power": 8800,
        # ── Duration ──
        "avg_duration_assets": 4.8,
        "avg_duration_liabs": 2.1,
        "repricing_gap": 18500,   # rate-sensitive assets - liabilities
    }
    return bank

def generate_macro_baseline():
    """8-quarter baseline macro path."""
    return pd.DataFrame({
        "quarter": [f"Q{i}" for i in range(1, 9)],
        "gdp_growth": [6.8, 7.1, 7.0, 6.9, 7.2, 7.3, 7.1, 7.4],
        "inflation": [4.2, 4.0, 3.9, 4.1, 4.3, 4.2, 4.0, 3.8],
        "repo_rate": [6.50]*8,
        "usd_inr": [83.2, 83.5, 83.8, 84.0, 84.2, 84.5, 84.7, 85.0],
        "nifty_index": [22500, 23100, 23600, 24000, 24500, 25000, 25400, 25900],
        "unemployment": [7.8]*8,
        "property_price_chg": [5.2, 5.5, 5.8, 5.5, 5.2, 5.0, 5.3, 5.6],
        "credit_spread_bps": [180]*8,
    })

# ─────────────────────────────────────────────────────────────────
# SCENARIO LIBRARY
# ─────────────────────────────────────────────────────────────────
SCENARIOS = {
    "🟢 Baseline": {
        "color": COLORS["green"],
        "severity": "Baseline",
        "description": "Business-as-usual with stable macro environment.",
        "gdp_shock": [0.0]*8,
        "rate_shock_bps": 0,
        "equity_shock_pct": 0.0,
        "inr_depreciation": 0.0,
        "npa_multiplier": 1.0,
        "deposit_runoff": 0.0,
        "property_shock": 0.0,
        "credit_spread_widening": 0,
        "fee_income_shock": 0.0,
    },
    "🟡 Mild Stress": {
        "color": COLORS["yellow"],
        "severity": "Mild",
        "description": "Moderate slowdown: RBI rate hike cycle, mild NPA uptick.",
        "gdp_shock": [-0.8, -1.0, -0.9, -0.7, -0.5, -0.4, -0.3, -0.2],
        "rate_shock_bps": 75,
        "equity_shock_pct": -0.12,
        "inr_depreciation": 0.04,
        "npa_multiplier": 1.3,
        "deposit_runoff": 0.02,
        "property_shock": -0.05,
        "credit_spread_widening": 50,
        "fee_income_shock": -0.08,
    },
    "🟠 Moderate Stress": {
        "color": COLORS["orange"],
        "severity": "Moderate",
        "description": "Significant macro downturn: global spillover, corporate NPA surge.",
        "gdp_shock": [-1.5, -2.5, -2.8, -2.2, -1.8, -1.2, -0.8, -0.5],
        "rate_shock_bps": 150,
        "equity_shock_pct": -0.28,
        "inr_depreciation": 0.10,
        "npa_multiplier": 1.8,
        "deposit_runoff": 0.06,
        "property_shock": -0.15,
        "credit_spread_widening": 120,
        "fee_income_shock": -0.18,
    },
    "🔴 Severe Stress": {
        "color": COLORS["red"],
        "severity": "Severe",
        "description": "Deep recession: GDP contraction, crisis-level NPA, equity crash.",
        "gdp_shock": [-2.5, -4.5, -5.2, -4.8, -3.5, -2.2, -1.5, -0.8],
        "rate_shock_bps": 250,
        "equity_shock_pct": -0.45,
        "inr_depreciation": 0.18,
        "npa_multiplier": 2.5,
        "deposit_runoff": 0.12,
        "property_shock": -0.28,
        "credit_spread_widening": 280,
        "fee_income_shock": -0.32,
    },
    "⚫ Extreme / Tail": {
        "color": "#cc66ff",
        "severity": "Extreme",
        "description": "Systemic crisis: GFC-2008 severity applied to Indian banking context.",
        "gdp_shock": [-3.5, -6.0, -7.3, -6.5, -5.0, -3.5, -2.0, -1.0],
        "rate_shock_bps": 350,
        "equity_shock_pct": -0.60,
        "inr_depreciation": 0.28,
        "npa_multiplier": 3.8,
        "deposit_runoff": 0.20,
        "property_shock": -0.40,
        "credit_spread_widening": 500,
        "fee_income_shock": -0.45,
    },
    "🏗️ Real Estate Collapse": {
        "color": "#ff9966",
        "severity": "Sector-specific",
        "description": "Property prices crash 35%: developer NPAs cascade to bank balance sheets.",
        "gdp_shock": [-1.2, -2.0, -2.5, -2.0, -1.5, -1.0, -0.8, -0.5],
        "rate_shock_bps": 50,
        "equity_shock_pct": -0.22,
        "inr_depreciation": 0.06,
        "npa_multiplier": 2.2,
        "deposit_runoff": 0.05,
        "property_shock": -0.35,
        "credit_spread_widening": 150,
        "fee_income_shock": -0.20,
    },
    "📈 Rate Shock": {
        "color": "#66ccff",
        "severity": "Market risk",
        "description": "Sudden 300 bps rate hike: MTM losses on investment portfolio, NII pressure.",
        "gdp_shock": [-0.5, -0.8, -0.7, -0.5, -0.3, -0.2, -0.1, 0.0],
        "rate_shock_bps": 300,
        "equity_shock_pct": -0.18,
        "inr_depreciation": -0.03,
        "npa_multiplier": 1.2,
        "deposit_runoff": 0.04,
        "property_shock": -0.10,
        "credit_spread_widening": 80,
        "fee_income_shock": -0.05,
    },
    "💧 Liquidity Crisis": {
        "color": "#33cccc",
        "severity": "Liquidity",
        "description": "IL&FS-style contagion: wholesale funding freeze, deposit run-off surge.",
        "gdp_shock": [-1.0, -1.5, -1.2, -0.8, -0.5, -0.3, -0.2, -0.1],
        "rate_shock_bps": 100,
        "equity_shock_pct": -0.30,
        "inr_depreciation": 0.08,
        "npa_multiplier": 1.6,
        "deposit_runoff": 0.18,
        "property_shock": -0.12,
        "credit_spread_widening": 220,
        "fee_income_shock": -0.25,
    },
}

# ─────────────────────────────────────────────────────────────────
# STRESS TEST ENGINE
# ─────────────────────────────────────────────────────────────────
def run_stress_test(bank, scenario, custom_params=None):
    """Core stress testing engine — projects 8 quarters."""
    sc = scenario.copy()
    if custom_params:
        sc.update(custom_params)

    quarters = 8
    macro = generate_macro_baseline()

    results = []
    capital = bank["cet1_capital"]
    tier1   = bank["tier1_capital"]
    total_capital = bank["total_capital"]
    rwa     = bank["rwa"]
    gross_npa = bank["gross_npa"]
    gross_loans = bank["gross_loans"]
    deposits = bank["total_deposits"]
    hqla    = bank["cash_hqla"]

    # One-time market shocks (applied at Q1)
    afs_mtm_loss = (bank["avg_duration_assets"] - 2.0) * \
                   (sc["rate_shock_bps"] / 10000) * bank["investments_afs"]
    equity_mtm_loss = abs(sc["equity_shock_pct"]) * bank["equity_portfolio"]
    fx_loss = sc["inr_depreciation"] * (bank["gross_loans"] * 0.04)  # ~4% forex exposure

    for q in range(quarters):
        # ── GDP shock path ──
        gdp_q = macro["gdp_growth"].iloc[q] + sc["gdp_shock"][q]

        # ── Credit Loss Model ──
        gdp_sensitivity = 2.8   # 1% GDP fall → 2.8% relative NPA increase
        gdp_drag = max(-sc["gdp_shock"][q], 0)
        npa_ratio_stressed = min(
            (bank["gross_npa_ratio"] / 100) * sc["npa_multiplier"] *
            (1 + gdp_sensitivity * gdp_drag / 100),
            0.28
        )
        gross_npa_q = gross_loans * npa_ratio_stressed

        # Incremental provisions (60% PCR on fresh NPAs)
        prev_npa = gross_npa if q == 0 else results[-1]["gross_npa"]
        incremental_npa = max(gross_npa_q - prev_npa, 0)
        incremental_provisions = incremental_npa * 0.60

        # ── NII Impact ──
        rate_bps = sc["rate_shock_bps"] if q < 2 else sc["rate_shock_bps"] * 0.7
        nii_rate_impact = bank["repricing_gap"] * (rate_bps / 10000) * 0.25
        nii_quarterly = bank["net_interest_income"] / 4 + nii_rate_impact * (1 if q >= 1 else 0)

        # ── Fee Income ──
        fee_q = bank["fee_income"] / 4 * (1 + sc["fee_income_shock"] * (q/7))

        # ── Trading / MTM (Q1 one-time shock) ──
        mtm_loss_q = (afs_mtm_loss + equity_mtm_loss + fx_loss) if q == 0 else 0
        trading_q = bank["trading_income"] / 4 - mtm_loss_q

        # ── P&L ──
        ppop_q = nii_quarterly + fee_q + trading_q - bank["operating_costs"] / 4
        net_profit_q = ppop_q - incremental_provisions - bank["tax"] / 4

        # ── Capital ──
        capital += net_profit_q
        tier1 += net_profit_q * 0.85
        total_capital += net_profit_q * 0.90
        rwa *= (1 + max(-sc["gdp_shock"][q], 0) * 0.003 + 0.003)  # mild RWA inflation

        cet1_ratio = capital / rwa * 100
        tier1_ratio = tier1 / rwa * 100
        crar = total_capital / rwa * 100

        # ── Liquidity ──
        deposit_runoff_q = sc["deposit_runoff"] * deposits * (1.2 if q == 0 else 0.5)
        deposits = max(deposits - deposit_runoff_q, deposits * 0.6)
        hqla = max(hqla - deposit_runoff_q * 0.3, 0)
        lcr = (hqla / max(deposits * 0.03 + deposit_runoff_q, 1)) * 100

        # ── Property / Collateral ──
        collateral_loss = abs(sc["property_shock"]) * bank["sector_real_estate"] * 0.20 \
                          if q == 0 else 0
        capital -= collateral_loss if q == 0 else 0

        results.append({
            "quarter": f"Q{q+1}",
            "gdp_growth": gdp_q,
            "gross_npa_ratio": npa_ratio_stressed * 100,
            "gross_npa": gross_npa_q,
            "incremental_provisions": incremental_provisions,
            "nii": nii_quarterly,
            "fee_income": fee_q,
            "trading_income": trading_q,
            "ppop": ppop_q,
            "net_profit": net_profit_q,
            "cet1_capital": capital,
            "tier1_capital": tier1,
            "total_capital": total_capital,
            "rwa": rwa,
            "cet1_ratio": cet1_ratio,
            "tier1_ratio": tier1_ratio,
            "crar": crar,
            "lcr": lcr,
            "deposits": deposits,
            "hqla": hqla,
            "mtm_loss": mtm_loss_q,
            "breach_cet1": cet1_ratio < 8.0,
            "breach_tier1": tier1_ratio < 9.5,
            "breach_crar": crar < 11.5,
            "breach_lcr": lcr < 100,
        })

    df = pd.DataFrame(results)
    # Summary stats
    summary = {
        "min_cet1": df["cet1_ratio"].min(),
        "min_crar": df["crar"].min(),
        "min_lcr": df["lcr"].min(),
        "peak_npa": df["gross_npa_ratio"].max(),
        "cumulative_provisions": df["incremental_provisions"].sum(),
        "cumulative_net_profit": df["net_profit"].sum(),
        "total_mtm_loss": df["mtm_loss"].sum(),
        "any_breach": df[["breach_cet1","breach_tier1","breach_crar","breach_lcr"]].any().any(),
        "cet1_breaches": df["breach_cet1"].sum(),
    }
    return df, summary

# ─────────────────────────────────────────────────────────────────
# SENSITIVITY ANALYSIS
# ─────────────────────────────────────────────────────────────────
def run_sensitivity(bank, variable, values, base_scenario):
    """Sweep a single variable and record min CET1."""
    results = []
    for val in values:
        sc = base_scenario.copy()
        sc[variable] = val
        if variable == "rate_shock_bps":
            pass
        elif variable == "npa_multiplier":
            sc["npa_multiplier"] = val
        elif variable == "equity_shock_pct":
            sc["equity_shock_pct"] = -abs(val)
        _, summary = run_stress_test(bank, sc)
        results.append({"value": val, "min_cet1": summary["min_cet1"],
                         "peak_npa": summary["peak_npa"]})
    return pd.DataFrame(results)

# ─────────────────────────────────────────────────────────────────
# PLOTTING HELPERS
# ─────────────────────────────────────────────────────────────────
def base_fig(**kwargs):
    fig = go.Figure(**kwargs)
    fig.update_layout(
        paper_bgcolor=PLOTLY_TEMPLATE["paper_bgcolor"],
        plot_bgcolor=PLOTLY_TEMPLATE["plot_bgcolor"],
        font=PLOTLY_TEMPLATE["font"],
        margin=dict(l=40, r=20, t=50, b=40),
        legend=dict(bgcolor="#0a1628", bordercolor="#003366", borderwidth=1,
                    font=dict(color=COLORS["text"])),
    )
    fig.update_xaxes(gridcolor="#1e3a5f", linecolor="#003366", tickfont=dict(color=COLORS["muted"]))
    fig.update_yaxes(gridcolor="#1e3a5f", linecolor="#003366", tickfont=dict(color=COLORS["muted"]))
    return fig

def plot_multi_scenario_cet1(bank, selected_scenarios):
    fig = base_fig()
    for s_name in selected_scenarios:
        sc = SCENARIOS[s_name]
        df, _ = run_stress_test(bank, sc)
        fig.add_trace(go.Scatter(
            x=df["quarter"], y=df["cet1_ratio"],
            name=s_name.split(" ", 1)[1],
            line=dict(color=sc["color"], width=2.5),
            mode="lines+markers",
            marker=dict(size=6),
        ))
    # Regulatory floors
    fig.add_hline(y=8.0, line=dict(color="#dc3545", dash="dash", width=1.5),
                  annotation_text="CET1 Min (8%)", annotation_font_color="#dc3545")
    fig.add_hline(y=9.5, line=dict(color="#fd7e14", dash="dot", width=1.2),
                  annotation_text="RBI Buffer (9.5%)", annotation_font_color="#fd7e14")
    fig.update_layout(title="CET1 Ratio Under Stress Scenarios (%)",
                      title_font=dict(color=COLORS["gold"], size=15, family="Playfair Display"),
                      yaxis_title="CET1 Ratio (%)", xaxis_title="Quarter",
                      height=380)
    return fig

def plot_npa_evolution(bank, selected_scenarios):
    fig = base_fig()
    for s_name in selected_scenarios:
        sc = SCENARIOS[s_name]
        df, _ = run_stress_test(bank, sc)
        fig.add_trace(go.Scatter(
            x=df["quarter"], y=df["gross_npa_ratio"],
            name=s_name.split(" ", 1)[1],
            line=dict(color=sc["color"], width=2.5),
            mode="lines+markers",
            fill="tozeroy" if s_name == list(selected_scenarios)[0] else "none",
            fillcolor=hex_to_rgba(sc["color"], 0.08),
        ))
    fig.update_layout(title="Gross NPA Ratio Evolution (%)",
                      title_font=dict(color=COLORS["gold"], size=15, family="Playfair Display"),
                      yaxis_title="Gross NPA (%)", xaxis_title="Quarter", height=380)
    return fig

def plot_capital_waterfall(bank, scenario_name):
    sc = SCENARIOS[scenario_name]
    df, summary = run_stress_test(bank, sc)

    items = ["Starting CET1", "Credit Losses", "MTM / Market", "NII Impact",
             "Fee & Other", "Ending CET1"]
    credit_loss = -summary["cumulative_provisions"] / 100
    mtm_loss    = -summary["total_mtm_loss"] / 100
    nii_gain    = (df["nii"].sum() - bank["net_interest_income"] * 2) / 100
    fee_loss    = (df["fee_income"].sum() - bank["fee_income"] * 2) / 100
    start       = bank["cet1_capital"] / 100
    end         = summary["min_cet1"] * bank["rwa"] / 10000

    values = [start, credit_loss, mtm_loss, max(nii_gain, -1), max(fee_loss, -1), end]
    measures = ["absolute", "relative", "relative", "relative", "relative", "total"]
    colors = [COLORS["gold"], COLORS["red"], COLORS["orange"],
              COLORS["green"] if nii_gain >= 0 else COLORS["red"],
              COLORS["green"] if fee_loss >= 0 else COLORS["red"],
              COLORS["lightblue"]]

    fig = go.Figure(go.Waterfall(
        name="Capital Walk",
        orientation="v",
        measure=measures,
        x=items,
        y=values,
        connector=dict(line=dict(color=COLORS["midblue"], width=1.5)),
        increasing=dict(marker=dict(color=COLORS["green"])),
        decreasing=dict(marker=dict(color=COLORS["red"])),
        totals=dict(marker=dict(color=COLORS["gold"])),
        text=[f"₹{v*100:.0f}Cr" for v in values],
        textposition="outside",
        textfont=dict(color=COLORS["text"], size=11),
    ))
    fig.update_layout(
        paper_bgcolor=PLOTLY_TEMPLATE["paper_bgcolor"],
        plot_bgcolor=PLOTLY_TEMPLATE["plot_bgcolor"],
        font=PLOTLY_TEMPLATE["font"],
        title=f"Capital Walk — {scenario_name.split(' ',1)[1]}",
        title_font=dict(color=COLORS["gold"], size=15, family="Playfair Display"),
        yaxis_title="₹ '00 Crore",
        height=380,
        margin=dict(l=40, r=20, t=50, b=40),
        showlegend=False,
    )
    fig.update_xaxes(tickfont=dict(color=COLORS["muted"]))
    fig.update_yaxes(gridcolor="#1e3a5f", tickfont=dict(color=COLORS["muted"]))
    return fig

def plot_sensitivity_heatmap(bank):
    """Create sensitivity heatmap: NPA multiplier vs. Rate shock."""
    npa_vals = [1.0, 1.3, 1.6, 2.0, 2.5, 3.0, 3.5]
    rate_vals = [0, 50, 100, 150, 200, 250, 300]
    base_sc = SCENARIOS["🔴 Severe Stress"].copy()
    z = []
    for nm in npa_vals:
        row = []
        for rb in rate_vals:
            sc = base_sc.copy()
            sc["npa_multiplier"] = nm
            sc["rate_shock_bps"] = rb
            _, s = run_stress_test(bank, sc)
            row.append(round(s["min_cet1"], 2))
        z.append(row)

    fig = go.Figure(go.Heatmap(
        z=z, x=[f"{r}bps" for r in rate_vals], y=[f"{n}x" for n in npa_vals],
        colorscale=[[0, "#dc3545"], [0.4, "#fd7e14"], [0.6, "#ffc107"],
                    [0.8, "#28a745"], [1.0, "#00cc66"]],
        zmin=6, zmax=12,
        text=[[f"{v:.1f}%" for v in row] for row in z],
        texttemplate="%{text}",
        textfont=dict(size=10, color="white"),
        colorbar=dict(title=dict(text="Min CET1 %", font=dict(color=COLORS["gold"])),
                      tickfont=dict(color=COLORS["text"]))
    ))
    fig.update_layout(
        paper_bgcolor=PLOTLY_TEMPLATE["paper_bgcolor"],
        plot_bgcolor=PLOTLY_TEMPLATE["plot_bgcolor"],
        font=PLOTLY_TEMPLATE["font"],
        title="Sensitivity Heatmap: NPA Multiplier × Rate Shock → Min CET1 (%)",
        title_font=dict(color=COLORS["gold"], size=14, family="Playfair Display"),
        xaxis=dict(title="Rate Shock (bps)", tickfont=dict(color=COLORS["muted"]),
                   title_font=dict(color=COLORS["muted"])),
        yaxis=dict(title="NPA Multiplier", tickfont=dict(color=COLORS["muted"]),
                   title_font=dict(color=COLORS["muted"])),
        height=360,
        margin=dict(l=60, r=40, t=50, b=50),
    )
    return fig

def plot_liquidity_stress(bank, selected_scenarios):
    fig = base_fig()
    for s_name in selected_scenarios:
        sc = SCENARIOS[s_name]
        df, _ = run_stress_test(bank, sc)
        fig.add_trace(go.Scatter(
            x=df["quarter"], y=df["lcr"],
            name=s_name.split(" ", 1)[1],
            line=dict(color=sc["color"], width=2.5),
            mode="lines+markers",
        ))
    fig.add_hline(y=100, line=dict(color="#dc3545", dash="dash", width=1.5),
                  annotation_text="LCR Min (100%)", annotation_font_color="#dc3545")
    fig.update_layout(title="Liquidity Coverage Ratio (LCR) Under Stress (%)",
                      title_font=dict(color=COLORS["gold"], size=15, family="Playfair Display"),
                      yaxis_title="LCR (%)", xaxis_title="Quarter", height=380)
    return fig

def plot_pl_decomposition(bank, scenario_name):
    sc = SCENARIOS[scenario_name]
    df, _ = run_stress_test(bank, sc)
    fig = base_fig()
    fig.add_trace(go.Bar(x=df["quarter"], y=df["nii"],
                         name="NII", marker_color=COLORS["midblue"]))
    fig.add_trace(go.Bar(x=df["quarter"], y=df["fee_income"],
                         name="Fee Income", marker_color=COLORS["lightblue"]))
    fig.add_trace(go.Bar(x=df["quarter"], y=df["trading_income"],
                         name="Trading", marker_color=hex_to_rgba(COLORS["gold"], 0.53)))
    fig.add_trace(go.Bar(x=df["quarter"],
                         y=[-bank["operating_costs"]/4]*8,
                         name="Operating Costs", marker_color="#cc3333"))
    fig.add_trace(go.Bar(x=df["quarter"], y=-df["incremental_provisions"],
                         name="Provisions", marker_color="#ff6666"))
    fig.add_trace(go.Scatter(x=df["quarter"], y=df["net_profit"],
                             name="Net Profit", mode="lines+markers",
                             line=dict(color=COLORS["gold"], width=3),
                             marker=dict(size=8)))
    fig.update_layout(
        barmode="relative",
        title=f"P&L Decomposition — {scenario_name.split(' ',1)[1]}",
        title_font=dict(color=COLORS["gold"], size=15, family="Playfair Display"),
        yaxis_title="₹ Crore", xaxis_title="Quarter", height=400,
    )
    return fig

def plot_radar_scenarios(bank):
    cats = ["CET1 Ratio", "Peak NPA", "LCR", "NII Stability", "Net Profit"]
    fig = go.Figure()
    for s_name, sc in list(SCENARIOS.items())[1:5]:
        df, summary = run_stress_test(bank, sc)
        nii_stability = max(0, (df["nii"].min() / (bank["net_interest_income"]/4)) * 10)
        np_score = max(0, (df["net_profit"].sum() / bank["pat"] + 2) * 2.5)
        values = [
            min(summary["min_cet1"] - 5, 10),
            max(0, 15 - summary["peak_npa"]),
            min(summary["min_lcr"] / 20, 10),
            min(nii_stability, 10),
            min(np_score, 10),
        ]
        values += [values[0]]
        cats_closed = cats + [cats[0]]
        fig.add_trace(go.Scatterpolar(
            r=values, theta=cats_closed,
            name=s_name.split(" ", 1)[1],
            line=dict(color=sc["color"], width=2),
            fill="toself", fillcolor=hex_to_rgba(sc["color"], 0.13),
        ))
    fig.update_layout(
        paper_bgcolor=PLOTLY_TEMPLATE["paper_bgcolor"],
        polar=dict(
            bgcolor="#112240",
            radialaxis=dict(visible=True, range=[0, 10],
                            gridcolor="#1e3a5f", tickfont=dict(color=COLORS["muted"])),
            angularaxis=dict(gridcolor="#1e3a5f",
                             tickfont=dict(color=COLORS["text"]))
        ),
        title="Bank Resilience Radar (Higher = More Resilient)",
        title_font=dict(color=COLORS["gold"], size=14, family="Playfair Display"),
        legend=dict(bgcolor="#0a1628", bordercolor="#003366",
                    font=dict(color=COLORS["text"])),
        height=400,
        font=dict(color=COLORS["text"]),
        margin=dict(l=40, r=40, t=50, b=40),
    )
    return fig

# ─────────────────────────────────────────────────────────────────
# UI HELPERS
# ─────────────────────────────────────────────────────────────────
def metric_card(label, value, delta=None, delta_type="pos"):
    delta_html = ""
    if delta is not None:
        cls = f"delta-{delta_type}"
        arrow = "▲" if delta_type == "pos" else ("▼" if delta_type == "neg" else "●")
        delta_html = f'<div class="metric-delta {cls}">{arrow} {delta}</div>'
    return f"""
    <div class="metric-card">
        <div class="metric-label">{label}</div>
        <div class="metric-value">{value}</div>
        {delta_html}
    </div>"""

def risk_badge(level):
    levels = {"Low":"green","Moderate":"yellow","High":"orange","Critical":"red","Pass":"green","Breach":"red"}
    cls = levels.get(level, "yellow")
    return f'<span class="badge-{cls}">{level}</span>'

def section_header(title, icon=""):
    return f'<div class="section-header">{icon} {title}</div>'

def scenario_summary_table(bank, selected_scenarios):
    rows = []
    for s_name in selected_scenarios:
        sc = SCENARIOS[s_name]
        df, summary = run_stress_test(bank, sc)
        breach = "🔴 YES" if summary["any_breach"] else "🟢 No"
        rows.append({
            "Scenario": s_name,
            "Severity": sc["severity"],
            "Min CET1 (%)": f"{summary['min_cet1']:.2f}%",
            "Peak NPA (%)": f"{summary['peak_npa']:.1f}%",
            "Min LCR (%)": f"{summary['min_lcr']:.0f}%",
            "Cum. Provisions (₹Cr)": f"₹{summary['cumulative_provisions']:,.0f}",
            "Cum. Net Profit (₹Cr)": f"₹{summary['cumulative_net_profit']:,.0f}",
            "Regulatory Breach": breach,
        })
    df_summary = pd.DataFrame(rows)
    header_row = "".join(f"<th>{c}</th>" for c in df_summary.columns)
    data_rows = ""
    for _, row in df_summary.iterrows():
        data_rows += "<tr>" + "".join(f"<td>{v}</td>" for v in row) + "</tr>"
    return f"""
    <table class="styled-table">
        <thead><tr>{header_row}</tr></thead>
        <tbody>{data_rows}</tbody>
    </table>"""

# ─────────────────────────────────────────────────────────────────
# MAIN APP
# ─────────────────────────────────────────────────────────────────
def main():
    apply_css()

    # ── Session state: bank data source ──────────────────────────
    if "bank_data" not in st.session_state:
        st.session_state.bank_data   = generate_bank_data()
        st.session_state.bank_source = "demo"
        st.session_state.bank_label  = "Mountain Path Bank Ltd. (Demo)"
    bank = st.session_state.bank_data

    # ── SIDEBAR ──────────────────────────────────────────────────
    with st.sidebar:
        st.markdown(f"""
        <div style='text-align:center; padding:10px 0 5px 0;'>
            <div style='font-family:Playfair Display,serif;font-size:1.15rem;
                        font-weight:900;color:{COLORS["gold"]};letter-spacing:1px;'>
                🏦 THE MOUNTAIN PATH
            </div>
            <div style='color:{COLORS["lightblue"]};font-size:0.72rem;
                        letter-spacing:2px;text-transform:uppercase;margin-top:2px;'>
                Bank Stress Testing Lab
            </div>
            <hr style='border-color:{COLORS["gold"]}44;margin:10px 0;'>
        </div>
        """, unsafe_allow_html=True)

        st.markdown(f"<div style='color:{COLORS['gold']};font-weight:700;font-size:0.85rem;'>📋 NAVIGATION</div>", unsafe_allow_html=True)
        page = st.selectbox("", [
            "🏠 Dashboard",
            "📊 Balance Sheet & Income",
            "🎯 Run Stress Tests",
            "📈 Capital Analysis",
            "💧 Liquidity Analysis",
            "🔥 Sensitivity & Heatmap",
            "🛠️ Custom Scenario Builder",
            "📋 Scenario Comparison",
            "🔄 Reverse Stress Test",
            "📤 Upload Bank Data",
            "🏛 About the Platform",
            "🎓 Education Hub",
        ], label_visibility="collapsed")

        # Data source indicator
        src_color = COLORS["gold"] if st.session_state.get("bank_source") == "demo" else COLORS["green"]
        src_icon  = "🔵" if st.session_state.get("bank_source") == "demo" else "🟢"
        st.markdown(
            "<div style='background:#0a1628;border:1px solid " + src_color + ";"
            "border-radius:6px;padding:7px 10px;margin:8px 0;font-size:0.78rem;'>"
            "<span style='color:" + src_color + ";font-weight:700;'>" + src_icon + " DATA SOURCE</span><br>"
            "<span style='color:#d0dff0;font-size:0.75rem;'>" +
            st.session_state.get("bank_label", "Demo Bank") +
            "</span></div>",
            unsafe_allow_html=True
        )

        st.markdown("<hr style='border-color:#1e3a5f;'>", unsafe_allow_html=True)
        st.markdown(f"<div style='color:{COLORS['gold']};font-weight:700;font-size:0.85rem;'>📌 SELECT SCENARIOS</div>", unsafe_allow_html=True)
        selected_scenarios = []
        for s_name, sc in SCENARIOS.items():
            checked = st.checkbox(s_name, value=(s_name in ["🟢 Baseline", "🟠 Moderate Stress", "🔴 Severe Stress"]),
                                  key=f"cb_{s_name}")
            if checked:
                selected_scenarios.append(s_name)
        if not selected_scenarios:
            selected_scenarios = ["🟢 Baseline"]

        st.markdown("<hr style='border-color:#1e3a5f;'>", unsafe_allow_html=True)
        st.markdown(f"<div style='color:{COLORS['gold']};font-weight:700;font-size:0.85rem;'>⚙️ BASELINE PARAMETERS</div>", unsafe_allow_html=True)
        bank["gross_npa_ratio"] = st.slider("Starting Gross NPA (%)", 3.0, 18.0, 7.0, 0.5)
        bank["cet1_ratio"]      = st.slider("Starting CET1 Ratio (%)", 8.0, 16.0, 9.57, 0.1)
        bank["lcr"]             = st.slider("Starting LCR (%)", 100.0, 200.0, 142.0, 5.0)
        bank["nim"]             = st.slider("Starting NIM (%)", 2.0, 5.5, 3.9, 0.1)

        st.markdown("<hr style='border-color:#1e3a5f;'>", unsafe_allow_html=True)
        st.markdown(f"""
        <div style='color:{COLORS['muted']};font-size:0.72rem;text-align:center;'>
            Prof. V. Ravichandran<br>
            <a href='https://www.linkedin.com/in/trichyravis' target='_blank'
               style='color:{COLORS["gold"]};'>LinkedIn</a> &nbsp;|&nbsp;
            <a href='https://github.com/trichyravis' target='_blank'
               style='color:{COLORS["gold"]};'>GitHub</a>
        </div>""", unsafe_allow_html=True)

    # ── HEADER ───────────────────────────────────────────────────
    st.markdown(f"""
    <div style='padding:10px 0 18px 0;'>
        <div class='brand-tag'>The Mountain Path — World of Finance</div>
        <div class='hero-title'>🏦 Bank Stress Testing Lab</div>
        <div class='hero-subtitle'>{bank["name"]} &nbsp;|&nbsp; {bank["type"]}
            &nbsp;|&nbsp; Rating: <span style='color:{COLORS["gold"]}'>{bank["rating"]}</span>
        </div>
    </div>
    <hr class='gold-divider'>
    """, unsafe_allow_html=True)

    # ═══════════════════════════════════════════════════════════════
    # PAGE 1: DASHBOARD
    # ═══════════════════════════════════════════════════════════════
    if page == "🏠 Dashboard":
        st.markdown(section_header("Key Risk Indicators — Current Position", "📊"), unsafe_allow_html=True)

        cols = st.columns(5)
        metrics = [
            ("CET1 Ratio", f"{bank['cet1_ratio']:.1f}%", "Above 8% min", "pos"),
            ("Gross NPA", f"{bank['gross_npa_ratio']:.1f}%", "Moderate risk", "warn"),
            ("LCR", f"{bank['lcr']:.0f}%", "Above 100% min", "pos"),
            ("CRAR", f"{bank['crar']:.1f}%", "Above 11.5% min", "pos"),
            ("NIM", f"{bank['nim']:.1f}%", "Healthy spread", "pos"),
        ]
        for col, (lbl, val, delta, dtype) in zip(cols, metrics):
            with col:
                st.markdown(metric_card(lbl, val, delta, dtype), unsafe_allow_html=True)

        cols2 = st.columns(5)
        metrics2 = [
            ("Total Assets", f"₹{bank['total_assets']:,}Cr", None, "pos"),
            ("Gross Loans", f"₹{bank['gross_loans']:,}Cr", "72.9% C/D ratio", "pos"),
            ("CASA Ratio", f"{bank['casa_ratio']:.0f}%", "Good funding mix", "pos"),
            ("Net Interest Income", f"₹{bank['net_interest_income']:,}Cr", "Annual", "pos"),
            ("Return on Equity", f"{bank['roe']:.1f}%", "FY Annualised", "pos"),
        ]
        for col, (lbl, val, delta, dtype) in zip(cols2, metrics2):
            with col:
                st.markdown(metric_card(lbl, val, delta, dtype), unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            st.markdown(section_header("CET1 Under All Scenarios", "💰"), unsafe_allow_html=True)
            st.plotly_chart(plot_multi_scenario_cet1(bank, selected_scenarios),
                            use_container_width=True)
        with c2:
            st.markdown(section_header("NPA Evolution", "📉"), unsafe_allow_html=True)
            st.plotly_chart(plot_npa_evolution(bank, selected_scenarios),
                            use_container_width=True)

        st.markdown(section_header("Quick Scenario Summary", "🎯"), unsafe_allow_html=True)
        st.markdown(scenario_summary_table(bank, selected_scenarios), unsafe_allow_html=True)

        # Radar
        st.markdown("<br>", unsafe_allow_html=True)
        c1, c2 = st.columns([1.2, 0.8])
        with c1:
            st.markdown(section_header("Resilience Radar", "🕸️"), unsafe_allow_html=True)
            st.plotly_chart(plot_radar_scenarios(bank), use_container_width=True)
        with c2:
            st.markdown(section_header("Sector Concentration Risk", "🏗️"), unsafe_allow_html=True)
            sectors = {
                "Real Estate": bank["sector_real_estate"],
                "Infrastructure": bank["sector_infrastructure"],
                "NBFC": bank["sector_nbfc"],
                "Power": bank["sector_power"],
                "Textile": bank["sector_textile"],
                "Gems/Jewellery": bank["sector_gems"],
            }
            fig_pie = go.Figure(go.Pie(
                labels=list(sectors.keys()),
                values=list(sectors.values()),
                hole=0.5,
                marker=dict(colors=[COLORS["darkblue"], COLORS["midblue"],
                                    COLORS["gold"], COLORS["lightblue"],
                                    COLORS["orange"], COLORS["green"]]),
                textfont=dict(color="white", size=11),
            ))
            fig_pie.update_layout(
                paper_bgcolor="#0a1628",
                font=dict(color=COLORS["text"]),
                legend=dict(bgcolor="#0a1628", font=dict(color=COLORS["text"])),
                margin=dict(l=10, r=10, t=30, b=10),
                height=380,
            )
            st.plotly_chart(fig_pie, use_container_width=True)

    # ═══════════════════════════════════════════════════════════════
    # PAGE 2: BALANCE SHEET & INCOME
    # ═══════════════════════════════════════════════════════════════
    elif page == "📊 Balance Sheet & Income":
        c1, c2 = st.columns(2)
        with c1:
            st.markdown(section_header("Balance Sheet — Assets (₹ Crore)", "📋"), unsafe_allow_html=True)
            bs_assets = pd.DataFrame({
                "Item": ["Gross Loans & Advances", "  ► Retail", "  ► Corporate",
                          "  ► MSME", "  ► Agriculture",
                          "HTM Investments", "AFS Investments", "Equity Portfolio",
                          "Cash & HQLA", "Fixed Assets", "Other Assets"],
                "Amount (₹Cr)": [
                    bank["gross_loans"], bank["retail_loans"], bank["corporate_loans"],
                    bank["msme_loans"], bank["agri_loans"],
                    bank["investments_htm"], bank["investments_afs"],
                    bank["equity_portfolio"], bank["cash_hqla"],
                    bank["fixed_assets"], bank["other_assets"]
                ],
                "% of Assets": [
                    f"{v/bank['total_assets']*100:.1f}%" for v in [
                        bank["gross_loans"], bank["retail_loans"], bank["corporate_loans"],
                        bank["msme_loans"], bank["agri_loans"],
                        bank["investments_htm"], bank["investments_afs"],
                        bank["equity_portfolio"], bank["cash_hqla"],
                        bank["fixed_assets"], bank["other_assets"]
                    ]
                ]
            })
            st.dataframe(bs_assets, use_container_width=True, height=380, hide_index=True)

        with c2:
            st.markdown(section_header("Balance Sheet — Liabilities & Capital (₹ Crore)", "📋"), unsafe_allow_html=True)
            bs_liab = pd.DataFrame({
                "Item": ["Total Deposits", "  ► CASA Deposits", "  ► Term Deposits",
                          "Wholesale Funding", "Subordinated Debt",
                          "CET1 Capital", "Additional Tier 1", "Tier 2 Capital",
                          "Total Capital"],
                "Amount (₹Cr)": [
                    bank["total_deposits"], bank["casa_deposits"], bank["term_deposits"],
                    bank["wholesale_funding"], bank["sub_debt"],
                    bank["cet1_capital"], bank["tier1_capital"] - bank["cet1_capital"],
                    bank["tier2_capital"], bank["total_capital"]
                ],
            })
            st.dataframe(bs_liab, use_container_width=True, height=380, hide_index=True)

        c3, c4 = st.columns(2)
        with c3:
            st.markdown(section_header("Income Statement — Annual (₹ Crore)", "💼"), unsafe_allow_html=True)
            is_data = pd.DataFrame({
                "Line Item": ["Net Interest Income", "Fee & Commission Income",
                               "Trading & MTM Income", "Other Income",
                               "Total Operating Income",
                               "Operating Costs", "Pre-Provision Operating Profit",
                               "Provisions & Write-offs", "Profit Before Tax",
                               "Taxes", "Profit After Tax"],
                "Amount (₹Cr)": [
                    bank["net_interest_income"], bank["fee_income"],
                    bank["trading_income"], bank["other_income"],
                    bank["net_interest_income"] + bank["fee_income"] + bank["trading_income"] + bank["other_income"],
                    -bank["operating_costs"], bank["pre_provision_profit"],
                    -bank["provisions_charge"], bank["pbt"], -bank["tax"], bank["pat"]
                ],
            })
            st.dataframe(is_data, use_container_width=True, height=380, hide_index=True)

        with c4:
            st.markdown(section_header("Capital & Asset Quality Ratios", "📐"), unsafe_allow_html=True)
            ratios = pd.DataFrame({
                "Ratio": ["CET1 Ratio", "Tier 1 Ratio", "CRAR (Total)", "Leverage Ratio",
                           "Gross NPA Ratio", "Net NPA Ratio", "SMA Ratio",
                           "Restructured Assets", "PCR",
                           "NIM", "ROE", "ROA", "CASA Ratio", "C/D Ratio",
                           "LCR", "NSFR"],
                "Value": ["9.57%","10.64%","12.90%","8.10%",
                          "7.00%","3.50%","3.20%","1.80%","50%",
                          "3.90%","13.70%","1.10%","40.0%","72.9%",
                          "142%","118%"],
                "Regulatory Min": ["8.0%","9.5%","11.5%","—",
                                    "—","—","—","—","—",
                                    "—","—","—","—","—",
                                    "100%","100%"],
                "Status": ["✅ Pass","✅ Pass","✅ Pass","✅ Pass",
                            "⚠️ Watch","⚠️ Watch","⚠️ Watch","⚠️ Watch","✅ Pass",
                            "✅ Pass","✅ Pass","✅ Pass","✅ Pass","✅ Pass",
                            "✅ Pass","✅ Pass"],
            })
            st.dataframe(ratios, use_container_width=True, height=450, hide_index=True)

    # ═══════════════════════════════════════════════════════════════
    # PAGE 3: RUN STRESS TESTS
    # ═══════════════════════════════════════════════════════════════
    elif page == "🎯 Run Stress Tests":
        sc_choice = st.selectbox("Select Scenario to Analyse in Detail", list(SCENARIOS.keys()))
        sc = SCENARIOS[sc_choice]
        df, summary = run_stress_test(bank, sc)

        st.markdown(f"""
        <div class='metric-card' style='margin-bottom:16px;'>
            <b style='color:{COLORS["gold"]};font-size:1.1rem;'>{sc_choice}</b><br>
            <span style='color:{COLORS["muted"]};'>{sc["description"]}</span>&nbsp;&nbsp;
            <span class='scenario-pill'>Severity: {sc["severity"]}</span>
        </div>""", unsafe_allow_html=True)

        # Shock parameters
        st.markdown(section_header("Scenario Shock Parameters", "⚡"), unsafe_allow_html=True)
        cols = st.columns(4)
        params = [
            ("Rate Shock", f"{sc['rate_shock_bps']} bps"),
            ("Equity Shock", f"{sc['equity_shock_pct']*100:.0f}%"),
            ("NPA Multiplier", f"{sc['npa_multiplier']}×"),
            ("INR Depreciation", f"{sc['inr_depreciation']*100:.0f}%"),
            ("Property Shock", f"{sc['property_shock']*100:.0f}%"),
            ("Deposit Run-off", f"{sc['deposit_runoff']*100:.0f}%"),
            ("Fee Income Shock", f"{sc['fee_income_shock']*100:.0f}%"),
            ("Credit Spread Widening", f"{sc['credit_spread_widening']} bps"),
        ]
        for i, (lbl, val) in enumerate(params):
            with cols[i % 4]:
                st.markdown(metric_card(lbl, val), unsafe_allow_html=True)

        # Results summary
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(section_header("Stress Test Results", "📊"), unsafe_allow_html=True)
        r_cols = st.columns(4)
        breach_status = "🔴 BREACH" if summary["any_breach"] else "🟢 PASS"
        with r_cols[0]:
            st.markdown(metric_card("Min CET1 Ratio", f"{summary['min_cet1']:.2f}%",
                "Below 8% min!" if summary["min_cet1"] < 8 else "Above regulatory min",
                "neg" if summary["min_cet1"] < 8 else "pos"), unsafe_allow_html=True)
        with r_cols[1]:
            st.markdown(metric_card("Peak Gross NPA", f"{summary['peak_npa']:.1f}%",
                "Stress peak", "neg"), unsafe_allow_html=True)
        with r_cols[2]:
            st.markdown(metric_card("Min LCR", f"{summary['min_lcr']:.0f}%",
                "Below 100!" if summary["min_lcr"] < 100 else "Above min",
                "neg" if summary["min_lcr"] < 100 else "pos"), unsafe_allow_html=True)
        with r_cols[3]:
            st.markdown(metric_card("Regulatory Status", breach_status,
                f"{summary['cet1_breaches']} CET1 breach qtr(s)",
                "neg" if summary["any_breach"] else "pos"), unsafe_allow_html=True)

        # Charts row
        c1, c2 = st.columns(2)
        with c1:
            st.plotly_chart(plot_pl_decomposition(bank, sc_choice), use_container_width=True)
        with c2:
            st.plotly_chart(plot_capital_waterfall(bank, sc_choice), use_container_width=True)

        # Quarterly data table
        st.markdown(section_header("Quarterly Projection Detail", "📅"), unsafe_allow_html=True)
        display_df = df[["quarter","gdp_growth","gross_npa_ratio","incremental_provisions",
                          "nii","net_profit","cet1_ratio","crar","lcr"]].copy()
        display_df.columns = ["Quarter","GDP Growth (%)","Gross NPA (%)","Provisions (₹Cr)",
                                "NII (₹Cr)","Net Profit (₹Cr)","CET1 (%)","CRAR (%)","LCR (%)"]
        display_df = display_df.round(2)
        st.dataframe(display_df, use_container_width=True, hide_index=True)

    # ═══════════════════════════════════════════════════════════════
    # PAGE 4: CAPITAL ANALYSIS
    # ═══════════════════════════════════════════════════════════════
    elif page == "📈 Capital Analysis":
        st.markdown(section_header("Capital Ratio Paths Under All Selected Scenarios", "💰"), unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            st.plotly_chart(plot_multi_scenario_cet1(bank, selected_scenarios), use_container_width=True)
        with c2:
            # CRAR
            fig = base_fig()
            for s_name in selected_scenarios:
                sc = SCENARIOS[s_name]
                df, _ = run_stress_test(bank, sc)
                fig.add_trace(go.Scatter(x=df["quarter"], y=df["crar"],
                    name=s_name.split(" ",1)[1],
                    line=dict(color=sc["color"], width=2.5), mode="lines+markers"))
            fig.add_hline(y=11.5, line=dict(color="#dc3545", dash="dash", width=1.5),
                          annotation_text="CRAR Min (11.5%)")
            fig.update_layout(title="Total CRAR Under Stress (%)",
                title_font=dict(color=COLORS["gold"], size=15, family="Playfair Display"),
                yaxis_title="CRAR (%)", height=380)
            st.plotly_chart(fig, use_container_width=True)

        # Waterfall for all selected
        st.markdown(section_header("Capital Walk — Waterfall Charts", "🏗️"), unsafe_allow_html=True)
        wf_cols = st.columns(min(len(selected_scenarios), 3))
        for i, s_name in enumerate(selected_scenarios[:3]):
            with wf_cols[i % 3]:
                st.plotly_chart(plot_capital_waterfall(bank, s_name), use_container_width=True)

        # Capital buffer analysis
        st.markdown(section_header("Capital Buffer Analysis", "🛡️"), unsafe_allow_html=True)
        buf_data = []
        for s_name in selected_scenarios:
            sc = SCENARIOS[s_name]
            _, summary = run_stress_test(bank, sc)
            buf_data.append({
                "Scenario": s_name,
                "Min CET1 (%)": round(summary["min_cet1"], 2),
                "Buffer above 8% floor (%)": round(summary["min_cet1"] - 8.0, 2),
                "Buffer above 9.5% incl. conservation (%)": round(summary["min_cet1"] - 9.5, 2),
                "Capital at Risk (₹Cr)": round(
                    max(0, (bank["cet1_ratio"] - summary["min_cet1"]) / 100 * bank["rwa"]), 0),
                "Pass / Fail CET1": "✅ Pass" if summary["min_cet1"] >= 8.0 else "❌ Fail",
            })
        st.dataframe(pd.DataFrame(buf_data), use_container_width=True, hide_index=True)

    # ═══════════════════════════════════════════════════════════════
    # PAGE 5: LIQUIDITY ANALYSIS
    # ═══════════════════════════════════════════════════════════════
    elif page == "💧 Liquidity Analysis":
        st.markdown(section_header("Liquidity Stress Testing", "💧"), unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            st.plotly_chart(plot_liquidity_stress(bank, selected_scenarios), use_container_width=True)
        with c2:
            # HQLA evolution
            fig = base_fig()
            for s_name in selected_scenarios:
                sc = SCENARIOS[s_name]
                df, _ = run_stress_test(bank, sc)
                fig.add_trace(go.Bar(x=df["quarter"], y=df["hqla"],
                    name=s_name.split(" ",1)[1], marker_color=sc["color"]))
            fig.update_layout(title="HQLA Buffer (₹ Crore)",
                title_font=dict(color=COLORS["gold"], size=15, family="Playfair Display"),
                yaxis_title="₹ Crore", barmode="group", height=380)
            st.plotly_chart(fig, use_container_width=True)

        st.markdown(section_header("Deposit Run-off Analysis", "🏃"), unsafe_allow_html=True)
        fig_dep = base_fig()
        for s_name in selected_scenarios:
            sc = SCENARIOS[s_name]
            df, _ = run_stress_test(bank, sc)
            fig_dep.add_trace(go.Scatter(x=df["quarter"], y=df["deposits"],
                name=s_name.split(" ",1)[1],
                line=dict(color=sc["color"], width=2.5), mode="lines+markers"))
        fig_dep.update_layout(title="Total Deposits Under Stress (₹ Crore)",
            title_font=dict(color=COLORS["gold"], size=15, family="Playfair Display"),
            yaxis_title="₹ Crore", height=360)
        st.plotly_chart(fig_dep, use_container_width=True)

        # Liquidity metrics table
        liq_data = []
        for s_name in selected_scenarios:
            sc = SCENARIOS[s_name]
            df, summary = run_stress_test(bank, sc)
            liq_data.append({
                "Scenario": s_name,
                "Starting LCR (%)": bank["lcr"],
                "Min LCR (%)": round(summary["min_lcr"], 0),
                "Deposit Run-off (%)": f"{sc['deposit_runoff']*100:.0f}%",
                "HQLA Floor (₹Cr)": round(df["hqla"].min(), 0),
                "Liquidity Status": "❌ Breach" if summary["min_lcr"] < 100 else "✅ Pass",
            })
        st.dataframe(pd.DataFrame(liq_data), use_container_width=True, hide_index=True)

    # ═══════════════════════════════════════════════════════════════
    # PAGE 6: SENSITIVITY & HEATMAP
    # ═══════════════════════════════════════════════════════════════
    elif page == "🔥 Sensitivity & Heatmap":
        st.markdown(section_header("Sensitivity Heatmap: NPA Multiplier × Rate Shock", "🔥"), unsafe_allow_html=True)
        st.plotly_chart(plot_sensitivity_heatmap(bank), use_container_width=True)

        st.markdown(section_header("Single-Variable Sensitivity Analysis", "📊"), unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        base_sc = SCENARIOS["🔴 Severe Stress"].copy()
        with c1:
            npa_vals = np.linspace(1.0, 4.0, 15)
            npa_sens = run_sensitivity(bank, "npa_multiplier", npa_vals, base_sc)
            fig = base_fig()
            fig.add_trace(go.Scatter(x=npa_vals, y=npa_sens["min_cet1"],
                name="Min CET1", line=dict(color=COLORS["gold"], width=2.5),
                fill="tozeroy", fillcolor=hex_to_rgba(COLORS["gold"], 0.08)))
            fig.add_hline(y=8.0, line=dict(color=COLORS["red"], dash="dash"),
                          annotation_text="CET1 Min 8%")
            fig.update_layout(title="CET1 Sensitivity to NPA Multiplier",
                title_font=dict(color=COLORS["gold"], size=14, family="Playfair Display"),
                xaxis_title="NPA Multiplier", yaxis_title="Min CET1 (%)", height=360)
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            rate_vals = np.linspace(0, 400, 15)
            rate_sens = run_sensitivity(bank, "rate_shock_bps", rate_vals, base_sc)
            fig2 = base_fig()
            fig2.add_trace(go.Scatter(x=rate_vals, y=rate_sens["min_cet1"],
                name="Min CET1", line=dict(color=COLORS["lightblue"], width=2.5),
                fill="tozeroy", fillcolor=hex_to_rgba(COLORS["lightblue"], 0.08)))
            fig2.add_hline(y=8.0, line=dict(color=COLORS["red"], dash="dash"),
                           annotation_text="CET1 Min 8%")
            fig2.update_layout(title="CET1 Sensitivity to Rate Shock (bps)",
                title_font=dict(color=COLORS["gold"], size=14, family="Playfair Display"),
                xaxis_title="Rate Shock (bps)", yaxis_title="Min CET1 (%)", height=360)
            st.plotly_chart(fig2, use_container_width=True)

        # Equity shock sensitivity
        eq_vals = np.linspace(0, 0.70, 15)
        eq_sens = run_sensitivity(bank, "equity_shock_pct", eq_vals, base_sc)
        fig3 = base_fig()
        fig3.add_trace(go.Scatter(x=eq_vals*100, y=eq_sens["min_cet1"],
            name="Min CET1", line=dict(color=COLORS["orange"], width=2.5),
            fill="tozeroy", fillcolor=hex_to_rgba(COLORS["orange"], 0.08)))
        fig3.add_hline(y=8.0, line=dict(color=COLORS["red"], dash="dash"),
                       annotation_text="CET1 Min 8%")
        fig3.update_layout(title="CET1 Sensitivity to Equity Market Crash (%)",
            title_font=dict(color=COLORS["gold"], size=14, family="Playfair Display"),
            xaxis_title="Equity Shock (%)", yaxis_title="Min CET1 (%)", height=350)
        st.plotly_chart(fig3, use_container_width=True)

    # ═══════════════════════════════════════════════════════════════
    # PAGE 7: CUSTOM SCENARIO BUILDER
    # ═══════════════════════════════════════════════════════════════
    elif page == "🛠️ Custom Scenario Builder":
        st.markdown(section_header("Build Your Own Stress Scenario", "🛠️"), unsafe_allow_html=True)

        with st.form("custom_scenario"):
            col1, col2, col3 = st.columns(3)
            with col1:
                st.markdown(f"<div style='color:{COLORS['gold']};font-weight:700;'>📉 Credit Risk</div>", unsafe_allow_html=True)
                npa_mult = st.slider("NPA Multiplier", 1.0, 5.0, 2.0, 0.1,
                                     help="How many times NPAs grow vs current")
                deposit_ro = st.slider("Deposit Run-off (%)", 0.0, 30.0, 8.0, 0.5) / 100
                property_sh = st.slider("Property Price Shock (%)", 0.0, 50.0, 20.0, 1.0) / 100

            with col2:
                st.markdown(f"<div style='color:{COLORS['gold']};font-weight:700;'>📈 Market Risk</div>", unsafe_allow_html=True)
                rate_sh = st.slider("Rate Shock (bps)", 0, 500, 200, 25)
                eq_sh = st.slider("Equity Market Fall (%)", 0.0, 70.0, 30.0, 1.0) / 100
                inr_sh = st.slider("INR Depreciation (%)", 0.0, 40.0, 10.0, 1.0) / 100

            with col3:
                st.markdown(f"<div style='color:{COLORS['gold']};font-weight:700;'>🌍 Macro Variables</div>", unsafe_allow_html=True)
                gdp_shock_val = st.slider("GDP Shock (pp, peak quarter)", 0.0, 10.0, 3.5, 0.1)
                fee_sh = st.slider("Fee Income Fall (%)", 0.0, 50.0, 20.0, 1.0) / 100
                cs_wide = st.slider("Credit Spread Widening (bps)", 0, 600, 200, 25)
                scenario_name_custom = st.text_input("Scenario Name", "My Custom Scenario")

            submitted = st.form_submit_button("🚀 Run Custom Stress Test",
                                              use_container_width=True)

        if submitted:
            gdp_path = np.array([0, -gdp_shock_val*0.5, -gdp_shock_val, -gdp_shock_val*0.9,
                                  -gdp_shock_val*0.7, -gdp_shock_val*0.4,
                                  -gdp_shock_val*0.2, -gdp_shock_val*0.1])
            custom_sc = {
                "color": "#cc99ff",
                "severity": "Custom",
                "description": scenario_name_custom,
                "gdp_shock": gdp_path.tolist(),
                "rate_shock_bps": rate_sh,
                "equity_shock_pct": -eq_sh,
                "inr_depreciation": inr_sh,
                "npa_multiplier": npa_mult,
                "deposit_runoff": deposit_ro,
                "property_shock": -property_sh,
                "credit_spread_widening": cs_wide,
                "fee_income_shock": -fee_sh,
            }
            df_c, summary_c = run_stress_test(bank, custom_sc)

            st.success(f"✅ Custom scenario '{scenario_name_custom}' executed successfully!")

            r1, r2, r3, r4 = st.columns(4)
            with r1: st.markdown(metric_card("Min CET1", f"{summary_c['min_cet1']:.2f}%",
                "BREACH!" if summary_c['min_cet1'] < 8 else "Pass",
                "neg" if summary_c['min_cet1'] < 8 else "pos"), unsafe_allow_html=True)
            with r2: st.markdown(metric_card("Peak NPA", f"{summary_c['peak_npa']:.1f}%"), unsafe_allow_html=True)
            with r3: st.markdown(metric_card("Min LCR", f"{summary_c['min_lcr']:.0f}%",
                "BREACH!" if summary_c['min_lcr'] < 100 else "Pass",
                "neg" if summary_c['min_lcr'] < 100 else "pos"), unsafe_allow_html=True)
            with r4: st.markdown(metric_card("Cum. Provisions",
                f"₹{summary_c['cumulative_provisions']:,.0f}Cr"), unsafe_allow_html=True)

            c_plot1, c_plot2 = st.columns(2)
            with c_plot1:
                st.plotly_chart(plot_pl_decomposition(bank, "🟢 Baseline"), use_container_width=True)
            with c_plot2:
                # Custom capital path
                fig_cap = base_fig()
                fig_cap.add_trace(go.Scatter(x=df_c["quarter"], y=df_c["cet1_ratio"],
                    name=scenario_name_custom, line=dict(color="#cc99ff", width=3),
                    mode="lines+markers", marker=dict(size=8)))
                bl_df, _ = run_stress_test(bank, SCENARIOS["🟢 Baseline"])
                fig_cap.add_trace(go.Scatter(x=bl_df["quarter"], y=bl_df["cet1_ratio"],
                    name="Baseline", line=dict(color=COLORS["green"], dash="dot", width=2)))
                fig_cap.add_hline(y=8.0, line=dict(color=COLORS["red"], dash="dash"))
                fig_cap.update_layout(title="CET1 Ratio: Custom vs Baseline",
                    title_font=dict(color=COLORS["gold"], size=14, family="Playfair Display"),
                    height=380)
                st.plotly_chart(fig_cap, use_container_width=True)

            st.dataframe(df_c[["quarter","gdp_growth","gross_npa_ratio",
                                "net_profit","cet1_ratio","lcr"]].round(2),
                         use_container_width=True, hide_index=True)

    # ═══════════════════════════════════════════════════════════════
    # PAGE 8: SCENARIO COMPARISON
    # ═══════════════════════════════════════════════════════════════
    elif page == "📋 Scenario Comparison":
        st.markdown(section_header("Full Scenario Comparison Table", "📋"), unsafe_allow_html=True)
        st.markdown(scenario_summary_table(bank, selected_scenarios), unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(section_header("Side-by-Side Capital & NPA Paths", "📊"), unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            st.plotly_chart(plot_multi_scenario_cet1(bank, selected_scenarios), use_container_width=True)
        with c2:
            st.plotly_chart(plot_npa_evolution(bank, selected_scenarios), use_container_width=True)

        c3, c4 = st.columns(2)
        with c3:
            st.plotly_chart(plot_liquidity_stress(bank, selected_scenarios), use_container_width=True)
        with c4:
            # Net profit comparison
            fig_np = base_fig()
            for s_name in selected_scenarios:
                sc = SCENARIOS[s_name]
                df, _ = run_stress_test(bank, sc)
                fig_np.add_trace(go.Bar(x=df["quarter"], y=df["net_profit"],
                    name=s_name.split(" ",1)[1], marker_color=sc["color"]))
            fig_np.update_layout(title="Quarterly Net Profit Under Stress (₹Cr)",
                title_font=dict(color=COLORS["gold"], size=14, family="Playfair Display"),
                barmode="group", yaxis_title="₹ Crore", height=380)
            st.plotly_chart(fig_np, use_container_width=True)

        # Provisions comparison
        st.markdown(section_header("Cumulative Provisions by Scenario (₹ Crore)", "💸"), unsafe_allow_html=True)
        fig_prov = base_fig()
        labels, prov_vals, colors_list = [], [], []
        for s_name in selected_scenarios:
            sc = SCENARIOS[s_name]
            _, summary = run_stress_test(bank, sc)
            labels.append(s_name.split(" ",1)[1])
            prov_vals.append(summary["cumulative_provisions"])
            colors_list.append(sc["color"])
        fig_prov.add_trace(go.Bar(x=labels, y=prov_vals,
            marker_color=colors_list, text=[f"₹{v:,.0f}Cr" for v in prov_vals],
            textposition="outside", textfont=dict(color=COLORS["text"])))
        fig_prov.update_layout(title="Cumulative Provisions (8Q)",
            title_font=dict(color=COLORS["gold"], size=14, family="Playfair Display"),
            yaxis_title="₹ Crore", height=380)
        st.plotly_chart(fig_prov, use_container_width=True)

    # ═══════════════════════════════════════════════════════════════
    # PAGE 9: REVERSE STRESS TEST
    # ═══════════════════════════════════════════════════════════════
    elif page == "🔄 Reverse Stress Test":
        st.markdown(section_header("Reverse Stress Testing — What Breaks the Bank?", "🔄"), unsafe_allow_html=True)
        st.markdown(f"""
        <div class='metric-card'>
            <b style='color:{COLORS["gold"]}'>What is Reverse Stress Testing?</b><br>
            <span style='color:{COLORS["muted"]};'>Instead of asking "what happens if GDP falls 5%?",
            reverse stress testing asks: <i>"what combination of shocks reduces CET1 below 8%?"</i>
            It identifies the exact Achilles' heel of the balance sheet.</span>
        </div>""", unsafe_allow_html=True)

        st.markdown(section_header("Critical Threshold Analysis", "⚠️"), unsafe_allow_html=True)

        # Find the NPA multiplier that breaks CET1
        base_sc = SCENARIOS["🟠 Moderate Stress"].copy()
        npa_breaks = None
        for nm in np.linspace(1.0, 5.0, 100):
            sc_test = base_sc.copy()
            sc_test["npa_multiplier"] = nm
            _, s = run_stress_test(bank, sc_test)
            if s["min_cet1"] < 8.0:
                npa_breaks = nm
                break

        rate_breaks = None
        for rb in range(0, 600, 5):
            sc_test = base_sc.copy()
            sc_test["rate_shock_bps"] = rb
            _, s = run_stress_test(bank, sc_test)
            if s["min_cet1"] < 8.0:
                rate_breaks = rb
                break

        equity_breaks = None
        for eq in np.linspace(0, 1.0, 100):
            sc_test = base_sc.copy()
            sc_test["equity_shock_pct"] = -eq
            _, s = run_stress_test(bank, sc_test)
            if s["min_cet1"] < 8.0:
                equity_breaks = eq
                break

        c1, c2, c3 = st.columns(3)
        with c1:
            val = f"{npa_breaks:.1f}×" if npa_breaks else "Resilient >5×"
            st.markdown(metric_card("NPA Break-Even Multiplier",
                val, "CET1 breaches at this level",
                "neg" if npa_breaks and npa_breaks < 3 else "warn"), unsafe_allow_html=True)
        with c2:
            val2 = f"{rate_breaks}bps" if rate_breaks else "Resilient >600bps"
            st.markdown(metric_card("Rate Shock Break-Even",
                val2, "CET1 breaches at this level",
                "neg" if rate_breaks and rate_breaks < 200 else "warn"), unsafe_allow_html=True)
        with c3:
            val3 = f"{equity_breaks*100:.0f}%" if equity_breaks else "Resilient >100%"
            st.markdown(metric_card("Equity Crash Break-Even",
                val3, "CET1 breaches at this level",
                "neg" if equity_breaks and equity_breaks < 0.4 else "warn"), unsafe_allow_html=True)

        # Break-even visualisation
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(section_header("Break-Even Frontier: NPA × Rate Shock", "📍"), unsafe_allow_html=True)

        npa_range = np.linspace(1.0, 4.5, 20)
        rate_range = np.linspace(0, 400, 20)
        frontier_npa, frontier_rate = [], []
        for nm in npa_range:
            for rb in rate_range:
                sc_test = base_sc.copy()
                sc_test["npa_multiplier"] = nm
                sc_test["rate_shock_bps"] = int(rb)
                _, s = run_stress_test(bank, sc_test)
                if s["min_cet1"] < 8.0:
                    frontier_npa.append(nm)
                    frontier_rate.append(rb)
                    break

        if frontier_npa:
            fig_rev = base_fig()
            fig_rev.add_trace(go.Scatter(
                x=frontier_npa, y=frontier_rate,
                mode="lines+markers",
                line=dict(color=COLORS["red"], width=3),
                fill="tozeroy", fillcolor=hex_to_rgba(COLORS["red"], 0.13),
                name="Failure Frontier",
                marker=dict(size=6, color=COLORS["red"]),
            ))
            fig_rev.add_annotation(
                x=np.mean(frontier_npa), y=np.mean(frontier_rate) * 0.4,
                text="BANK SURVIVES ✅",
                font=dict(color=COLORS["green"], size=14, family="Playfair Display"),
                showarrow=False,
            )
            fig_rev.add_annotation(
                x=np.mean(frontier_npa), y=np.mean(frontier_rate) * 1.5,
                text="CET1 BREACH ❌",
                font=dict(color=COLORS["red"], size=14, family="Playfair Display"),
                showarrow=False,
            )
            fig_rev.update_layout(
                title="Reverse Stress: Failure Frontier (NPA Multiplier vs Rate Shock)",
                title_font=dict(color=COLORS["gold"], size=15, family="Playfair Display"),
                xaxis_title="NPA Multiplier",
                yaxis_title="Rate Shock (bps)",
                height=420,
            )
            st.plotly_chart(fig_rev, use_container_width=True)

        # Management actions
        st.markdown(section_header("Management Actions Available", "🔧"), unsafe_allow_html=True)
        actions = pd.DataFrame({
            "Action": [
                "Capital raising via QIP/Rights issue",
                "Accelerate NPA recoveries / NARCL referral",
                "Reduce risk appetite — cut new corporate lending",
                "Sell AFS portfolio to crystallise gains early",
                "Increase PCR to 70% pre-emptively",
                "Activate contingent liquidity facilities (RBI SLF)",
                "Increase CASA mobilisation",
                "Reduce dividend payout ratio",
            ],
            "Capital Impact (₹Cr est.)": ["+3,000–5,000", "+500–1,200", "+800 RWA relief",
                                          "+400–700", "-300 provisions", "Liquidity only",
                                          "Funding cost -15bps", "+200–400"],
            "Timeline": ["3–6 months","6–12 months","Immediate","1–2 months",
                          "Immediate","Immediate","6–18 months","Annual"],
            "Priority": ["⭐⭐⭐ Critical","⭐⭐⭐ Critical","⭐⭐ High","⭐⭐ High",
                          "⭐⭐ High","⭐ Medium","⭐ Medium","⭐ Medium"],
        })
        st.dataframe(actions, use_container_width=True, hide_index=True)



    # ═══════════════════════════════════════════════════════════════
    # PAGE: UPLOAD BANK DATA
    # ═══════════════════════════════════════════════════════════════
    elif page == "📤 Upload Bank Data":
        _g  = COLORS["gold"]; _db = COLORS["darkblue"]; _cb = COLORS["cardBg"]
        _mb = COLORS["midblue"]; _bd = COLORS["bgDark"]; _lb = COLORS["lightblue"]
        _gr = COLORS["green"]; _rd = COLORS["red"]; _mt = COLORS["muted"]

        # Hero
        st.markdown(
            "<div style='background:linear-gradient(135deg," + _db + "," + _cb + ");"
            "border:2px solid " + _g + ";border-radius:10px;"
            "padding:22px 32px;margin-bottom:20px;'>"
            "<div style='font-family:Playfair Display,serif;font-size:1.7rem;"
            "font-weight:900;color:" + _g + ";'>📤 Upload Your Bank's Financial Data</div>"
            "<div style='color:" + _lb + ";font-size:0.9rem;margin-top:6px;'>"
            "Download the template → fill in your bank's data → upload → run stress tests on real data"
            "</div></div>",
            unsafe_allow_html=True
        )

        # ── Step 1: Download template ────────────────────────────
        st.markdown(
            "<div style='color:" + _g + ";font-family:Playfair Display,serif;"
            "font-size:1.1rem;font-weight:700;border-bottom:2px solid " + _mb + ";"
            "padding-bottom:6px;margin:0 0 12px 0;'>Step 1 — Download the Input Template</div>",
            unsafe_allow_html=True
        )

        col_dl, col_info = st.columns([1, 2])
        with col_dl:
            template_bytes = create_excel_template()
            st.download_button(
                label="⬇️  Download Template (.xlsx)",
                data=template_bytes,
                file_name="Mountain_Path_Bank_Stress_Test_Template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                help="Download the Excel template, fill in your bank data, then upload below"
            )
        with col_info:
            st.markdown(
                "<div style='background:" + _bd + ";border-left:3px solid " + _g + ";"
                "border-radius:0 6px 6px 0;padding:10px 14px;font-size:0.83rem;'>"
                "<b style='color:" + _g + ";'>Template contains 6 sheets:</b><br>"
                "<span style='color:#d0dff0;'>"
                "📊 BS_Assets &nbsp;·&nbsp; 📊 BS_Liabilities &nbsp;·&nbsp; 📊 Income_Stmt<br>"
                "📊 Asset_Quality &nbsp;·&nbsp; 📊 Capital_Ratios &nbsp;·&nbsp; 📊 Duration_Risk"
                "</span><br><br>"
                "<b style='color:" + _g + ";'>Fill only the light-blue cells.</b>"
                "<span style='color:" + _mt + ";'> All values in ₹ Crore (unless stated). "
                "Percentages as numbers (7.5 not 0.075).</span>"
                "</div>",
                unsafe_allow_html=True
            )

        st.markdown("<br>", unsafe_allow_html=True)

        # ── Step 2: Template sheet guide ────────────────────────
        st.markdown(
            "<div style='color:" + _g + ";font-family:Playfair Display,serif;"
            "font-size:1.1rem;font-weight:700;border-bottom:2px solid " + _mb + ";"
            "padding-bottom:6px;margin:0 0 12px 0;'>Step 2 — Fill in Your Data</div>",
            unsafe_allow_html=True
        )
        sheets_info = [
            ("📊 BS_Assets",       "Balance Sheet: Assets",
             "Total assets, gross loans (retail/corp/MSME/agri), investments (HTM/AFS), HQLA, equity portfolio",
             _rd),
            ("📊 BS_Liabilities",  "Balance Sheet: Liabilities & Capital",
             "Deposits (CASA/term), wholesale funding, sub-debt, CET1, AT1, Tier 2, RWA, net worth",
             "#fd7e14"),
            ("📊 Income_Stmt",     "Annual Income Statement",
             "NII, fee income, trading income, operating costs, provisions, PAT — auto-calculates PPOP and PBT",
             _g),
            ("📊 Asset_Quality",   "NPA Ratios & Sector Exposure",
             "Gross/Net NPA %, SMA-2, PCR, C/D ratio, plus ₹Cr exposure to 8 sectors",
             "#cc66ff"),
            ("📊 Capital_Ratios",  "Capital & Liquidity Ratios",
             "CET1, CRAR, Leverage, LCR, NSFR, CASA ratio, NIM, ROE, ROA, Cost-to-Income",
             _gr),
            ("📊 Duration_Risk",   "Duration & Market Risk",
             "Avg. modified duration of AFS/HTM/Liabilities, repricing gap, equity beta, net FX position",
             _lb),
        ]
        cols = st.columns(3)
        for i, (sname, stitle, sdesc, scolor) in enumerate(sheets_info):
            with cols[i % 3]:
                st.markdown(
                    "<div style='background:" + _cb + ";border:1px solid " + scolor + "33;"
                    "border-top:3px solid " + scolor + ";border-radius:8px;"
                    "padding:12px 14px;margin:4px 0;min-height:110px;'>"
                    "<div style='color:" + scolor + ";font-weight:700;font-size:0.85rem;"
                    "font-family:Playfair Display,serif;'>" + sname + "</div>"
                    "<div style='color:#ffffff;font-size:0.8rem;font-weight:600;margin:3px 0;'>" + stitle + "</div>"
                    "<div style='color:" + _mt + ";font-size:0.76rem;line-height:1.5;'>" + sdesc + "</div>"
                    "</div>",
                    unsafe_allow_html=True
                )

        st.markdown("<br>", unsafe_allow_html=True)

        # ── Step 3: Upload ───────────────────────────────────────
        st.markdown(
            "<div style='color:" + _g + ";font-family:Playfair Display,serif;"
            "font-size:1.1rem;font-weight:700;border-bottom:2px solid " + _mb + ";"
            "padding-bottom:6px;margin:0 0 12px 0;'>Step 3 — Upload Completed Template</div>",
            unsafe_allow_html=True
        )

        uploaded = st.file_uploader(
            "Upload your completed Excel template",
            type=["xlsx"],
            help="Upload the filled Mountain Path Bank Stress Test Template",
            label_visibility="collapsed"
        )

        if uploaded is not None:
            file_bytes = uploaded.read()
            with st.spinner("Parsing your bank data..."):
                parsed_bank, errors = parse_uploaded_excel(file_bytes)

            if errors:
                for err in errors:
                    st.error(err)
                st.markdown(
                    "<div style='background:#330d10;border:1px solid " + _rd + ";"
                    "border-radius:6px;padding:12px 16px;color:#ffaaaa;font-size:0.85rem;'>"
                    "⚠️ Please fix the errors above and re-upload the corrected file."
                    "</div>",
                    unsafe_allow_html=True
                )
            else:
                # Success — save to session state
                bank_name = parsed_bank.get("name", "Uploaded Bank")
                st.session_state.bank_data   = parsed_bank
                st.session_state.bank_source = "uploaded"
                st.session_state.bank_label  = bank_name + " (Uploaded)"
                bank = parsed_bank  # update local reference

                st.success(f"✅ Data loaded successfully for **{bank_name}** — all pages now use your data!")

                # ── Preview: key metrics ─────────────────────────
                st.markdown("<br>", unsafe_allow_html=True)
                st.markdown(
                    "<div style='color:" + _g + ";font-family:Playfair Display,serif;"
                    "font-size:1.0rem;font-weight:700;margin-bottom:10px;'>"
                    "📊 Uploaded Data Preview</div>",
                    unsafe_allow_html=True
                )

                # Row 1 metrics
                m1, m2, m3, m4, m5 = st.columns(5)
                metrics_r1 = [
                    ("Total Assets", f"₹{bank['total_assets']:,.0f} Cr", ""),
                    ("CET1 Ratio",   f"{bank['cet1_ratio']:.2f}%",
                     "🟢 Pass" if bank["cet1_ratio"] >= 8.0 else "🔴 Below Min"),
                    ("Gross NPA",    f"{bank['gross_npa_ratio']:.2f}%",
                     "🟢 <7%" if bank["gross_npa_ratio"] < 7 else "🟠 >7%"),
                    ("LCR",          f"{bank['lcr']:.1f}%",
                     "🟢 Pass" if bank["lcr"] >= 100 else "🔴 Below Min"),
                    ("NIM",          f"{bank['nim']:.2f}%", ""),
                ]
                for col, (lbl, val, delta) in zip([m1,m2,m3,m4,m5], metrics_r1):
                    with col:
                        st.markdown(
                            "<div style='background:" + _cb + ";border:1px solid " + _mb + ";"
                            "border-top:3px solid " + _g + ";border-radius:8px;"
                            "padding:12px 14px;'>"
                            "<div style='color:" + _mt + ";font-size:0.72rem;text-transform:uppercase;"
                            "letter-spacing:1px;'>" + lbl + "</div>"
                            "<div style='color:" + _g + ";font-size:1.5rem;font-weight:700;"
                            "font-family:Playfair Display,serif;'>" + val + "</div>"
                            "<div style='font-size:0.78rem;color:#d0dff0;'>" + delta + "</div>"
                            "</div>",
                            unsafe_allow_html=True
                        )

                st.markdown("<br>", unsafe_allow_html=True)

                # Two-column detail tables
                tc1, tc2 = st.columns(2)
                with tc1:
                    st.markdown(
                        "<div style='color:" + _g + ";font-weight:700;font-size:0.88rem;"
                        "margin-bottom:6px;'>📋 Balance Sheet Summary</div>",
                        unsafe_allow_html=True
                    )
                    bs_items = [
                        ("Gross Loans",     f"₹{bank['gross_loans']:,.0f} Cr"),
                        ("Gross NPA",       f"₹{bank['gross_npa']:,.0f} Cr"),
                        ("AFS Investments", f"₹{bank['investments_afs']:,.0f} Cr"),
                        ("HQLA",            f"₹{bank['cash_hqla']:,.0f} Cr"),
                        ("Total Deposits",  f"₹{bank['total_deposits']:,.0f} Cr"),
                        ("CASA Deposits",   f"₹{bank['casa_deposits']:,.0f} Cr ({bank['casa_ratio']:.1f}%)"),
                        ("CET1 Capital",    f"₹{bank['cet1_capital']:,.0f} Cr"),
                        ("RWA",             f"₹{bank['rwa']:,.0f} Cr"),
                    ]
                    rows_html = "".join(
                        "<tr><td style='color:#d0dff0;padding:5px 10px;border-bottom:1px solid #1e3a5f;"
                        "font-size:0.82rem;'>" + lbl + "</td>"
                        "<td style='color:" + _g + ";font-weight:700;padding:5px 10px;"
                        "border-bottom:1px solid #1e3a5f;font-size:0.82rem;text-align:right;'>"
                        + val + "</td></tr>"
                        for lbl, val in bs_items
                    )
                    st.markdown(
                        "<table style='width:100%;background:" + _cb + ";"
                        "border-radius:8px;border-collapse:collapse;'>"
                        "<thead><tr>"
                        "<th style='background:" + _db + ";color:" + _g + ";padding:7px 10px;"
                        "font-size:0.78rem;text-align:left;'>Item</th>"
                        "<th style='background:" + _db + ";color:" + _g + ";padding:7px 10px;"
                        "font-size:0.78rem;text-align:right;'>Value</th>"
                        "</tr></thead><tbody>" + rows_html + "</tbody></table>",
                        unsafe_allow_html=True
                    )

                with tc2:
                    st.markdown(
                        "<div style='color:" + _g + ";font-weight:700;font-size:0.88rem;"
                        "margin-bottom:6px;'>📋 Income & Capital Summary</div>",
                        unsafe_allow_html=True
                    )
                    inc_items = [
                        ("Net Interest Income",  f"₹{bank['net_interest_income']:,.0f} Cr"),
                        ("Fee Income",           f"₹{bank['fee_income']:,.0f} Cr"),
                        ("Operating Costs",      f"₹{bank['operating_costs']:,.0f} Cr"),
                        ("Provisions Charge",    f"₹{bank['provisions_charge']:,.0f} Cr"),
                        ("PAT",                  f"₹{bank['pat']:,.0f} Cr"),
                        ("CET1 Ratio",           f"{bank['cet1_ratio']:.2f}%"),
                        ("Total CRAR",           f"{bank['crar']:.2f}%"),
                        ("Gross NPA Ratio",      f"{bank['gross_npa_ratio']:.2f}%"),
                    ]
                    rows_html2 = "".join(
                        "<tr><td style='color:#d0dff0;padding:5px 10px;border-bottom:1px solid #1e3a5f;"
                        "font-size:0.82rem;'>" + lbl + "</td>"
                        "<td style='color:" + _g + ";font-weight:700;padding:5px 10px;"
                        "border-bottom:1px solid #1e3a5f;font-size:0.82rem;text-align:right;'>"
                        + val + "</td></tr>"
                        for lbl, val in inc_items
                    )
                    st.markdown(
                        "<table style='width:100%;background:" + _cb + ";"
                        "border-radius:8px;border-collapse:collapse;'>"
                        "<thead><tr>"
                        "<th style='background:" + _db + ";color:" + _g + ";padding:7px 10px;"
                        "font-size:0.78rem;text-align:left;'>Item</th>"
                        "<th style='background:" + _db + ";color:" + _g + ";padding:7px 10px;"
                        "font-size:0.78rem;text-align:right;'>Value</th>"
                        "</tr></thead><tbody>" + rows_html2 + "</tbody></table>",
                        unsafe_allow_html=True
                    )

                st.markdown("<br>", unsafe_allow_html=True)
                st.info("✅ Navigate to any page from the sidebar — all stress tests will now run on your uploaded bank data.")

        # ── Reset to demo ────────────────────────────────────────
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(
            "<div style='color:" + _mt + ";font-size:0.8rem;margin-bottom:6px;'>"
            "Currently loaded: <b style='color:#d0dff0;'>" +
            st.session_state.get("bank_label", "Demo Bank") + "</b></div>",
            unsafe_allow_html=True
        )
        if st.button("🔄 Reset to Demo Bank (Mountain Path Bank Ltd.)",
                     use_container_width=False):
            st.session_state.bank_data   = generate_bank_data()
            st.session_state.bank_source = "demo"
            st.session_state.bank_label  = "Mountain Path Bank Ltd. (Demo)"
            st.success("Reset to demo bank data.")
            st.rerun()

    # ═══════════════════════════════════════════════════════════════
    # PAGE 10: ABOUT THE PLATFORM
    # ═══════════════════════════════════════════════════════════════
    elif page == "🏛 About the Platform":
        # Hero banner
        st.markdown(f"""
        <div style='background:linear-gradient(135deg,{COLORS["darkblue"]},{COLORS["midblue"]});
                    border:2px solid {COLORS["gold"]};border-radius:12px;
                    padding:36px 40px;margin-bottom:24px;text-align:center;'>
            <div style='font-family:Playfair Display,serif;font-size:2.4rem;
                        font-weight:900;color:{COLORS["gold"]};letter-spacing:2px;'>
                🏦 Bank Stress Testing Lab
            </div>
            <div style='color:{COLORS["lightblue"]};font-size:1.05rem;
                        margin:10px 0 6px 0;letter-spacing:1px;'>
                THE MOUNTAIN PATH — World of Finance
            </div>
            <div style='color:{COLORS["muted"]};font-size:0.88rem;'>
                Advanced Financial Risk Modelling Platform &nbsp;|&nbsp;
                Powered by Python · Streamlit · Plotly
            </div>
        </div>
        """, unsafe_allow_html=True)

        c1, c2 = st.columns([1.1, 0.9])
        with c1:
            st.markdown(section_header("About This Platform", "🏦"), unsafe_allow_html=True)
            st.markdown(f"""
            <div class='metric-card' style='line-height:1.9;'>
                <p style='color:{COLORS["text"]};'>
                The <b style='color:{COLORS["gold"]};'>Bank Stress Testing Lab</b> is a
                comprehensive, interactive financial risk modelling platform built to simulate
                and analyse the resilience of a bank's balance sheet, capital position, and
                liquidity buffers under a wide spectrum of adverse macroeconomic and
                market scenarios.
                </p>
                <p style='color:{COLORS["text"]};'>
                This platform implements industry-standard methodologies aligned with
                <b style='color:{COLORS["lightblue"]};'>RBI Stress Testing Guidelines</b>,
                <b style='color:{COLORS["lightblue"]};'>Basel III / Basel IV (FRTB)</b>
                frameworks, and global best practices from the EBA and US Federal Reserve's
                DFAST programme.
                </p>
                <p style='color:{COLORS["text"]};'>
                All bank data used is <b style='color:{COLORS["gold"]};'>fully synthetic</b>
                — designed to represent a realistic mid-sized Indian private sector scheduled
                commercial bank — enabling safe, unrestricted exploration of stress scenarios
                without any confidentiality concerns.
                </p>
            </div>
            """, unsafe_allow_html=True)

            st.markdown(section_header("What the Platform Does", "⚙️"), unsafe_allow_html=True)
            features = [
                ("🏠 Dashboard", "Live KPI tracking, scenario CET1 paths, sector concentration, resilience radar"),
                ("📊 Balance Sheet & Income", "Full synthetic B/S (assets, liabilities, capital) and P&L with ratio analysis"),
                ("🎯 Run Stress Tests", "Deep-dive single-scenario analysis: P&L decomposition, capital waterfall, quarterly projections"),
                ("📈 Capital Analysis", "CET1 / CRAR paths, capital buffer analysis, capital walk waterfall for all scenarios"),
                ("💧 Liquidity Analysis", "LCR stress paths, HQLA buffer evolution, deposit run-off modelling"),
                ("🔥 Sensitivity & Heatmap", "Two-variable sensitivity heatmap (NPA × Rate), single-variable sweep charts"),
                ("🛠️ Custom Scenario Builder", "Build any bespoke stress scenario with 8 independent shock parameters"),
                ("📋 Scenario Comparison", "Side-by-side comparison of all selected scenarios across capital, NPA, liquidity, P&L"),
                ("🔄 Reverse Stress Test", "Identifies break-even shock levels and plots the failure frontier"),
                ("🎓 Education Hub", "Conceptual deep-dives on EVT, Basel, VaR, ES, satellite models and more"),
            ]
            for icon_name, desc in features:
                st.markdown(f"""
                <div style='display:flex;align-items:flex-start;gap:10px;
                            padding:8px 12px;border-left:3px solid {COLORS["gold"]};
                            margin:5px 0;background:{COLORS["cardBg"]};border-radius:0 6px 6px 0;'>
                    <div style='font-weight:700;color:{COLORS["gold"]};min-width:200px;
                                font-size:0.85rem;'>{icon_name}</div>
                    <div style='color:{COLORS["muted"]};font-size:0.83rem;'>{desc}</div>
                </div>
                """, unsafe_allow_html=True)

        with c2:
            st.markdown(section_header("About the Author", "👨‍🏫"), unsafe_allow_html=True)
            st.markdown(f"""
            <div style='background:linear-gradient(135deg,{COLORS["darkblue"]},{COLORS["cardBg"]});
                        border:2px solid {COLORS["gold"]};border-radius:10px;padding:28px 24px;
                        text-align:center;'>
                <div style='font-family:Playfair Display,serif;font-size:1.5rem;
                            font-weight:900;color:{COLORS["gold"]};'>
                    Prof. V. Ravichandran
                </div>
                <div style='color:{COLORS["lightblue"]};font-size:0.88rem;
                            margin:6px 0 16px 0;letter-spacing:0.5px;'>
                    28+ Years Corporate Finance & Banking<br>
                    10+ Years Academic Excellence
                </div>
                <hr style='border-color:{COLORS["gold"]}44;margin:12px 0;'>
                <div style='text-align:left;color:{COLORS["text"]};
                            font-size:0.85rem;line-height:1.8;'>
                    <b style='color:{COLORS["gold"]};'>Visiting Faculty at:</b><br>
                    • BITS Pilani (WILP)<br>
                    • Christ University, Bangalore<br>
                    • Goa Institute of Management<br>
                    • ICFAI Centre for Higher Education, Bangalore
                    <br><br>
                    <b style='color:{COLORS["gold"]};'>Courses Taught:</b><br>
                    • Financial Risk Management<br>
                    • Fixed Income Securities & Analysis<br>
                    • Financial Derivatives<br>
                    • Investment Banking<br>
                    • Alternative Investment Markets<br>
                    • Value Risk & Capital Markets
                    <br><br>
                    <b style='color:{COLORS["gold"]};'>Target Audience:</b><br>
                    MBA · CFA · FRM Students & Practitioners
                </div>
                <hr style='border-color:{COLORS["gold"]}44;margin:16px 0 12px 0;'>
                <a href='https://www.linkedin.com/in/trichyravis' target='_blank'
                   style='color:{COLORS["gold"]};font-weight:700;
                          text-decoration:none;font-size:0.85rem;'>
                    🔗 LinkedIn Profile
                </a>
                &nbsp;&nbsp;
                <a href='https://github.com/trichyravis' target='_blank'
                   style='color:{COLORS["gold"]};font-weight:700;
                          text-decoration:none;font-size:0.85rem;'>
                    💻 GitHub
                </a>
            </div>
            """, unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown(section_header("Tech Stack", "🔧"), unsafe_allow_html=True)
            tech = [
                ("Python 3.10+", "Core language", "#3776ab"),
                ("Streamlit", "App framework", "#ff4b4b"),
                ("Plotly", "Interactive charts", "#3d4db7"),
                ("Pandas / NumPy", "Data & numerics", "#150458"),
                ("SciPy", "Statistical models", "#8caae6"),
            ]
            for lib, role, col in tech:
                st.markdown(f"""
                <div style='display:flex;justify-content:space-between;align-items:center;
                            padding:7px 14px;background:{COLORS["cardBg"]};
                            border-left:3px solid {col};
                            border-radius:0 6px 6px 0;margin:4px 0;'>
                    <span style='color:{COLORS["text"]};font-weight:700;
                                 font-size:0.84rem;'>{lib}</span>
                    <span style='color:{COLORS["muted"]};font-size:0.80rem;'>{role}</span>
                </div>
                """, unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown(section_header("Regulatory Alignment", "📜"), unsafe_allow_html=True)
            regs = [
                ("RBI Stress Testing Guidelines", "Master Circular 2015 + updates"),
                ("Basel III / FRTB", "CET1, ES@97.5%, CRAR floors"),
                ("ICAAP (Pillar 2)", "Internal capital adequacy"),
                ("LCR / NSFR", "Basel III liquidity standards"),
                ("DFAST / EBA methodology", "Scenario design best practice"),
            ]
            for reg, desc in regs:
                st.markdown(f"""
                <div style='padding:6px 12px;background:{COLORS["cardBg"]};
                            border-left:3px solid {COLORS["lightblue"]};
                            border-radius:0 6px 6px 0;margin:4px 0;'>
                    <span style='color:{COLORS["gold"]};font-weight:700;
                                 font-size:0.83rem;'>{reg}</span>
                    <span style='color:{COLORS["muted"]};font-size:0.79rem;'>
                        &nbsp;—&nbsp;{desc}</span>
                </div>
                """, unsafe_allow_html=True)

        # Disclaimer
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(f"""
        <div style='background:#1a0a0a;border:1px solid #cc3333;border-radius:8px;
                    padding:14px 20px;color:{COLORS["muted"]};font-size:0.78rem;'>
            <b style='color:#ff6666;'>⚠️ Disclaimer:</b> This platform uses entirely
            synthetic data for educational and research purposes. No real bank data,
            confidential information, or proprietary models are used. All stress test
            results are illustrative. Nothing on this platform constitutes financial,
            regulatory, or investment advice.
        </div>
        """, unsafe_allow_html=True)

    # ═══════════════════════════════════════════════════════════════
    # PAGE 11: EDUCATION HUB
    # ═══════════════════════════════════════════════════════════════
    elif page == "🎓 Education Hub":
        st.markdown(f"""
        <div style='background:linear-gradient(135deg,{COLORS["darkblue"]},{COLORS["cardBg"]});
                    border:2px solid {COLORS["gold"]};border-radius:10px;
                    padding:24px 32px;margin-bottom:20px;'>
            <div style='font-family:Playfair Display,serif;font-size:1.8rem;
                        font-weight:900;color:{COLORS["gold"]};'>
                🎓 Stress Testing Education Hub
            </div>
            <div style='color:{COLORS["lightblue"]};font-size:0.92rem;margin-top:6px;'>
                Conceptual foundations, regulatory frameworks, and modelling methodologies
                for bank stress testing — curated for MBA, CFA & FRM students.
            </div>
        </div>
        """, unsafe_allow_html=True)

        edu_tab1, edu_tab2, edu_tab3, edu_tab4, edu_tab5, edu_tab6 = st.tabs([
            "📖 Foundations",
            "🏛️ Regulatory Framework",
            "🔬 Modelling Approaches",
            "📐 Risk Measures",
            "🌍 Historical Crises",
            "📚 Key References",
        ])

        # ── TAB 1: FOUNDATIONS ──────────────────────────────────
        with edu_tab1:
            st.markdown(section_header("What is Bank Stress Testing?", "📖"), unsafe_allow_html=True)
            st.markdown(f"""
            <div class='metric-card' style='line-height:1.9;margin-bottom:14px;'>
                <b style='color:{COLORS["gold"]};font-size:1.0rem;'>Definition</b><br>
                <span style='color:{COLORS["text"]};'>
                Bank stress testing is the process of subjecting a bank's balance sheet,
                income statement, capital adequacy ratios, and liquidity buffers to
                hypothetical but <i>plausible adverse</i> macroeconomic and market scenarios
                to assess whether the institution can remain solvent and liquid through
                the stress period. It answers the fundamental question:
                <b style='color:{COLORS["gold"]};'>"How bad can it get — and can we survive it?"</b>
                </span>
            </div>
            """, unsafe_allow_html=True)

            c1, c2 = st.columns(2)
            with c1:
                st.markdown(f"""
                <div class='metric-card'>
                    <b style='color:{COLORS["gold"]};'>🎯 Objectives of Stress Testing</b>
                    <ul style='color:{COLORS["text"]};line-height:2.0;margin-top:8px;'>
                        <li>Identify vulnerabilities in the balance sheet before they materialise</li>
                        <li>Quantify capital and liquidity gaps under adverse conditions</li>
                        <li>Inform risk appetite and concentration limits</li>
                        <li>Satisfy regulatory capital adequacy requirements (ICAAP)</li>
                        <li>Support strategic planning and contingency funding</li>
                        <li>Enable board-level risk oversight and governance</li>
                    </ul>
                </div>
                """, unsafe_allow_html=True)
            with c2:
                st.markdown(f"""
                <div class='metric-card'>
                    <b style='color:{COLORS["gold"]};'>🏗️ The Four-Layer Architecture</b>
                    <div style='margin-top:10px;'>
                """, unsafe_allow_html=True)
                layers = [
                    ("Layer 1", "Macro Scenario Engine", "GDP, rates, FX, equity path generation"),
                    ("Layer 2", "Satellite Models", "Macro → credit losses, NII, fee income"),
                    ("Layer 3", "P&L & B/S Projection", "Integrated financial statement projection"),
                    ("Layer 4", "Capital & Liquidity", "CET1, CRAR, LCR, NSFR breach detection"),
                ]
                for num, name, desc in layers:
                    st.markdown(f"""
                    <div style='display:flex;align-items:center;gap:10px;
                                padding:7px 10px;margin:4px 0;
                                background:{COLORS["darkblue"]};border-radius:6px;'>
                        <span style='background:{COLORS["gold"]};color:{COLORS["darkblue"]};
                                     font-weight:900;font-size:0.72rem;border-radius:3px;
                                     padding:2px 7px;min-width:56px;text-align:center;'>{num}</span>
                        <div>
                            <div style='color:{COLORS["gold"]};font-weight:700;
                                        font-size:0.83rem;'>{name}</div>
                            <div style='color:{COLORS["muted"]};font-size:0.76rem;'>{desc}</div>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
                st.markdown("</div></div>", unsafe_allow_html=True)

            st.markdown(section_header("Types of Stress Tests", "🔍"), unsafe_allow_html=True)
            types_data = [
                ("Sensitivity Analysis", "Single variable shocked in isolation",
                 "Rate +200bps; NPA ratio doubles",
                 "Identify which variables matter most; risk appetite calibration"),
                ("Historical Scenario", "Replay of past crisis with observed co-movements",
                 "2008 GFC; 2020 COVID; 2013 Taper Tantrum",
                 "Realistic, credible; uses actual correlation structure"),
                ("Hypothetical Scenario", "Forward-looking plausible adverse narrative",
                 "India recession + INR collapse + real estate crash",
                 "Captures novel risks; forward-looking; regulatory stress"),
                ("Reverse Stress Test", "Work backwards from failure to find break-even shocks",
                 "What NPA level breaks CET1 < 8%?",
                 "Identifies Achilles' heel; forces management action planning"),
            ]
            for ttype, method, example, use in types_data:
                st.markdown(f"""
                <div style='background:{COLORS["cardBg"]};border-left:4px solid {COLORS["gold"]};
                            border-radius:0 8px 8px 0;padding:12px 16px;margin:8px 0;'>
                    <div style='font-family:Playfair Display,serif;font-weight:700;
                                color:{COLORS["gold"]};font-size:0.95rem;'>{ttype}</div>
                    <div style='display:grid;grid-template-columns:1fr 1fr 1fr;
                                gap:8px;margin-top:8px;'>
                        <div><span style='color:{COLORS["muted"]};font-size:0.75rem;
                                         text-transform:uppercase;'>Method</span><br>
                             <span style='color:{COLORS["text"]};font-size:0.82rem;'>{method}</span></div>
                        <div><span style='color:{COLORS["muted"]};font-size:0.75rem;
                                         text-transform:uppercase;'>Example</span><br>
                             <span style='color:{COLORS["lightblue"]};font-size:0.82rem;'>{example}</span></div>
                        <div><span style='color:{COLORS["muted"]};font-size:0.75rem;
                                         text-transform:uppercase;'>Use Case</span><br>
                             <span style='color:{COLORS["text"]};font-size:0.82rem;'>{use}</span></div>
                    </div>
                </div>
                """, unsafe_allow_html=True)

        # ── TAB 2: REGULATORY FRAMEWORK ─────────────────────────
        with edu_tab2:
            st.markdown(section_header("Basel III / IV Capital Requirements", "🏛️"), unsafe_allow_html=True)
            c1, c2 = st.columns(2)
            with c1:
                capital_reqs = [
                    ("CET1 Ratio (Minimum)", "4.5%", "6.0% (RBI)"),
                    ("CET1 + Conservation Buffer", "7.0%", "8.0% (RBI)"),
                    ("Tier 1 Ratio (Minimum)", "6.0%", "7.5% (RBI)"),
                    ("Total CRAR (Minimum)", "8.0%", "11.5% (RBI)"),
                    ("Capital Conservation Buffer", "2.5%", "2.5%"),
                    ("Countercyclical Buffer", "0–2.5%", "0% (currently)"),
                    ("D-SIB Surcharge (SBI)", "+0.6%", "Additional"),
                    ("Leverage Ratio", "3.0%", "3.5% (RBI)"),
                ]
                st.markdown(f"""
                <table class='styled-table'>
                    <thead><tr>
                        <th>Capital Metric</th>
                        <th>Basel III Min</th>
                        <th>RBI Requirement</th>
                    </tr></thead>
                    <tbody>
                        {"".join(f"<tr><td>{r[0]}</td><td>{r[1]}</td><td style='color:{COLORS['gold']};font-weight:700;'>{r[2]}</td></tr>" for r in capital_reqs)}
                    </tbody>
                </table>
                """, unsafe_allow_html=True)
            with c2:
                st.markdown(f"""
                <div class='metric-card'>
                    <b style='color:{COLORS["gold"]};'>📋 FRTB — Fundamental Review of the Trading Book</b>
                    <div style='color:{COLORS["text"]};line-height:1.9;margin-top:10px;font-size:0.85rem;'>
                        <b style='color:{COLORS["lightblue"]};'>Key Change:</b>
                        Replaces 99% VaR with <b>97.5% Expected Shortfall (ES)</b> as the
                        primary market risk metric.<br><br>
                        <b style='color:{COLORS["lightblue"]};'>Why ES over VaR?</b>
                        ES captures the <i>average loss beyond the VaR threshold</i>,
                        penalising fat-tailed distributions that VaR ignores.
                        ES is a <b>coherent risk measure</b> (subadditive); VaR is not.<br><br>
                        <b style='color:{COLORS["lightblue"]};'>Liquidity Horizons:</b>
                        Different holding periods (10–120 days) assigned by asset class —
                        equities 20 days, credit products 40–120 days.<br><br>
                        <b style='color:{COLORS["lightblue"]};'>Backtesting:</b>
                        P&L attribution tests required; IMA approval withdrawn if
                        excessive P&L exceptions detected.
                    </div>
                </div>
                """, unsafe_allow_html=True)

            st.markdown(section_header("RBI Stress Testing Guidelines (India)", "🇮🇳"), unsafe_allow_html=True)
            rbi_items = [
                ("Scope", "All scheduled commercial banks with assets > ₹100 crore"),
                ("Frequency", "Annual minimum; quarterly for Systemically Important Banks (D-SIBs)"),
                ("Scenarios", "Minimum 3: mild, moderate, severe. RBI publishes benchmark scenarios annually"),
                ("Risk Coverage", "Credit risk, market risk, liquidity risk, and combined scenarios mandatory"),
                ("Governance", "Board Risk Management Committee sign-off; CEO attestation required"),
                ("ICAAP Integration", "Stress test results feed directly into Pillar 2 capital planning"),
                ("Reporting", "Results submitted to RBI annually; published in Financial Stability Report"),
                ("NPR Mechanism", "RBI can mandate capital addition if stress tests reveal inadequacy"),
            ]
            for item, desc in rbi_items:
                st.markdown(f"""
                <div style='display:flex;gap:12px;padding:7px 12px;
                            background:{COLORS["cardBg"]};border-radius:6px;margin:4px 0;
                            border-left:3px solid {COLORS["midblue"]};'>
                    <span style='color:{COLORS["gold"]};font-weight:700;min-width:160px;
                                 font-size:0.83rem;'>{item}</span>
                    <span style='color:{COLORS["text"]};font-size:0.83rem;'>{desc}</span>
                </div>
                """, unsafe_allow_html=True)

            st.markdown(section_header("Liquidity Standards: LCR & NSFR", "💧"), unsafe_allow_html=True)
            c3, c4 = st.columns(2)
            with c3:
                st.markdown(f"""
                <div class='metric-card'>
                    <b style='color:{COLORS["gold"]};'>LCR — Liquidity Coverage Ratio</b>
                    <div style='font-size:1.4rem;color:{COLORS["lightblue"]};
                                font-family:Playfair Display,serif;margin:8px 0;'>
                        LCR = HQLA / Net Cash Outflows (30-day) ≥ 100%
                    </div>
                    <div style='color:{COLORS["text"]};font-size:0.83rem;line-height:1.8;'>
                        <b>HQLA Tiers:</b><br>
                        • Level 1: Cash, central bank reserves, sovereign bonds (0% haircut)<br>
                        • Level 2A: High-grade corporate bonds (15% haircut)<br>
                        • Level 2B: RMBS, lower-rated corporates (25–50% haircut)<br><br>
                        <b>Stress Outflow Rates:</b><br>
                        • Retail deposits (stable): 3–5% run-off<br>
                        • Retail deposits (less stable): 10% run-off<br>
                        • Wholesale (non-financial): 25–40% run-off<br>
                        • Interbank / financial: 100% run-off
                    </div>
                </div>
                """, unsafe_allow_html=True)
            with c4:
                st.markdown(f"""
                <div class='metric-card'>
                    <b style='color:{COLORS["gold"]};'>NSFR — Net Stable Funding Ratio</b>
                    <div style='font-size:1.4rem;color:{COLORS["lightblue"]};
                                font-family:Playfair Display,serif;margin:8px 0;'>
                        NSFR = Available Stable Funding / Required Stable Funding ≥ 100%
                    </div>
                    <div style='color:{COLORS["text"]};font-size:0.83rem;line-height:1.8;'>
                        <b>Available Stable Funding (ASF):</b><br>
                        • Equity & Tier 1 capital: 100% weight<br>
                        • Retail deposits (>1yr): 95% weight<br>
                        • Wholesale funding (>1yr): 50–100% weight<br><br>
                        <b>Required Stable Funding (RSF):</b><br>
                        • Cash & short-term assets: 0–5% weight<br>
                        • Loans to corporates (<1yr): 50% weight<br>
                        • Mortgages: 65% weight<br>
                        • Non-performing loans: 100% weight
                    </div>
                </div>
                """, unsafe_allow_html=True)

        # ── TAB 3: MODELLING APPROACHES ─────────────────────────
        with edu_tab3:
            st.markdown(section_header("Satellite Models — The Technical Heart", "🔬"), unsafe_allow_html=True)
            st.markdown(f"""
            <div class='metric-card' style='margin-bottom:16px;'>
                <b style='color:{COLORS["gold"]};'>What are Satellite Models?</b><br>
                <span style='color:{COLORS["text"]};line-height:1.9;'>
                Satellite models are econometric or statistical models that translate
                macroeconomic shock variables (GDP growth, interest rates, unemployment,
                exchange rates) into bank-specific financial outcomes (NPA ratios,
                net interest income, fee income, trading losses). Each major risk type
                requires its own satellite model.
                </span>
            </div>
            """, unsafe_allow_html=True)

            models = [
                {
                    "name": "📉 Credit Loss Model (Most Critical)",
                    "formula": "ΔNPA_t = α + β₁·ΔGDP_t + β₂·Δrate_t + β₃·Δunemployment_t + β₄·NPA_{t-1} + β₅·Δproperty_t + ε_t",
                    "inputs": "GDP growth, repo rate, unemployment, property prices, sector output",
                    "output": "Quarterly NPA ratio by portfolio segment → provisions → capital depletion",
                    "estimation": "OLS regression on 10–15 years of historical quarterly data; estimated separately for retail, corporate, MSME, agriculture",
                    "key_param": "GDP sensitivity coefficient β₁ typically −2.0 to −3.5 (1% GDP fall → 2–3.5% relative NPA increase)",
                },
                {
                    "name": "📈 Net Interest Income (NII) Model",
                    "formula": "ΔNII = Σ (Repricing Gap_k) × (Δrate_k) across maturity buckets k",
                    "inputs": "Yield curve shift, repricing schedule of assets and liabilities, CASA ratio, loan mix",
                    "output": "NII compression or expansion under rate shock scenarios",
                    "estimation": "Asset-liability management gap analysis; duration mismatch quantification",
                    "key_param": "Asset-liability repricing gap (rate-sensitive assets minus liabilities); CASA deposits are sticky, wholesale reprices fast",
                },
                {
                    "name": "💹 Market Risk / MTM Model",
                    "formula": "ΔP ≈ −Duration × Δyield × P_AFS + β_equity × ΔEquity_index + ΔFXNOP × ΔRate",
                    "inputs": "Modified duration of AFS portfolio, equity beta, net open FX position, credit spread widening",
                    "output": "Mark-to-market losses on AFS investments, equity portfolio, and FX positions",
                    "estimation": "Duration from bond analytics; equity beta from regression; FX exposure from treasury systems",
                    "key_param": "Average modified duration of AFS portfolio (typically 3–6 years for Indian banks)",
                },
                {
                    "name": "💰 Fee & Non-Interest Income Model",
                    "formula": "Fee_t = α + β₁·GDP_t + β₂·CreditGrowth_t + β₃·Nifty_t + ε_t",
                    "inputs": "Economic activity proxies, transaction volume indicators, equity market levels",
                    "output": "Fee income compression during recession (transaction banking, trade finance, wealth management)",
                    "estimation": "Regression of fee income on macroeconomic and market activity proxies",
                    "key_param": "Fee income typically falls 15–35% under severe stress as transaction volumes, credit card spends, and trade finance volumes contract",
                },
            ]
            for m in models:
                _cb  = COLORS["cardBg"]
                _mb  = COLORS["midblue"]
                _g   = COLORS["gold"]
                _db  = COLORS["darkblue"]
                _bd  = COLORS["bgDark"]
                _lb  = COLORS["lightblue"]
                _grid_sep = hex_to_rgba(COLORS["midblue"], 0.27)
                _html = (
                    "<div style='background:" + _cb + ";border:1px solid " + _mb + ";"
                    "border-left:5px solid " + _g + ";border-radius:8px;margin:12px 0;overflow:hidden;'>"

                    "<div style='background:linear-gradient(135deg," + _db + "," + _cb + ");"
                    "padding:13px 18px;border-bottom:1px solid " + _mb + ";'>"
                    "<span style='font-family:Playfair Display,serif;font-size:1.0rem;"
                    "font-weight:700;color:#ffffff;letter-spacing:0.3px;'>"
                    + m["name"] +
                    "</span></div>"

                    "<div style='background:" + _db + ";padding:11px 18px;"
                    "border-bottom:1px solid " + _mb + ";"
                    "font-family:monospace;font-size:0.86rem;'>"
                    "<span style='color:" + _g + ";font-weight:700;'>Model Equation:&nbsp;</span>"
                    "<span style='color:" + _lb + ";'>" + m["formula"] + "</span>"
                    "</div>"

                    "<div style='display:grid;grid-template-columns:1fr 1fr;gap:1px;"
                    "background:" + _grid_sep + ";'>"

                    "<div style='background:" + _bd + ";padding:13px 16px;'>"
                    "<div style='color:" + _g + ";font-size:0.72rem;font-weight:700;"
                    "text-transform:uppercase;letter-spacing:1px;margin-bottom:6px;'>📥 Inputs</div>"
                    "<div style='color:#d0dff0;font-size:0.84rem;line-height:1.7;'>" + m["inputs"] + "</div>"
                    "</div>"

                    "<div style='background:" + _bd + ";padding:13px 16px;'>"
                    "<div style='color:" + _g + ";font-size:0.72rem;font-weight:700;"
                    "text-transform:uppercase;letter-spacing:1px;margin-bottom:6px;'>📤 Output</div>"
                    "<div style='color:#d0dff0;font-size:0.84rem;line-height:1.7;'>" + m["output"] + "</div>"
                    "</div>"

                    "<div style='background:" + _cb + ";padding:13px 16px;'>"
                    "<div style='color:" + _g + ";font-size:0.72rem;font-weight:700;"
                    "text-transform:uppercase;letter-spacing:1px;margin-bottom:6px;'>⚙️ Estimation</div>"
                    "<div style='color:#d0dff0;font-size:0.84rem;line-height:1.7;'>" + m["estimation"] + "</div>"
                    "</div>"

                    "<div style='background:" + _cb + ";padding:13px 16px;"
                    "border-left:2px solid " + _g + ";'>"
                    "<div style='color:" + _g + ";font-size:0.72rem;font-weight:700;"
                    "text-transform:uppercase;letter-spacing:1px;margin-bottom:6px;'>🔑 Key Parameter</div>"
                    "<div style='color:#ffffff;font-size:0.84rem;line-height:1.7;font-weight:500;'>" + m["key_param"] + "</div>"
                    "</div>"

                    "</div></div>"
                )
                st.markdown(_html, unsafe_allow_html=True)

            st.markdown(section_header("P&L and Capital Projection Logic", "🧮"), unsafe_allow_html=True)
            st.markdown(f"""
            <div style='background:{COLORS["darkblue"]};border-radius:8px;
                        padding:16px 20px;font-family:monospace;
                        color:{COLORS["lightblue"]};font-size:0.85rem;
                        border:1px solid {COLORS["midblue"]};line-height:2.0;'>
                <b style='color:{COLORS["gold"]};'>Quarterly Capital Projection Algorithm:</b><br><br>
                Pre-Provision Operating Profit (PPOP)<br>
                &nbsp;&nbsp;&nbsp;&nbsp;= Stressed NII + Stressed Fee Income + Stressed Trading Income<br>
                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;− Operating Costs (relatively fixed)<br><br>
                Net Profit = PPOP − Incremental Provisions − Taxes<br><br>
                Retained Earnings += Net Profit<br>
                CET1 Capital += Retained Earnings (after dividend)<br>
                RWA = RWA × (1 + credit migration factor + growth factor)<br><br>
                <b style='color:{COLORS["gold"]};'>CET1 Ratio = CET1 Capital / Risk-Weighted Assets</b><br><br>
                <span style='color:{COLORS["muted"]};'>Breach check each quarter: CET1 &lt; 8.0% → regulatory intervention threshold</span>
            </div>
            """, unsafe_allow_html=True)

        # ── TAB 4: RISK MEASURES ────────────────────────────────
        with edu_tab4:
            st.markdown(section_header("Value at Risk (VaR) vs Expected Shortfall (ES)", "📐"), unsafe_allow_html=True)
            c1, c2 = st.columns(2)
            with c1:
                st.markdown(f"""
                <div class='metric-card'>
                    <b style='color:{COLORS["gold"]};font-size:0.95rem;'>
                        📊 Value at Risk (VaR)
                    </b>
                    <div style='font-size:1.5rem;color:{COLORS["lightblue"]};
                                font-family:Playfair Display,serif;margin:10px 0;'>
                        VaR_q = inf{{x : P(L > x) ≤ 1−q}}
                    </div>
                    <div style='color:{COLORS["text"]};font-size:0.83rem;line-height:1.9;'>
                        <b>Definition:</b> The maximum loss not exceeded with probability q
                        over a given holding period.<br>
                        <b>Intuition:</b> "We are 99% confident our daily loss will not exceed
                        ₹X crore."<br>
                        <b>Limitations:</b>
                        <ul style='margin:4px 0;'>
                            <li>Not <i>subadditive</i> — not a coherent risk measure</li>
                            <li>Says nothing about <i>how bad</i> losses are beyond the threshold</li>
                            <li>Underestimates tail risk for fat-tailed distributions</li>
                            <li>Susceptible to manipulation by concentrated positions</li>
                        </ul>
                        <b>Regulatory status:</b> Replaced by ES under FRTB (Basel IV)
                        for internal model approval.
                    </div>
                </div>
                """, unsafe_allow_html=True)
            with c2:
                st.markdown(f"""
                <div class='metric-card'>
                    <b style='color:{COLORS["gold"]};font-size:0.95rem;'>
                        📊 Expected Shortfall (ES) / CVaR
                    </b>
                    <div style='font-size:1.5rem;color:{COLORS["lightblue"]};
                                font-family:Playfair Display,serif;margin:10px 0;'>
                        ES_q = E[L | L > VaR_q]
                    </div>
                    <div style='color:{COLORS["text"]};font-size:0.83rem;line-height:1.9;'>
                        <b>Definition:</b> The expected loss <i>given that</i> the loss
                        exceeds VaR — the average of the worst (1−q)% outcomes.<br>
                        <b>Intuition:</b> "Given that we breach the 99% threshold, our
                        average loss will be ₹Y crore."<br>
                        <b>Advantages:</b>
                        <ul style='margin:4px 0;'>
                            <li><b>Coherent risk measure</b> — satisfies subadditivity</li>
                            <li>Captures the full severity of tail losses</li>
                            <li>Penalises fat tails that VaR ignores</li>
                            <li>ES ≥ VaR always; the gap grows with tail heaviness</li>
                        </ul>
                        <b>Regulatory status:</b> Mandated at <b>97.5%</b> confidence
                        under FRTB for market risk capital.
                    </div>
                </div>
                """, unsafe_allow_html=True)

            st.markdown(section_header("Key Capital Adequacy Ratios", "🏦"), unsafe_allow_html=True)
            ratios = [
                ("CET1 Ratio", "Common Equity Tier 1 / Risk-Weighted Assets",
                 "The highest quality capital — ordinary shares + retained earnings only. "
                 "The core loss-absorbing layer. RBI minimum: 6% (+ 2.5% conservation buffer = 8.5%)."),
                ("Tier 1 Ratio", "Tier 1 Capital / Risk-Weighted Assets",
                 "CET1 + Additional Tier 1 (AT1 bonds, perpetual debt). "
                 "AT1 instruments absorb losses through coupon cancellation or principal write-down. "
                 "RBI minimum: 7.5%."),
                ("CRAR (Total Capital)", "Total Capital / Risk-Weighted Assets",
                 "Tier 1 + Tier 2 (sub-debt, general provisions). "
                 "RBI minimum: 11.5% including all buffers. Breach triggers PCA (Prompt Corrective Action)."),
                ("Leverage Ratio", "Tier 1 Capital / Total Exposure (on + off balance sheet)",
                 "Non-risk-based backstop to prevent excessive leverage regardless of RWA model. "
                 "Basel III minimum 3%; RBI requires 3.5%. Prevents model arbitrage."),
                ("PCR (Provision Coverage Ratio)", "Cumulative Provisions / Gross NPA",
                 "Measures adequacy of provisions against NPAs. "
                 "Higher PCR = more conservative; reduces future P&L volatility. "
                 "RBI advisory minimum: 70% (was mandatory 2009–2011)."),
            ]
            for name, formula, description in ratios:
                _cb = COLORS["cardBg"]; _mb = COLORS["midblue"]
                _g  = COLORS["gold"];   _db = COLORS["darkblue"]
                _bd = COLORS["bgDark"]; _lb = COLORS["lightblue"]
                _html = (
                    "<div style='background:" + _cb + ";border:1px solid " + _mb + ";"
                    "border-left:5px solid " + _g + ";border-radius:8px;margin:8px 0;overflow:hidden;'>"
                    "<div style='background:linear-gradient(135deg," + _db + "," + _cb + ");"
                    "padding:11px 18px;border-bottom:1px solid " + _mb + ";'>"
                    "<span style='font-family:Playfair Display,serif;font-size:0.95rem;"
                    "font-weight:700;color:#ffffff;'>📐 " + name + "</span></div>"
                    "<div style='background:" + _db + ";padding:10px 18px;"
                    "border-bottom:1px solid " + _mb + ";"
                    "font-family:monospace;font-size:0.88rem;'>"
                    "<span style='color:" + _g + ";font-weight:700;'>Formula:&nbsp;</span>"
                    "<span style='color:" + _lb + ";'>" + formula + "</span></div>"
                    "<div style='padding:12px 18px;background:" + _bd + ";'>"
                    "<div style='color:#d0dff0;font-size:0.84rem;line-height:1.8;'>" + description + "</div>"
                    "</div></div>"
                )
                st.markdown(_html, unsafe_allow_html=True)

            st.markdown(section_header("Gross NPA vs Net NPA vs SMA", "📉"), unsafe_allow_html=True)
            npa_stages = [
                ("Standard Asset", "0–30 DPD", "0.25–1%", "Normal performing; no concern"),
                ("SMA-0", "0 DPD but stress signals", "—", "Overdue principal/interest 1–30 days"),
                ("SMA-1", "31–60 DPD", "—", "Early warning; enhanced monitoring"),
                ("SMA-2", "61–90 DPD", "—", "Close watch; pre-NPA stage"),
                ("Sub-Standard (NPA)", "91–365 DPD", "15–25%", "Declared NPA; provisioning begins"),
                ("Doubtful (NPA)", ">1 year NPA", "25–100%", "Secured: 25→40→100%; Unsecured: 100%"),
                ("Loss (NPA)", "Identified loss", "100%", "Written off or fully provisioned"),
            ]
            st.markdown(f"""
            <table class='styled-table'>
                <thead><tr>
                    <th>Asset Classification</th><th>Days Past Due</th>
                    <th>Provision Rate</th><th>Description</th>
                </tr></thead>
                <tbody>
                    {"".join(f"<tr><td style='color:{COLORS["gold"]};font-weight:700;'>{r[0]}</td><td>{r[1]}</td><td style='color:{COLORS["red"] if "NPA" in r[0] or "Loss" in r[0] else COLORS["green"]};font-weight:700;'>{r[2]}</td><td>{r[3]}</td></tr>" for r in npa_stages)}
                </tbody>
            </table>
            """, unsafe_allow_html=True)

        # ── TAB 5: HISTORICAL CRISES ─────────────────────────────
        with edu_tab5:
            st.markdown(section_header("Historical Stress Episodes — India & Global", "🌍"), unsafe_allow_html=True)
            crises = [
                {
                    "name": "2008 Global Financial Crisis (GFC)",
                    "period": "Sep 2008 – Mar 2009",
                    "trigger": "US subprime mortgage collapse → Lehman Brothers bankruptcy → global credit freeze",
                    "india_impact": "Nifty −60% peak-to-trough; GDP growth slowed to 6.7%; FII outflows $13bn; INR depreciated 25%",
                    "banking": "Indian banks relatively insulated (low US exposure); credit growth crashed from 30% to 17%; liquidity stress",
                    "lesson": "Contagion through capital markets and trade finance channels even with limited direct exposure; CASA banks outperformed",
                    "color": COLORS["red"],
                },
                {
                    "name": "2013 Taper Tantrum",
                    "period": "May – Sep 2013",
                    "trigger": "Fed signals tapering of QE; sudden EM capital outflows; INR depreciation spiral",
                    "india_impact": "INR depreciated 20% (₹55 to ₹68); 10Y G-sec yield +200bps; Nifty −10%; RBI emergency rate hike 50bps",
                    "banking": "MTM losses on AFS bond portfolios; NII compression; wholesale funding costs spiked",
                    "lesson": "Duration mismatch in investment portfolios is a severe vulnerability; CASA funding outperforms under rate shock",
                    "color": COLORS["orange"],
                },
                {
                    "name": "2018 IL&FS / NBFC Crisis",
                    "period": "Sep 2018 – Mar 2019",
                    "trigger": "IL&FS defaults on commercial paper; contagion to entire NBFC sector; funding freeze",
                    "india_impact": "Nifty −15%; NBFC stocks −40–80%; mutual fund redemptions; CP market paralysed",
                    "banking": "Banks with high NBFC/HFC exposure suffered NPA spike; inter-bank funding stress; liquidity hoarding",
                    "lesson": "Concentration in NBFC sector creates tail risk; wholesale funding dependency is acute vulnerability; liquidity crisis ≠ solvency crisis",
                    "color": COLORS["yellow"],
                },
                {
                    "name": "2020 COVID-19 Shock",
                    "period": "Feb – Apr 2020",
                    "trigger": "Pandemic lockdowns; global demand collapse; RBI moratorium on loan repayments",
                    "india_impact": "GDP −7.3% (FY21); Nifty −38% in 6 weeks; RBI cut repo 115bps; ₹20 lakh crore stimulus package",
                    "banking": "Moratorium covered 50%+ of loan books; deferred NPA recognition; provisioning surged; credit growth near zero",
                    "lesson": "Regulatory forbearance (moratorium) can delay but not prevent NPA crystallisation; banks with high retail floating rate loans benefited from rate cuts",
                    "color": "#cc66ff",
                },
                {
                    "name": "2023 Adani Short-Seller Report",
                    "period": "Jan – Mar 2023",
                    "trigger": "Hindenburg Research report on Adani Group; sudden market confidence shock",
                    "india_impact": "Adani Group market cap fell $120bn; Nifty −5%; Indian banks with Adani exposure under scrutiny",
                    "banking": "PSU banks with Adani infrastructure exposure flagged; SBI led damage control; LIC holdings scrutinised",
                    "lesson": "Concentrated single-group exposure creates reputational + credit risk simultaneously; board oversight of large borrower exposure critical",
                    "color": COLORS["lightblue"],
                },
            ]
            for crisis in crises:
                _cc  = crisis["color"]
                _cb  = COLORS["cardBg"]
                _db  = COLORS["darkblue"]
                _bd  = COLORS["bgDark"]
                _g   = COLORS["gold"]
                _bdr = hex_to_rgba(_cc, 0.13)   # border separator
                _hdr = hex_to_rgba(_cc, 0.27)   # grid sep
                _badge_bg  = hex_to_rgba(_cc, 0.13)
                _badge_bdr = hex_to_rgba(_cc, 0.40)
                _html = (
                    "<div style='background:" + _cb + ";border:1px solid " + _badge_bdr + ";"
                    "border-left:5px solid " + _cc + ";border-radius:8px;margin:10px 0;overflow:hidden;'>"

                    "<div style='background:linear-gradient(135deg," + _db + "," + _cb + ");"
                    "padding:12px 18px;border-bottom:1px solid " + _bdr + ";"
                    "display:flex;justify-content:space-between;align-items:center;'>"
                    "<span style='font-family:Playfair Display,serif;font-size:1.0rem;"
                    "font-weight:700;color:#ffffff;letter-spacing:0.3px;'>📅 " + crisis["name"] + "</span>"
                    "<span style='background:" + _badge_bg + ";color:" + _cc + ";"
                    "border:1px solid " + _badge_bdr + ";border-radius:20px;"
                    "padding:3px 12px;font-size:0.75rem;font-weight:700;white-space:nowrap;'>"
                    + crisis["period"] + "</span></div>"

                    "<div style='display:grid;grid-template-columns:1fr 1fr;"
                    "gap:1px;background:" + _hdr + ";'>"

                    "<div style='background:" + _bd + ";padding:14px 16px;'>"
                    "<div style='color:" + _g + ";font-size:0.75rem;font-weight:700;"
                    "text-transform:uppercase;letter-spacing:1px;margin-bottom:6px;'>🔥 Trigger</div>"
                    "<div style='color:#d0dff0;font-size:0.85rem;line-height:1.7;'>" + crisis["trigger"] + "</div>"
                    "</div>"

                    "<div style='background:" + _bd + ";padding:14px 16px;'>"
                    "<div style='color:" + _g + ";font-size:0.75rem;font-weight:700;"
                    "text-transform:uppercase;letter-spacing:1px;margin-bottom:6px;'>🇮🇳 India Impact</div>"
                    "<div style='color:#d0dff0;font-size:0.85rem;line-height:1.7;'>" + crisis["india_impact"] + "</div>"
                    "</div>"

                    "<div style='background:" + _cb + ";padding:14px 16px;'>"
                    "<div style='color:" + _g + ";font-size:0.75rem;font-weight:700;"
                    "text-transform:uppercase;letter-spacing:1px;margin-bottom:6px;'>🏦 Banking Sector</div>"
                    "<div style='color:#d0dff0;font-size:0.85rem;line-height:1.7;'>" + crisis["banking"] + "</div>"
                    "</div>"

                    "<div style='background:" + _cb + ";padding:14px 16px;"
                    "border-left:2px solid " + _cc + ";'>"
                    "<div style='color:" + _cc + ";font-size:0.75rem;font-weight:700;"
                    "text-transform:uppercase;letter-spacing:1px;margin-bottom:6px;'>💡 Key Lesson</div>"
                    "<div style='color:#ffffff;font-size:0.85rem;line-height:1.7;font-weight:500;'>" + crisis["lesson"] + "</div>"
                    "</div>"

                    "</div></div>"
                )
                st.markdown(_html, unsafe_allow_html=True)

        # ── TAB 6: REFERENCES ───────────────────────────────────
        with edu_tab6:
            st.markdown(section_header("Key References & Further Reading", "📚"), unsafe_allow_html=True)
            refs = {
                "📘 Foundational Textbooks": [
                    ("Duffie & Singleton", "Credit Risk: Pricing, Measurement, and Management", "Princeton University Press, 2003"),
                    ("McNeil, Frey & Embrechts", "Quantitative Risk Management (2nd ed.)", "Princeton University Press, 2015"),
                    ("Hull", "Risk Management and Financial Institutions (5th ed.)", "Wiley Finance, 2018"),
                    ("Bessis", "Risk Management in Banking (4th ed.)", "Wiley, 2015"),
                ],
                "🏛️ Regulatory Documents": [
                    ("Basel Committee on Banking Supervision", "Basel III: A Global Regulatory Framework", "BIS, 2010 (rev. 2017)"),
                    ("Basel Committee on Banking Supervision", "Minimum Capital Requirements for Market Risk (FRTB)", "BIS, 2019"),
                    ("Reserve Bank of India", "Master Circular on Stress Testing", "RBI, 2015 + updates"),
                    ("RBI", "Guidelines on Liquidity Risk Management (LCR, NSFR)", "RBI, 2014–2020"),
                    ("RBI", "Prudential Framework for Resolution of Stressed Assets", "RBI, 2019"),
                ],
                "📄 Key Academic Papers": [
                    ("Borio, Drehmann & Tsatsaronis", "Stress-testing macro stress testing: does it live up to expectations?", "BIS Working Paper No. 369, 2012"),
                    ("Quagliariello", "Stress-testing the Banking System: Methodologies and Applications", "Cambridge University Press, 2009"),
                    ("Foglia", "Stress Testing Credit Risk: A Survey of Authorities' Approaches", "International Journal of Central Banking, 2009"),
                    ("Schuermann", "Stress Testing Banks", "Wharton Financial Institutions Center, 2012"),
                ],
                "🔗 Online Resources": [
                    ("BIS", "Basel Committee publications & stress testing papers", "www.bis.org/bcbs"),
                    ("RBI", "Financial Stability Reports (half-yearly)", "www.rbi.org.in"),
                    ("IMF", "Financial Sector Assessment Program (FSAP) methodology", "www.imf.org/fsap"),
                    ("Federal Reserve", "DFAST (Dodd-Frank Stress Tests) documentation", "www.federalreserve.gov/supervisionreg/dfast"),
                ],
            }
            for category, items in refs.items():
                st.markdown(f"""
                <div style='color:{COLORS["gold"]};font-family:Playfair Display,serif;
                            font-weight:700;font-size:1.0rem;margin:16px 0 8px 0;'>
                    {category}
                </div>
                """, unsafe_allow_html=True)
                for i, (author, title, source) in enumerate(items, 1):
                    st.markdown(f"""
                    <div style='display:flex;gap:10px;padding:8px 12px;
                                background:{COLORS["cardBg"]};border-radius:6px;
                                margin:3px 0;border-left:3px solid {COLORS["midblue"]};'>
                        <span style='color:{COLORS["gold"]};font-weight:700;
                                     min-width:22px;font-size:0.82rem;'>[{i}]</span>
                        <div>
                            <span style='color:{COLORS["lightblue"]};font-weight:700;
                                         font-size:0.83rem;'>{author}.</span>
                            <span style='color:{COLORS["text"]};font-size:0.83rem;'>
                                &nbsp;<i>"{title}"</i>
                            </span><br>
                            <span style='color:{COLORS["muted"]};font-size:0.78rem;'>
                                {source}
                            </span>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown(f"""
            <div style='background:linear-gradient(135deg,{COLORS["darkblue"]},{COLORS["cardBg"]});
                        border:1px solid {COLORS["gold"]};border-radius:10px;
                        padding:20px 24px;text-align:center;'>
                <div style='font-family:Playfair Display,serif;font-size:1.1rem;
                            font-weight:700;color:{COLORS["gold"]};margin-bottom:6px;'>
                    THE MOUNTAIN PATH — World of Finance
                </div>
                <div style='color:{COLORS["text"]};font-size:0.84rem;margin-bottom:12px;'>
                    Explore more advanced finance content, models, and educational materials
                </div>
                <a href='https://www.linkedin.com/in/trichyravis' target='_blank'
                   style='color:{COLORS["gold"]};font-weight:700;text-decoration:none;
                          font-size:0.88rem;'>
                    🔗 Connect on LinkedIn
                </a>
                &nbsp;&nbsp;&nbsp;
                <a href='https://github.com/trichyravis' target='_blank'
                   style='color:{COLORS["gold"]};font-weight:700;text-decoration:none;
                          font-size:0.88rem;'>
                    💻 View on GitHub
                </a>
            </div>
            """, unsafe_allow_html=True)


    else:
        st.warning(f"Page not found: `{page}` — please select from the navigation menu.")

    # ── FOOTER ───────────────────────────────────────────────────
    st.markdown("<hr class='gold-divider'>", unsafe_allow_html=True)
    st.markdown(f"""
    <div style='text-align:center; padding:8px 0; color:{COLORS["muted"]}; font-size:0.78rem;'>
        <span style='color:{COLORS["gold"]};font-weight:700;font-family:Playfair Display,serif;'>
            THE MOUNTAIN PATH — World of Finance
        </span><br>
        Bank Stress Testing Lab &nbsp;|&nbsp; Prof. V. Ravichandran &nbsp;|&nbsp;
        28+ Yrs Corporate Finance &amp; Banking &nbsp;|&nbsp;
        <a href='https://www.linkedin.com/in/trichyravis' target='_blank'
           style='color:{COLORS["gold"]};text-decoration:none;'>LinkedIn</a>
        &nbsp;|&nbsp;
        <a href='https://github.com/trichyravis' target='_blank'
           style='color:{COLORS["gold"]};text-decoration:none;'>GitHub</a>
    </div>""", unsafe_allow_html=True)

if __name__ == "__main__":
    main()
