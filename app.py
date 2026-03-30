#!/usr/bin/env python3
"""
INFINITE GZ + AI SMART TECH — Invoice Management System  v3
=============================================================
Changes vs v2:
  • IGZ: Items (JSON) → 4 cols: Product Item / Qty / Unit Price (RM) / Total Amount (RM)
  • IGZ Product Item dropdown: 51 transaction types from IGZ Transaction Reference doc
  • IGZ 2026 market reference prices added alongside each transaction type
  • Both companies: Payment Type & Card Type dropdowns in Excel + UI
  • Screen 1 right column: shows BOTH company price reference tables (IGZ top, AST bottom)
  • Receipt Link & Invoice Link: wide columns + horizontal scroll
  • Invoice No: uses actual receipt invoice number (INV- prefix stripped automatically)
"""

import streamlit as st
import pytesseract
from PIL import Image
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
import re, os, io, base64, datetime, subprocess, json
from pathlib import Path

import sys
sys.path.insert(0, str(Path(__file__).parent))
from invoice_infinitegz_v7 import make_igz
from invoice_ast import make_ast

# ── Paths (with /mnt fallback to local) ────────────────────────────────
def _safe_base() -> Path:
    """Use ./output/ relative to script location — works on any OS."""
    base = Path(__file__).parent / "output"
    base.mkdir(parents=True, exist_ok=True)
    return base

OUT_DIR       = _safe_base()
IGZ_EXCEL     = OUT_DIR / "IGZ_invoices.xlsx"
AST_EXCEL     = OUT_DIR / "AST_invoices.xlsx"
RECEIPTS_DIR  = OUT_DIR / "receipts"
INVOICES_DIR  = OUT_DIR / "invoices"
for d in [RECEIPTS_DIR, INVOICES_DIR]:
    d.mkdir(parents=True, exist_ok=True)

# ══════════════════════════════════════════════════════════════════════
#  INFINITE GZ Product Catalogue  (from IGZ Transaction Reference v1.0)
#  Source: IGZ_Transaction_Reference.docx  |  2026 Malaysia market prices
# ══════════════════════════════════════════════════════════════════════
IGZ_PRODUCTS = [
    # ── GROUP A: Financing Arrangement ──────────────────────────────
    ("[A-001] Card Balance Alignment Package",           1500),
    ("[A-002] Debt Position Adjustment Package",         2000),
    ("[A-003] Credit Profile Care Plan",                  800),
    ("[A-004] Personal Financing Arrangement Pack",      1200),
    ("[A-005] Home Financing Setup Pack",                3000),
    ("[A-006] Business Funding Setup Pack",              5000),
    ("[A-007] Emergency Liquidity Arrangement Pack",     1500),
    ("[A-008] Credit Limit Restructuring Pack",          1200),
    # ── GROUP B: Record & Report Handling ────────────────────────────
    ("[B-001] Monthly Entry Handling Pack",               500),
    ("[B-002] Year-End Report Preparation Pack",         2500),
    ("[B-003] Tax Working File Pack",                    1500),
    ("[B-004] Business Figures Summary Pack",             800),
    ("[B-005] Bank Movement Matching Package",            500),
    ("[B-006] Payroll Processing Pack",                   450),
    ("[B-007] Account Reconciliation Pack",               800),
    ("[B-008] Compliance Documentation Pack",            1500),
    # ── GROUP C: Digital & Business Services ─────────────────────────
    ("[C-001] Website Build & Launch Pack",              5000),
    ("[C-002] Online Reach Boost Package",               3000),
    ("[C-003] Space Layout & Sourcing Pack",             3500),
    ("[C-004] Business Direction Planning Pack",         2500),
    ("[C-005] Document Draft & Filing Pack",              800),
    ("[C-006] Financial Skills Session",                 1200),
    ("[C-007] Brand Identity Setup Pack",                4500),
    ("[C-008] Digital Tools Integration Pack",           3000),
    # ── GROUP D: Outcome Sharing ─────────────────────────────────────
    ("[D-001] Business Setup Integration Pack",          3000),
    ("[D-002] Financing Cost Reduction Outcome",         1000),
    ("[D-003] Extra Income Channel Setup Pack",          2000),
    ("[D-004] Debt Clearing Flow Pack",                  1000),
    ("[D-005] VIP Rotation Access Pack",                 1500),
    ("[D-006] Program Net Outcome Share",                 800),
    ("[D-007] Network Outcome Share",                     800),
    ("[D-008] Product Option Matching Pack",              800),
    ("[D-009] Ongoing Collaboration Outcome Share",       800),
    ("[D-010] Referral Outcome Share",                    500),
    ("[D-011] Agent Payout Share",                        500),
    ("[D-012] POS Terminal MDR Settlement",               200),
    # ── GROUP E: Bank Account Operations ─────────────────────────────
    ("[E-001] Inward Fund Transfer – Bank In",             50),
    ("[E-002] Outward Fund Transfer – Bank Out",           50),
    ("[E-003] IBG Inward – IBG In",                        30),
    ("[E-004] IBG Outward – IBG Out",                      30),
    ("[E-005] Cash Deposit – CASH_DEP",                    50),
    ("[E-006] Cash Withdrawal – CASH_WD",                  50),
    ("[E-007] Cheque Payment – CHEQUE",                    80),
    ("[E-008] Bill Payment – Bill Pay",                    30),
    ("[E-009] Loan Disbursement Received – Loan Disb",    100),
    ("[E-010] Loan Repayment Made – Loan Repay",           80),
    ("[E-011] POS Payment Received – POS Pay",             50),
    ("[E-012] Intercompany Transfer",                     100),
    ("[E-013] MDR Charge Deduction",                       50),
    ("[E-014] Standing Instruction Payment",               50),
    # ── GROUP U: Uncertain / Pending ─────────────────────────────────
    ("[U-001] Pending Review – 待确认",                     0),
]

# ══════════════════════════════════════════════════════════════════════
#  AI SMART TECH Product Catalogue  (2026 Malaysia market prices)
# ══════════════════════════════════════════════════════════════════════
AST_PRODUCTS = [
    ("Website Development – Basic (up to 10 pages)",            3800),
    ("Website Development – Business (custom design)",          8500),
    ("Website Development – E-Commerce",                       18000),
    ("Website Development – Enterprise / Portal",              45000),
    ("Mobile App Development – iOS / Android (basic)",         15000),
    ("Mobile App Development – Custom Full-Stack",             50000),
    ("Software Development – Custom System (SME)",             20000),
    ("Software Development – Enterprise Platform",             80000),
    ("CRM System – Setup & Configuration",                      8000),
    ("CRM System – Custom Development & Integration",          25000),
    ("CRM System – Monthly Support & Maintenance",              1500),
    ("AI Automation – Workflow Bot (single process)",           5000),
    ("AI Automation – Multi-Process Automation Suite",         20000),
    ("AI Automation – Enterprise AI Agent System",             75000),
    ("AI Chatbot – Basic FAQ Bot",                              3500),
    ("AI Chatbot – Advanced NLP Customer Service Bot",         12000),
    ("Hardware Setup – Workstation Configuration (per unit)",    350),
    ("Hardware Setup – Network & Server Infrastructure",        8500),
    ("Hardware Setup – Enterprise IT Infrastructure",          35000),
    ("Brand Strategy & Positioning Package",                    6500),
    ("Brand Identity Design (logo, colours, guidelines)",       4500),
    ("Digital Marketing – Monthly Management (2 platforms)",    4500),
    ("Digital Marketing – Full Campaign Management",            8000),
    ("SEO Optimisation – Monthly Package",                      2500),
    ("Google / Meta Ads Management – Monthly",                  2000),
    ("Social Media Content Creation – Monthly (10 posts)",      2200),
    ("Technical Support – Monthly Retainer (8×5)",              2500),
    ("Technical Support – Monthly Retainer (24×7)",             5000),
    ("Technical Support – Per-Incident On-Site",                 500),
    ("IT Consultation – Per Session (2 hours)",                  800),
    ("Cybersecurity Audit & Assessment",                        6000),
    ("Cloud Migration & Deployment",                           12000),
    ("Data Analytics Dashboard Setup",                          9000),
    ("Microsoft 365 / Google Workspace Deployment",             3500),
    ("Staff IT Training Programme (per session)",               1200),
    ("VPN & Network Security Configuration",                    4500),
    ("2-Year Enterprise Technical Support Package",            18000),
]

# ══════════════════════════════════════════════════════════════════════
#  Excel column definitions
# ══════════════════════════════════════════════════════════════════════

# --- INFINITE GZ columns (v3: Items split into 4 cols) ---------------
IGZ_COLS = [
    "Invoice No",
    "Date",
    "Due Date",
    "Bill To",
    "Company",
    "Payment Type",        # dropdown
    "Card Type",           # dropdown
    "Card No",
    "Approval Code",
    "Receipt No",
    "Ref No",
    # ↓ replaces "Items (JSON)"
    "Product Item",        # dropdown (IGZ transaction types)
    "Qty",
    "Unit Price (RM)",
    "Total Amount (RM)",
    # ↓ kept
    "Subtotal (RM)",
    "Promo Rebate (RM)",
    "Tax",
    "Total (RM)",
    "Remarks",
    "Receipt Link",        # wide + scroll
    "Invoice Link",        # wide + scroll
]

# --- AI SMART TECH columns (unchanged from v2) -----------------------
AST_COLS = [
    "Invoice No",
    "Date",
    "Due Date",
    "Bill To",
    "Company",
    "Payment Type",
    "Card Type",
    "Card No",
    "Approval Code",
    "Receipt No",
    "Ref No",
    "Product Item",
    "Qty",
    "Unit Price (RM)",
    "Total Amount (RM)",
    "Subtotal (RM)",
    "Promo Rebate (RM)",
    "Tax",
    "Total (RM)",
    "Remarks",
    "Receipt Link",
    "Invoice Link",
]

def get_cols(company: str):
    return AST_COLS if _is_ast(company) else IGZ_COLS

def _is_ast(company: str) -> bool:
    return "AI SMART" in company or "AST" in company

def _products(company: str):
    return AST_PRODUCTS if _is_ast(company) else IGZ_PRODUCTS

# ══════════════════════════════════════════════════════════════════════
#  Excel helpers
# ══════════════════════════════════════════════════════════════════════
def _thin():
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)

def init_excel(path: Path, company: str):
    if path.exists():
        wb = openpyxl.load_workbook(path)
    else:
        wb = openpyxl.Workbook()
        wb.active.title = "Invoices"
    ws = wb["Invoices"]
    if ws.max_row == 0 or (ws.max_row == 1 and ws.cell(1,1).value is None):
        _write_header(ws, company)
        _add_dropdowns(ws, company)
        wb.save(path)
    return wb, ws


def _write_header(ws, company: str):
    COLS = get_cols(company)
    if _is_ast(company):
        hdr_fill = PatternFill("solid", fgColor="6B5B95")
    else:
        hdr_fill = PatternFill("solid", fgColor="1A1A1A")
    hdr_font = Font(name="Calibri", bold=True, color="FFFFFF", size=10)

    # Row 1: Banner
    ws.merge_cells(f"A1:{get_column_letter(len(COLS))}1")
    c = ws["A1"]
    c.value     = f"{company}  —  Invoice Records"
    c.font      = Font(name="Calibri", bold=True, color="FFFFFF", size=13)
    c.fill      = hdr_fill
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26

    # Row 2: Column headers
    COL_W = _col_widths(company)
    for ci, col in enumerate(COLS, 1):
        cell = ws.cell(row=2, column=ci, value=col)
        cell.font      = hdr_font
        cell.fill      = hdr_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border    = _thin()
        ws.column_dimensions[get_column_letter(ci)].width = COL_W.get(col, 14)
    ws.row_dimensions[2].height = 32
    ws.freeze_panes = "A3"


def _col_widths(company: str) -> dict:
    # Both companies now have the same 4-column product structure
    base = {
        "Invoice No": 13, "Date": 14, "Due Date": 14, "Bill To": 16,
        "Company": 22, "Payment Type": 14, "Card Type": 12, "Card No": 22,
        "Approval Code": 14, "Receipt No": 12, "Ref No": 18,
        "Product Item": 48,
        "Qty": 6, "Unit Price (RM)": 15, "Total Amount (RM)": 16,
        "Subtotal (RM)": 14, "Promo Rebate (RM)": 17,
        "Tax": 7, "Total (RM)": 13,
        "Remarks": 30,
        "Receipt Link": 42,
        "Invoice Link": 42,
    }
    return base


def _add_dropdowns(ws, company: str):
    """Add Excel DataValidation dropdowns for all companies."""
    COLS = get_cols(company)
    is_ast = _is_ast(company)

    def col_letter(name):
        idx = COLS.index(name) + 1
        return get_column_letter(idx)

    # ── Payment Type dropdown ──────────────────────────────────────
    if "Payment Type" in COLS:
        dv_pay = DataValidation(
            type="list",
            formula1='"Cash,Visa,Master,Amex,Transfer,E-Wallet"',
            allow_blank=True, showDropDown=False,
        )
        dv_pay.sqref       = f"{col_letter('Payment Type')}3:{col_letter('Payment Type')}500"
        dv_pay.prompt      = "请选择付款方式"
        dv_pay.promptTitle = "Payment Type"
        ws.add_data_validation(dv_pay)

    # ── Card Type dropdown ────────────────────────────────────────
    if "Card Type" in COLS:
        dv_card = DataValidation(
            type="list",
            formula1='"Credit,Debit"',
            allow_blank=True, showDropDown=False,
        )
        dv_card.sqref      = f"{col_letter('Card Type')}3:{col_letter('Card Type')}500"
        dv_card.prompt     = "Credit 或 Debit"
        dv_card.promptTitle= "Card Type"
        ws.add_data_validation(dv_card)

    # ── Product Item dropdown (both companies, different lists) ──
    if "Product Item" in COLS:
        products = _products(company)
        sheet_name = "IGZProductList" if not is_ast else "ASTProductList"
        try:
            if sheet_name not in [s.title for s in ws.parent.worksheets]:
                pl_ws = ws.parent.create_sheet(sheet_name)
            else:
                pl_ws = ws.parent[sheet_name]

            for i, (label, price) in enumerate(products, 1):
                pl_ws.cell(row=i, column=1, value=label)
                pl_ws.cell(row=i, column=2, value=price)

            max_row = len(products)
            dv_prod = DataValidation(
                type="list",
                formula1=f"{sheet_name}!$A$1:$A${max_row}",
                allow_blank=True, showDropDown=False,
            )
            dv_prod.sqref       = f"{col_letter('Product Item')}3:{col_letter('Product Item')}500"
            dv_prod.prompt      = "从下拉选择服务/交易类型"
            dv_prod.promptTitle = "Product Item"
            ws.add_data_validation(dv_prod)
            pl_ws.sheet_state = "hidden"
        except Exception as e:
            print(f"[WARN] Product Item dropdown: {e}")


def append_row(wb, ws, data: dict, path: Path, company: str):
    COLS = get_cols(company)
    is_ast = _is_ast(company)
    even_fill = PatternFill("solid", fgColor="F9F5FF" if is_ast else "FFF8F0")
    odd_fill  = PatternFill("solid", fgColor="FFFFFF")
    lc = "6B5B95" if is_ast else "1A1A1A"

    row_num = ws.max_row + 1
    fill    = even_fill if row_num % 2 == 0 else odd_fill

    for ci, col in enumerate(COLS, 1):
        val  = data.get(col, "")
        cell = ws.cell(row=row_num, column=ci, value=val)
        cell.fill      = fill
        cell.alignment = Alignment(vertical="center", wrap_text=(col == "Product Item"))
        cell.border    = _thin()
        if col in ("Receipt Link", "Invoice Link") and val:
            cell.font = Font(name="Calibri", color=lc, underline="single", size=9)
        else:
            cell.font = Font(name="Calibri", size=9)

    ws.row_dimensions[row_num].height = 20
    wb.save(path)
    return row_num


def update_cell(wb, ws, row_num: int, col_name: str, value, path: Path, company: str):
    COLS = get_cols(company)
    if col_name not in COLS:
        return
    ci   = COLS.index(col_name) + 1
    cell = ws.cell(row=row_num, column=ci, value=value)
    lc   = "6B5B95" if _is_ast(company) else "1A1A1A"
    if col_name in ("Receipt Link", "Invoice Link") and value:
        cell.font = Font(name="Calibri", color=lc, underline="single", size=9)
    wb.save(path)


# ══════════════════════════════════════════════════════════════════════
#  OCR + parser
# ══════════════════════════════════════════════════════════════════════
def ocr_image(img: Image.Image) -> str:
    return pytesseract.image_to_string(img, lang="eng")


def _clean_inv_no(raw: str) -> str:
    """Strip INV- / INV_ prefix; return bare number/code string."""
    s = raw.strip()
    s = re.sub(r"(?i)^INV[-_\s]*", "", s)
    return s


def parse_receipt(text: str, company: str) -> dict:
    d = {}
    def _s(patterns):
        for pat in patterns:
            m = re.search(pat, text, re.IGNORECASE)
            if m:
                return m.group(1).strip()
        return ""

    raw_inv = _s([r"invoice\s*no[\s.:]+([A-Z0-9_\-]+)",
                  r"INV[-_]?\s*([0-9]{3,})"])
    d["Invoice No"] = _clean_inv_no(raw_inv) if raw_inv else ""

    d["Receipt No"] = _s([r"receipt\s*no[\s.:]+([0-9]+)"])
    d["Date"]       = _s([r"date[\s.:]+(\\d{1,2}[\s/\-]\\w+[\s/\-]\\d{4})",
                          r"(\d{1,2}\s+\w+\s+\d{4})"])
    d["Due Date"]   = d["Date"]
    d["Bill To"]    = _s([r"bill\s*to[\s.:]+([A-Za-z ]+?)(?:\n|company|payment)",
                          r"issued\s*to[\s.:]+([A-Za-z ]+?)(?:\n|company)"])
    d["Bill To"]    = re.sub(r"\s+", " ", d["Bill To"]).strip()
    d["Payment Type"] = _s([r"payment[\s.:]+([A-Za-z ]+?)(?:\n|card|approval)",
                             r"(visa|mastercard|cash|transfer|amex|e.?wallet)"])
    d["Card Type"]  = _s([r"card\s*type[\s.:]+([A-Za-z ]+?)(?:\n|payment|card\s*no)",
                          r"(visa\s*card|mastercard|credit|debit)"])
    d["Card No"]    = _s([r"card\s*no[\s.:]+([0-9*X ]{10,})",
                          r"(\d{4}[\s*]+\d{2,4}[\s*]+\*+[\s*]+\d{4})"])
    d["Approval Code"] = _s([r"approval(?:\s*code)?[\s.:]+([A-Z0-9]+)"])
    d["Ref No"]     = _s([r"ref(?:erence)?\s*no[\s.:]+([0-9A-Z]+)"])

    total_s = _s([r"total\s*amount[\s.:]+RM\s*([\d,]+\.?\d*)",
                  r"total[\s.:]+RM\s*([\d,]+\.?\d*)"])
    d["Total (RM)"] = float(total_s.replace(",","")) if total_s else ""

    sub_s = _s([r"subtotal[\s.:]+RM\s*([\d,]+\.?\d*)"])
    d["Subtotal (RM)"] = float(sub_s.replace(",","")) if sub_s else d.get("Total (RM)","")

    promo_s = _s([r"promo\s*rebate[\s.:(]+RM\s*([\d,]+\.?\d*)",
                  r"discount[\s.:(]+RM\s*([\d,]+\.?\d*)"])
    d["Promo Rebate (RM)"] = f"-{promo_s}" if promo_s else ""
    d["Tax"]     = _s([r"tax[\s.:]+([^\n]+)"]) or "-"
    d["Remarks"] = _s([r"remarks?[\s.:]+([^\n]+)"]) or ""

    # Product item defaults
    d["Product Item"]      = ""
    d["Qty"]               = 1
    d["Unit Price (RM)"]   = d.get("Total (RM)", "")
    d["Total Amount (RM)"] = d.get("Total (RM)", "")

    return d


# ══════════════════════════════════════════════════════════════════════
#  Invoice PDF generator wrapper
# ══════════════════════════════════════════════════════════════════════
def generate_invoice_pdf(company: str, data: dict, inv_number: str,
                         items_list: list = None) -> Path:
    items = items_list or []
    if not items:
        try:
            unit_p = float(str(data.get("Unit Price (RM)","") or data.get("Total (RM)","0")).replace(",",""))
            qty    = int(data.get("Qty", 1) or 1)
            amt    = float(str(data.get("Total Amount (RM)","") or data.get("Total (RM)","0")).replace(",",""))
            desc   = data.get("Product Item","Services rendered") or "Services rendered"
            # Strip code prefix [X-000] for cleaner PDF display
            desc_clean = re.sub(r"^\[[A-Z]-\d{3}\]\s*", "", desc)
            items  = [{"desc": desc_clean, "unit_price": unit_p, "qty": qty, "amount": amt}]
        except Exception:
            total = float(str(data.get("Total (RM)", 0) or 0).replace(",",""))
            items = [{"desc": "Services rendered", "unit_price": total, "qty": 1, "amount": total}]

    inv_no_clean = _clean_inv_no(str(inv_number))

    inv = {
        "inv_no":       inv_no_clean,
        "date":         data.get("Date",""),
        "due_date":     data.get("Due Date", data.get("Date","")),
        "bill_to":      data.get("Bill To",""),
        "company":      data.get("Company","–"),
        "payment_type": data.get("Payment Type",""),
        "card_type":    data.get("Card Type",""),
        "card_no":      data.get("Card No",""),
        "approval":     data.get("Approval Code",""),
        "receipt_no":   data.get("Receipt No",""),
        "ref_no":       data.get("Ref No",""),
        "tax":          data.get("Tax","-"),
        "total":        float(str(data.get("Total (RM)",0) or 0).replace(",","")),
        "items":        items,
        "remarks":      [data.get("Remarks","")] if data.get("Remarks") else [],
    }
    promo = data.get("Promo Rebate (RM)","")
    if promo:
        try: inv["promo_rebate"] = float(str(promo).replace(",",""))
        except: pass

    ts       = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_no  = re.sub(r"[^A-Za-z0-9_\-]","_", inv_no_clean)
    pdf_path = INVOICES_DIR / f"invoice_{safe_no}_{ts}.pdf"

    if _is_ast(company):
        make_ast(inv, str(pdf_path))
    else:
        make_igz(inv, str(pdf_path))
    return pdf_path


# ══════════════════════════════════════════════════════════════════════
#  Streamlit UI
# ══════════════════════════════════════════════════════════════════════
def main():
    st.set_page_config(
        page_title="Invoice Management System",
        page_icon="🧾",
        layout="wide",
        initial_sidebar_state="collapsed",
    )

    st.markdown("""
    <style>
    .stApp { background: #F7F8FA; }
    div[data-testid="stDataFrame"] > div { overflow-x: auto !important; }
    div[data-testid="stDataFrame"] iframe { width: 100% !important; }
    .section-title { font-size:17px; font-weight:700; margin:4px 0 10px 0; color:#1A1A1A; }
    .info-box { background:#fff; border-radius:10px; padding:14px 18px;
                border:1px solid #E5E5E5; margin-bottom:12px; }
    .price-row { font-size:11px; padding:2px 0; border-bottom:1px solid #eee; }
    </style>
    """, unsafe_allow_html=True)

    # ── Session state ─────────────────────────────────────────────
    for k, v in [("company",None),("parsed_data",{}),("receipt_path",None),
                 ("excel_row",None),("invoice_done",False),("invoice_path",None),
                 ("items_list",[])]:
        if k not in st.session_state:
            st.session_state[k] = v

    # ══════════════════════════════════════════════════════════════
    #  SCREEN 1 — Company selection
    # ══════════════════════════════════════════════════════════════
    if st.session_state.company is None:
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("## 🧾 Invoice Management System")
        st.markdown("请选择公司 / Select company to proceed:")
        st.markdown("<br>", unsafe_allow_html=True)

        c1, c2, c3 = st.columns([1, 1, 1.2])

        # ── IGZ card ──────────────────────────────────────────────
        with c1:
            st.markdown("""
            <div style='background:linear-gradient(135deg,#1A1A1A,#333);color:#fff;
                        border-radius:14px;padding:28px 20px;text-align:center;'>
                <div style='font-size:34px'>🏢</div>
                <div style='font-size:19px;font-weight:800;margin:8px 0'>INFINITE GZ SDN BHD</div>
                <div style='font-size:11px;opacity:.7'>202401019141 (1564990-X)</div>
                <div style='font-size:11px;opacity:.6;margin-top:4px'>Hong Leong Bank · Black Minimal Style</div>
                <div style='font-size:10px;opacity:.5;margin-top:4px'>51 Transaction Types | 5 Groups</div>
            </div>""", unsafe_allow_html=True)
            if st.button("选择 INFINITE GZ  →", key="btn_igz", use_container_width=True):
                st.session_state.company = "INFINITE GZ SDN BHD"; st.rerun()

        # ── AST card ──────────────────────────────────────────────
        with c2:
            st.markdown("""
            <div style='background:linear-gradient(135deg,#6B5B95,#9B8EC4);color:#fff;
                        border-radius:14px;padding:28px 20px;text-align:center;'>
                <div style='font-size:34px'>💼</div>
                <div style='font-size:19px;font-weight:800;margin:8px 0'>AI SMART TECH SDN BHD</div>
                <div style='font-size:11px;opacity:.7'>202401043356 (1589202-V)</div>
                <div style='font-size:11px;opacity:.6;margin-top:4px'>Purple Corporate · AI Tech Services</div>
                <div style='font-size:10px;opacity:.5;margin-top:4px'>37 AI Services | 2026 Market Prices</div>
            </div>""", unsafe_allow_html=True)
            if st.button("选择 AI SMART TECH →", key="btn_ast", use_container_width=True):
                st.session_state.company = "AI SMART TECH SDN BHD"; st.rerun()

        # ── RIGHT COLUMN: both price tables ──────────────────────
        with c3:
            tab_igz, tab_ast = st.tabs(["🏢 IGZ 服务价格", "💼 AST 服务价格"])

            with tab_igz:
                st.markdown("**INFINITE GZ — 2026 业务服务参考价**")
                # Group headers
                groups = {
                    "Group A · 融资安排": [(l,p) for l,p in IGZ_PRODUCTS if l.startswith("[A-")],
                    "Group B · 记录处理": [(l,p) for l,p in IGZ_PRODUCTS if l.startswith("[B-")],
                    "Group C · 数字服务": [(l,p) for l,p in IGZ_PRODUCTS if l.startswith("[C-")],
                    "Group D · 成果分享": [(l,p) for l,p in IGZ_PRODUCTS if l.startswith("[D-")],
                    "Group E · 账户操作": [(l,p) for l,p in IGZ_PRODUCTS if l.startswith("[E-")],
                }
                for gname, items in groups.items():
                    st.markdown(f"<div style='font-size:10px;font-weight:700;color:#555;margin-top:6px'>{gname}</div>",
                                unsafe_allow_html=True)
                    for label, price in items[:4]:
                        short = re.sub(r"^\[[A-Z]-\d{3}\]\s*","",label)
                        price_str = f"RM {price:,}" if price > 0 else "–"
                        st.markdown(
                            f"<div class='price-row'><b>{price_str}</b> · {short[:38]}</div>",
                            unsafe_allow_html=True)
                st.caption("完整列表见 Product Item 下拉菜单（51项）")

            with tab_ast:
                st.markdown("**AI SMART TECH — 2026 AI 服务参考价**")
                for label, price in AST_PRODUCTS[:14]:
                    st.markdown(
                        f"<div class='price-row'><b>RM {price:,}</b> · {label[:42]}</div>",
                        unsafe_allow_html=True)
                st.caption("完整列表见 Product Item 下拉菜单（37项）")

        return

    # ══════════════════════════════════════════════════════════════
    #  SCREEN 2 — Main workflow
    # ══════════════════════════════════════════════════════════════
    company = st.session_state.company
    is_ast  = _is_ast(company)
    bar_col = "#6B5B95" if is_ast else "#1A1A1A"

    st.markdown(f"""
    <div style='background:{bar_col};color:#fff;padding:13px 20px;border-radius:10px;
                display:flex;align-items:center;justify-content:space-between;margin-bottom:16px'>
        <span style='font-size:19px;font-weight:800'>{'💼' if is_ast else '🏢'} {company}</span>
        <span style='font-size:11px;opacity:.7'>Invoice Management System v3</span>
    </div>""", unsafe_allow_html=True)

    if st.button("← 切换公司 / Switch Company"):
        for k in ["company","parsed_data","receipt_path","excel_row",
                  "invoice_done","invoice_path","items_list"]:
            st.session_state[k] = {} if k in ("parsed_data",) else ([] if k=="items_list" else None)
        st.session_state.invoice_done = False
        st.rerun()

    excel_path = AST_EXCEL if is_ast else IGZ_EXCEL
    wb, ws = init_excel(excel_path, company)

    # ── STEP 1: Upload receipt ─────────────────────────────────────
    st.markdown("---")
    st.markdown("<div class='section-title'>📤 Step 1 · 上传收据</div>", unsafe_allow_html=True)
    uploaded = st.file_uploader("上传收据图片 (PNG / JPG)", type=["png","jpg","jpeg","webp"])

    if uploaded and st.session_state.receipt_path is None:
        ts  = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        ext = Path(uploaded.name).suffix or ".png"
        rec_path = RECEIPTS_DIR / f"receipt_{ts}{ext}"
        with open(rec_path,"wb") as f: f.write(uploaded.read())
        st.session_state.receipt_path = str(rec_path)

        with st.spinner("🔍 正在 OCR 解析收据…"):
            img    = Image.open(rec_path)
            raw    = ocr_image(img)
            parsed = parse_receipt(raw, company)
            parsed["Receipt Link"] = str(rec_path)
            st.session_state.parsed_data = parsed
            row_num = append_row(wb, ws, parsed, excel_path, company)
            st.session_state.excel_row = row_num
        st.success(f"✅ 收据已解析，已写入 Excel 第 {row_num} 行")
        st.rerun()

    if st.session_state.receipt_path:
        with st.expander("📎 已上传的收据", expanded=False):
            try: st.image(st.session_state.receipt_path, width=420)
            except: st.write(st.session_state.receipt_path)

    # ── STEP 2: Review & Edit ─────────────────────────────────────
    if st.session_state.parsed_data:
        st.markdown("---")
        st.markdown("<div class='section-title'>✏️ Step 2 · 确认 / 编辑内容</div>",
                    unsafe_allow_html=True)
        st.info("📌 系统已自动解析收据。**唯一需手动填写**的是「Company 客户公司名称」。"
                "下拉菜单字段直接选择即可。")

        d = st.session_state.parsed_data
        c1, c2 = st.columns(2)

        # ── Left column ──────────────────────────────────────────
        with c1:
            d["Invoice No"]  = st.text_input(
                "Invoice No（收据实际编号）",
                d.get("Invoice No",""),
                help="系统自动去除 INV- 前缀，保留收据原始编号")
            d["Date"]        = st.text_input("Date", d.get("Date",""))
            d["Due Date"]    = st.text_input("Due Date", d.get("Due Date",""))
            d["Bill To"]     = st.text_input("Bill To (客户姓名)", d.get("Bill To",""))
            d["Company"]     = st.text_input(
                "⭐ Company (手动填写，无则填 -)",
                d.get("Company",""),
                help="唯一需手动输入的字段")

            # Payment Type — dropdown for BOTH companies
            pay_opts = ["","Cash","Visa","Master","Amex","Transfer","E-Wallet"]
            cur_pay  = d.get("Payment Type","")
            idx_pay  = pay_opts.index(cur_pay) if cur_pay in pay_opts else 0
            d["Payment Type"] = st.selectbox("Payment Type", pay_opts, index=idx_pay)

            # Card Type — dropdown for BOTH companies
            card_opts = ["","Credit","Debit"]
            cur_card  = d.get("Card Type","")
            idx_card  = card_opts.index(cur_card) if cur_card in card_opts else 0
            d["Card Type"] = st.selectbox("Card Type", card_opts, index=idx_card)

        # ── Right column ─────────────────────────────────────────
        with c2:
            d["Card No"]       = st.text_input("Card No",       d.get("Card No",""))
            d["Approval Code"] = st.text_input("Approval Code", d.get("Approval Code",""))
            d["Receipt No"]    = st.text_input("Receipt No",    d.get("Receipt No",""))
            d["Ref No"]        = st.text_input("Ref No",        d.get("Ref No",""))
            d["Tax"]           = st.text_input("Tax",           d.get("Tax","-"))
            d["Subtotal (RM)"] = st.text_input("Subtotal (RM)", str(d.get("Subtotal (RM)","")))
            d["Promo Rebate (RM)"] = st.text_input("Promo Rebate (RM)",
                                                    str(d.get("Promo Rebate (RM)","")))
            d["Total (RM)"]    = st.text_input("Total (RM)",    str(d.get("Total (RM)","")))

        # ── Product Item section (BOTH companies now use 4 columns) ──
        st.markdown("---")
        label_color = "#6B5B95" if is_ast else "#1A1A1A"
        co_label    = "AI 服务项目" if is_ast else "IGZ 交易类型"
        st.markdown(
            f"<div class='section-title' style='color:{label_color}'>🛒 {co_label}</div>",
            unsafe_allow_html=True)

        products    = _products(company)
        prod_labels = [""] + [label for label, _ in products]
        cur_prod    = d.get("Product Item","")
        idx_prod    = prod_labels.index(cur_prod) if cur_prod in prod_labels else 0

        help_text = ("选择服务后自动带入 2026 参考单价" if is_ast
                     else "选择 IGZ 交易类型（来源：IGZ Transaction Reference v1.0）")
        selected_prod = st.selectbox(
            f"Product Item（{co_label}）",
            prod_labels,
            index=idx_prod,
            help=help_text)

        auto_price = ""
        if selected_prod:
            for lbl, price in products:
                if lbl == selected_prod:
                    auto_price = str(price) if price > 0 else ""
                    break

        d["Product Item"] = selected_prod

        pc1, pc2, pc3 = st.columns([3, 1, 1])
        with pc1:
            st.text_input("已选项目", selected_prod, disabled=True)
        with pc2:
            d["Qty"] = st.number_input("Qty", min_value=1, max_value=999,
                                       value=int(d.get("Qty",1) or 1))
        with pc3:
            d["Unit Price (RM)"] = st.text_input(
                "Unit Price (RM)",
                value=d.get("Unit Price (RM)","") or auto_price,
                help="选择后自动带入参考价，可手动修改")

        # Auto-calc line total
        try:
            line_total = float(str(d["Unit Price (RM)"]).replace(",","")) * int(d["Qty"])
            d["Total Amount (RM)"] = f"{line_total:.2f}"
        except:
            d["Total Amount (RM)"] = d.get("Total Amount (RM)","")

        st.markdown(f"**Line Total : RM {d['Total Amount (RM)']}**")

        d["Remarks"] = st.text_area("Remarks", d.get("Remarks",""), height=60)
        st.session_state.parsed_data = d

        # ── STEP 3: Generate Invoice ──────────────────────────────
        st.markdown("---")
        st.markdown("<div class='section-title'>🚀 Step 3 · 生成 Invoice</div>",
                    unsafe_allow_html=True)

        if not st.session_state.invoice_done:
            if st.button("📄  GENERATE INVOICE  —  生成收据",
                         type="primary", use_container_width=True):
                with st.spinner("⚙️ 正在生成 Invoice PDF…"):
                    try:
                        inv_no = d.get("Invoice No") or datetime.datetime.now().strftime("%Y%m%d%H%M")
                        items_list = [{
                            "desc": re.sub(r"^\[[A-Z]-\d{3}\]\s*","",
                                           d.get("Product Item","Services rendered") or "Services rendered"),
                            "unit_price": float(str(d.get("Unit Price (RM)","0")).replace(",","")),
                            "qty":        int(d.get("Qty",1)),
                            "amount":     float(str(d.get("Total Amount (RM)","0")).replace(",","")),
                        }]

                        pdf_p    = generate_invoice_pdf(company, d, inv_no, items_list)
                        inv_link = str(pdf_p)

                        COLS = get_cols(company)
                        for col in COLS:
                            if col not in ("Receipt Link","Invoice Link"):
                                update_cell(wb, ws, st.session_state.excel_row, col,
                                            d.get(col,""), excel_path, company)
                        update_cell(wb, ws, st.session_state.excel_row,
                                    "Invoice Link", inv_link, excel_path, company)

                        st.session_state.invoice_path = str(pdf_p)
                        st.session_state.invoice_done = True
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ 生成失败: {e}")
        else:
            st.success("✅ Invoice 已生成！")
            inv_path = st.session_state.invoice_path
            with open(inv_path,"rb") as f:
                st.download_button("⬇️ 下载 Invoice PDF", data=f,
                                   file_name=Path(inv_path).name,
                                   mime="application/pdf",
                                   use_container_width=True)
            # PNG preview
            png_p = str(Path(inv_path).with_suffix("")) + "_prev.png"
            try:
                subprocess.run(["pdftoppm","-r","150","-png","-singlefile",
                                inv_path, png_p.replace(".png","")],
                               check=True, capture_output=True)
                if os.path.exists(png_p):
                    st.image(png_p, caption="Invoice Preview", use_container_width=True)
            except: pass

            if st.button("📝 新增下一张收据"):
                for k in ["parsed_data","receipt_path","excel_row",
                          "invoice_done","invoice_path","items_list"]:
                    st.session_state[k] = {} if k=="parsed_data" else ([] if k=="items_list" else None)
                st.session_state.invoice_done = False
                st.rerun()

    # ── STEP 4: Excel summary table ───────────────────────────────
    st.markdown("---")
    st.markdown(f"<div class='section-title'>📊 汇总表格 — {company}</div>",
                unsafe_allow_html=True)

    with open(excel_path,"rb") as f:
        st.download_button(
            label=f"⬇️ 下载 Excel 汇总表 ({Path(excel_path).name})",
            data=f, file_name=Path(excel_path).name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    import pandas as pd
    COLS = get_cols(company)
    lc   = "6B5B95" if is_ast else "1A1A1A"
    COL_W = {
        "Invoice No":130,"Date":110,"Due Date":110,"Bill To":120,"Company":160,
        "Payment Type":115,"Card Type":100,"Card No":175,"Approval Code":115,
        "Receipt No":105,"Ref No":150,
        "Product Item":310,
        "Qty":65,"Unit Price (RM)":120,"Total Amount (RM)":135,
        "Subtotal (RM)":120,"Promo Rebate (RM)":140,"Tax":65,
        "Total (RM)":110,"Remarks":200,
        "Receipt Link":280,"Invoice Link":280,
    }

    try:
        df = pd.read_excel(excel_path, skiprows=1, header=0)
        col_cfg = {}
        for col in df.columns:
            w = COL_W.get(col, 120)
            if col in ("Receipt Link","Invoice Link"):
                col_cfg[col] = st.column_config.LinkColumn(col, width=w, display_text="🔗 查看")
            elif col in ("Subtotal (RM)","Unit Price (RM)","Total Amount (RM)","Total (RM)"):
                col_cfg[col] = st.column_config.NumberColumn(col, width=w, format="RM %.2f")
            elif col == "Qty":
                col_cfg[col] = st.column_config.NumberColumn(col, width=w, format="%d")
            else:
                col_cfg[col] = st.column_config.TextColumn(col, width=w)

        st.markdown(
            f"<p style='font-size:11px;color:#888'>共 {len(df)} 条记录 · "
            f"{len(df.columns)} 个字段 · 👉 <b>横向滑动查看全部列</b></p>",
            unsafe_allow_html=True)

        st.dataframe(df, column_config=col_cfg,
                     use_container_width=False,
                     width=2200,
                     height=420,
                     hide_index=True)
    except Exception as e:
        st.warning(f"暂无记录: {e}")

    st.caption(f"📁 {excel_path}  ·  Invoices: {INVOICES_DIR}")


if __name__ == "__main__":
    main()
