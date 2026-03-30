#!/usr/bin/env python3
"""
AI SMART TECH SDN BHD  Invoice Generator
=========================================
Design: Replicates the purple-accent corporate style shown in sample image.
  • Company header top-left  |  "INVOICE" top-right
  • Purple horizontal rule under header
  • Info grid (Invoice No, Date, Bill To, etc.)
  • Table with NO. / DESCRIPTION / QTY / UNIT PRICE(RM) / AMOUNT(RM)
    - Dark header row with white text
    - Numbered item rows
    - Package Promo Rebate row in purple italic (optional)
  • Subtotal / Promo Rebate / TOTAL AMOUNT block
  • PAYMENT DETAILS box (bottom-left)  |  REMARKS box (bottom-right)
  • Footer: thin rule + disclaimer text
"""

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os

# ── Fonts ─────────────────────────────────────────────────────────────
FONT_DIR = "/usr/share/fonts/truetype/liberation/"
pdfmetrics.registerFont(TTFont("AST_Reg",    FONT_DIR + "LiberationSans-Regular.ttf"))
pdfmetrics.registerFont(TTFont("AST_Bold",   FONT_DIR + "LiberationSans-Bold.ttf"))
pdfmetrics.registerFont(TTFont("AST_Serif",  FONT_DIR + "LiberationSerif-Regular.ttf"))
pdfmetrics.registerFont(TTFont("AST_SerifI", FONT_DIR + "LiberationSerif-Italic.ttf"))
pdfmetrics.registerFont(TTFont("AST_SerifB", FONT_DIR + "LiberationSerif-Bold.ttf"))

# ── Colour Palette ─────────────────────────────────────────────────────
WHITE   = colors.HexColor("#FFFFFF")
BLACK   = colors.HexColor("#000000")
DARK    = colors.HexColor("#1A1A1A")
DGRAY   = colors.HexColor("#333333")
MGRAY   = colors.HexColor("#666666")
LGRAY   = colors.HexColor("#CCCCCC")
XLGRAY  = colors.HexColor("#F4F4F4")
PURPLE  = colors.HexColor("#6B5B95")   # primary accent
PURPLE2 = colors.HexColor("#8472AA")   # lighter purple for italic text

# ── Page geometry ──────────────────────────────────────────────────────
W, H   = A4
PAD_L  = 18 * mm
PAD_R  = 18 * mm
PAD_T  = 14 * mm
PAD_B  = 14 * mm
ML     = PAD_L
MR     = W - PAD_R
TW     = MR - ML   # ≈ 174 mm


# ─────────────────────────────────────────────────────────────────────
#  Helpers
# ─────────────────────────────────────────────────────────────────────
def _fmt_unit(val):
    if val == int(val):
        return f"{int(val):,}"
    return f"{val:,.2f}"

def _fmt_amount(val):
    return f"{val:,.2f}"

def _wrap(text, font, size, max_w):
    words = text.split()
    lines, cur = [], ""
    for w in words:
        test = (cur + " " + w).strip()
        if pdfmetrics.stringWidth(test, font, size) <= max_w:
            cur = test
        else:
            if cur:
                lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    return lines or [""]


# ─────────────────────────────────────────────────────────────────────
#  Main draw function
# ─────────────────────────────────────────────────────────────────────
def draw_ast_invoice(c, inv):
    """Draw AI SMART TECH SDN BHD invoice on canvas."""

    # White background
    c.setFillColor(WHITE)
    c.rect(0, 0, W, H, fill=1, stroke=0)

    y = H - PAD_T

    # ══════════════════════════════════════════════════════════════════
    #  1. Header: Company (left) | INVOICE (right)
    # ══════════════════════════════════════════════════════════════════
    CNAME_SZ  = 20
    INV_SZ    = 32

    # Company name
    c.setFont("AST_Bold", CNAME_SZ)
    c.setFillColor(DARK)
    c.drawString(ML, y, "AI SMART TECH SDN BHD")
    y -= 5.5 * mm

    # Sub-lines
    sub_lines = [
        "Registration No : 202401043356 (1589202-V)",
        "L2-K02A CENTRAL i-CITY, NO 1, PERSIARAN MULTIMEDIA",
        "SEKSYEN 7, JALAN PLUMBUM 7/102, SHAH ALAM, SELANGOR",
        "Email : admin@aismarttech.com",
    ]
    for sl in sub_lines:
        c.setFont("AST_Reg", 7.5)
        c.setFillColor(MGRAY)
        c.drawString(ML, y, sl)
        y -= 4.5 * mm

    # "INVOICE" top-right (aligned to top of company name block)
    inv_top_y = H - PAD_T
    c.setFont("AST_Bold", INV_SZ)
    c.setFillColor(DARK)
    inv_w = pdfmetrics.stringWidth("INVOICE", "AST_Bold", INV_SZ)
    c.drawString(MR - inv_w, inv_top_y, "INVOICE")

    # Purple horizontal rule under header block
    y -= 3 * mm
    c.setStrokeColor(PURPLE)
    c.setLineWidth(1.5)
    c.line(ML, y, MR, y)
    y -= 5 * mm

    # ══════════════════════════════════════════════════════════════════
    #  2. Info grid  (left col | right col)
    # ══════════════════════════════════════════════════════════════════
    LBL_SZ   = 8
    VAL_SZ   = 8.5
    ROW_H    = 5.8 * mm
    LEFT_X   = ML
    RIGHT_X  = ML + TW * 0.50
    VAL_OFF  = 28 * mm   # indent from label start to value

    left_fields = [
        ("Invoice No :",   str(inv.get("inv_no", ""))),
        ("Bill To :",      str(inv.get("bill_to", ""))),
        ("Company :",      str(inv.get("company", "–"))),
        ("Payment :",      str(inv.get("payment_type", ""))),
        ("Approval :",     str(inv.get("approval", ""))),
    ]
    right_fields = [
        ("Date :",         str(inv.get("date", ""))),
        ("Receipt No :",   str(inv.get("receipt_no", ""))),
        ("Card No :",      str(inv.get("card_no", ""))),
        ("Ref No :",       str(inv.get("ref_no", ""))),
    ]

    fy = y
    for lbl, val in left_fields:
        c.setFont("AST_Reg", LBL_SZ)
        c.setFillColor(MGRAY)
        c.drawString(LEFT_X, fy, lbl)
        c.setFont("AST_Reg", VAL_SZ)
        c.setFillColor(DARK)
        c.drawString(LEFT_X + VAL_OFF, fy, val)
        fy -= ROW_H

    fy2 = y
    for lbl, val in right_fields:
        c.setFont("AST_Reg", LBL_SZ)
        c.setFillColor(MGRAY)
        c.drawString(RIGHT_X, fy2, lbl)
        c.setFont("AST_Reg", VAL_SZ)
        if lbl == "Card No :":
            c.setFont("AST_Bold", VAL_SZ)
        c.setFillColor(DARK)
        c.drawString(RIGHT_X + VAL_OFF, fy2, val)
        fy2 -= ROW_H

    y = min(fy, fy2) - 8 * mm

    # ══════════════════════════════════════════════════════════════════
    #  3. Item table
    # ══════════════════════════════════════════════════════════════════
    # Column positions
    C_NO    = ML                     # NO. col left edge
    C_DESC  = ML + 11 * mm           # description left
    C_QTY_C = MR - 70 * mm          # qty centre
    C_UP_R  = MR - 33 * mm          # unit price right
    C_AMT_R = MR                     # amount right
    DESC_W  = C_QTY_C - 8*mm - C_DESC

    THEAD_H = 9 * mm
    # Table header background (dark)
    c.setFillColor(DARK)
    c.rect(ML, y - THEAD_H, TW, THEAD_H, fill=1, stroke=0)

    # Header text (white)
    th_y = y - THEAD_H + 3 * mm
    c.setFont("AST_Bold", 8)
    c.setFillColor(WHITE)
    c.drawString(C_NO,   th_y, "NO.")
    c.drawString(C_DESC, th_y, "DESCRIPTION")
    c.drawCentredString(C_QTY_C, th_y, "QTY")
    c.drawRightString(C_UP_R, th_y, "UNIT PRICE (RM)")
    c.drawRightString(C_AMT_R, th_y, "AMOUNT (RM)")

    y -= THEAD_H

    # ── Item rows ─────────────────────────────────────────────────────
    ROW_LINE_H = 4.8 * mm
    subtotal = 0
    items = inv.get("items", [])

    for idx, item in enumerate(items, start=1):
        desc_lines = _wrap(item["desc"], "AST_Reg", 8, DESC_W)
        row_h = max(5.5 * mm, len(desc_lines) * ROW_LINE_H + 2 * mm)

        # Alternating very-light background
        if idx % 2 == 0:
            c.setFillColor(XLGRAY)
            c.rect(ML, y - row_h, TW, row_h, fill=1, stroke=0)

        ty = y - 2.5 * mm
        c.setFont("AST_Reg", 8)
        c.setFillColor(DARK)

        c.drawString(C_NO, ty, str(idx))

        for li, dl in enumerate(desc_lines):
            c.drawString(C_DESC, ty - li * ROW_LINE_H, dl)

        if item.get("unit_price") is not None:
            c.drawRightString(C_UP_R, ty, _fmt_unit(item["unit_price"]))
        if item.get("qty") is not None:
            c.drawCentredString(C_QTY_C, ty, str(item["qty"]))
        if item.get("amount") is not None:
            c.drawRightString(C_AMT_R, ty, _fmt_amount(item["amount"]))
            subtotal += item["amount"]

        y -= row_h

    # Package Promo Rebate row (purple italic, if present)
    promo = inv.get("promo_rebate")
    if promo is not None:
        row_h = 7 * mm
        ty = y - 2.5 * mm
        c.setFont("AST_SerifI", 8.5)
        c.setFillColor(PURPLE)
        c.drawString(C_DESC, ty, "Package Promo Rebate")
        c.drawRightString(C_AMT_R, ty, f"(RM {_fmt_amount(abs(promo))})")
        y -= row_h

    # Bottom table border line
    c.setStrokeColor(LGRAY)
    c.setLineWidth(0.5)
    c.line(ML, y, MR, y)
    y -= 5 * mm

    # ══════════════════════════════════════════════════════════════════
    #  4. Totals block (right-aligned)
    # ══════════════════════════════════════════════════════════════════
    TLBL_X  = MR - 70 * mm
    TVAL_X  = MR

    # Subtotal
    c.setFont("AST_Reg", 9)
    c.setFillColor(DARK)
    c.drawRightString(TLBL_X, y, "Subtotal")
    c.drawRightString(TVAL_X, y, f"RM {_fmt_amount(subtotal)}")
    y -= 6.5 * mm

    # Promo rebate line (purple, if present)
    if promo is not None:
        c.setFont("AST_SerifI", 9)
        c.setFillColor(PURPLE)
        c.drawRightString(TLBL_X, y, "Package Promo Rebate")
        c.drawRightString(TVAL_X, y, f"(RM {_fmt_amount(abs(promo))})")
        y -= 6.5 * mm

    # Thin separator
    c.setStrokeColor(LGRAY)
    c.setLineWidth(0.4)
    c.line(TLBL_X - 5*mm, y + 2*mm, TVAL_X, y + 2*mm)
    y -= 3 * mm

    # TOTAL AMOUNT — full-width dark bar
    total = inv.get("total", subtotal)
    if promo is not None:
        total = subtotal + promo  # promo is negative
    TOTAL_H = 10 * mm
    c.setFillColor(DARK)
    c.rect(ML, y - TOTAL_H, TW, TOTAL_H, fill=1, stroke=0)

    # Label left, value right
    tot_ty = y - TOTAL_H + 3.2 * mm
    c.setFont("AST_Bold", 10)
    c.setFillColor(WHITE)
    c.drawString(ML + 4*mm, tot_ty, "TOTAL AMOUNT")
    c.drawRightString(MR - 2*mm, tot_ty, f"RM {_fmt_amount(total)}")
    y -= TOTAL_H + 5 * mm

    # ══════════════════════════════════════════════════════════════════
    #  5. Payment Details (left box)  |  Remarks (right)
    # ══════════════════════════════════════════════════════════════════
    BOX_W   = TW * 0.48
    BOX_X   = ML
    REM_X   = ML + TW * 0.52
    REM_W   = TW * 0.48

    pd_fields = [
        ("Card Type :",      inv.get("card_type", "–")),
        ("Payment Type :",   inv.get("payment_type", "–")),
        ("Card No :",        inv.get("card_no", "–")),
        ("Approval Code :",  inv.get("approval", "–")),
        ("Ref No :",         inv.get("ref_no", "–")),
        ("Amount :",         f"RM {_fmt_amount(total)}"),
    ]

    # Box height estimate
    BOX_H = (len(pd_fields) + 1) * 6 * mm + 8 * mm
    box_top = y
    box_bottom = box_top - BOX_H

    # Draw bordered box
    c.setStrokeColor(LGRAY)
    c.setFillColor(WHITE)
    c.setLineWidth(0.8)
    c.rect(BOX_X, box_bottom, BOX_W, BOX_H, fill=1, stroke=1)

    # "PAYMENT DETAILS" header inside box
    pd_label_y = box_top - 6 * mm
    c.setFont("AST_Bold", 8.5)
    c.setFillColor(DARK)
    c.drawString(BOX_X + 4*mm, pd_label_y, "PAYMENT DETAILS")
    pd_label_y -= 5 * mm

    # Thin line under PD header
    c.setStrokeColor(LGRAY)
    c.setLineWidth(0.4)
    c.line(BOX_X + 3*mm, pd_label_y + 1*mm, BOX_X + BOX_W - 3*mm, pd_label_y + 1*mm)
    pd_label_y -= 3.5 * mm

    for lbl, val in pd_fields:
        c.setFont("AST_Reg", 7.5)
        c.setFillColor(MGRAY)
        c.drawString(BOX_X + 4*mm, pd_label_y, lbl)
        c.setFillColor(DARK)
        if lbl == "Amount :":
            c.setFont("AST_Bold", 7.5)
        c.drawString(BOX_X + 4*mm + 26*mm, pd_label_y, val)
        pd_label_y -= 5.5 * mm

    # Remarks (right side)
    rem_y = y
    c.setFont("AST_Bold", 8.5)
    c.setFillColor(DARK)
    c.drawString(REM_X, rem_y, "REMARKS")
    rem_y -= 5 * mm

    c.setStrokeColor(LGRAY)
    c.setLineWidth(0.4)
    c.line(REM_X, rem_y + 1*mm, REM_X + REM_W, rem_y + 1*mm)
    rem_y -= 3.5 * mm

    remarks = inv.get("remarks", [
        f"Date / Time : {inv.get('date', '')}",
        "2-Year enterprise support from invoice date.",
        "Hardware under 1-Year Huawei warranty.",
    ])
    for rline in remarks:
        c.setFont("AST_Reg", 7.5)
        c.setFillColor(DARK)
        c.drawString(REM_X, rem_y, rline)
        rem_y -= 5 * mm

    # ══════════════════════════════════════════════════════════════════
    #  6. Footer
    # ══════════════════════════════════════════════════════════════════
    FT_Y = PAD_B + 8 * mm

    c.setStrokeColor(LGRAY)
    c.setLineWidth(0.6)
    c.line(ML, FT_Y + 6*mm, MR, FT_Y + 6*mm)

    c.setFont("AST_Reg", 7)
    c.setFillColor(MGRAY)
    c.drawCentredString(W / 2, FT_Y + 2*mm,
        "This is a computer-generated invoice. No signature required.")
    c.drawCentredString(W / 2, FT_Y - 3*mm,
        "Thank you for your business with AI Smart Tech Sdn Bhd.")

    # Safety check
    if y < FT_Y + 32 * mm:
        print(f"  [WARN] {inv.get('inv_no','')} content too long, may overlap footer!")
    else:
        print(f"  [OK]   AST {inv.get('inv_no','')}  clearance = {(y - FT_Y - 32*mm)/mm:.1f}mm")


# ─────────────────────────────────────────────────────────────────────
#  Public API
# ─────────────────────────────────────────────────────────────────────
def make_ast(inv, path):
    """Generate one AI SMART TECH invoice PDF."""
    c = canvas.Canvas(path, pagesize=A4)
    c.setTitle(f"Invoice {inv['inv_no']}")
    draw_ast_invoice(c, inv)
    c.showPage()
    c.save()
    return path


# ─────────────────────────────────────────────────────────────────────
#  Self-test
# ─────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    import subprocess

    test_inv = {
        "inv_no":       "000959",
        "date":         "26 March 2026",
        "bill_to":      "Max Lai",
        "company":      "Senselite Sdn Bhd",
        "payment_type": "Visa Credit",
        "card_type":    "Visa Card",
        "approval":     "038549",
        "receipt_no":   "000959",
        "card_no":      "4141 70** **** 6003",
        "ref_no":       "608514293775",
        "tax":          "-",
        "promo_rebate": -2757.00,
        "total":        18888.00,
        "remarks": [
            "Date / Time : 26 Mar 2026  22:29:38",
            "2-Year enterprise support from invoice date.",
            "Hardware under 1-Year Huawei warranty.",
        ],
        "items": [
            {"desc": "Huawei MateBook X Pro 2024 Laptop — Intel Core Ultra 7, 16GB RAM, 1TB SSD, 14.2\" OLED 2.5K Touch Display, 980g Ultra-Light", "unit_price": 7999, "qty": 1, "amount": 7999.00},
            {"desc": "Huawei Mate 80 Pro 16GB+512GB Smartphone — True-to-Colour Camera System, HarmonyOS 5, IP68 Waterproof Rating", "unit_price": 3499, "qty": 1, "amount": 3499.00},
            {"desc": "Huawei MatePad 11.5 S 2026 Tablet — 8GB+256GB, Kirin 9000WL, 144Hz TFT LCD, 8800mAh, NearLink M-Pencil Support", "unit_price": 1599, "qty": 1, "amount": 1599.00},
            {"desc": "Huawei AirEngine 5761-11 Indoor WiFi 6 Access Point + S1730S-L8P-M 8-Port PoE Gigabit Smart Switch Bundle", "unit_price": 1099, "qty": 1, "amount": 1099.00},
            {"desc": "Huawei FreeClip 2 Open-Ear Wireless Earbuds — Bluetooth 5.3, Dual-Diaphragm 10.8mm Driver, 36H Total Battery Life", "unit_price": 649, "qty": 1, "amount": 649.00},
            {"desc": "System Setup: Enterprise IT Infrastructure Deployment — Server Rack, Fortinet Firewall & Managed Switch Configuration", "unit_price": 1200, "qty": 1, "amount": 1200.00},
            {"desc": "System Setup: Microsoft 365 Business Premium Deployment, CRM Integration & Custom Workflow Automation", "unit_price": 1000, "qty": 1, "amount": 1000.00},
            {"desc": "System Setup: Enterprise Network Security & Site-to-Site VPN Configuration (FortiGate / Cisco ASA)", "unit_price": 800, "qty": 1, "amount": 800.00},
            {"desc": "System Setup: Cloud Data Migration & Automated Incremental Backup (Microsoft Azure / AWS S3 / Synology NAS)", "unit_price": 500, "qty": 1, "amount": 500.00},
            {"desc": "System Setup: Staff IT Training & Onboarding Program — 3 Structured Sessions (Max 15 Users per Session)", "unit_price": 300, "qty": 1, "amount": 300.00},
            {"desc": "2-Year Enterprise Technical Support & Priority Maintenance — Dedicated Account Manager, SLA On-Site Response < 2 Hours", "unit_price": 3000, "qty": 1, "amount": 3000.00},
        ],
    }

    OUT = "/mnt/user-data/outputs"
    os.makedirs(OUT, exist_ok=True)
    pdf_path = f"{OUT}/AST_preview.pdf"
    png_path = f"{OUT}/preview_ast.png"

    make_ast(test_inv, pdf_path)
    print(f"PDF saved: {pdf_path}")

    subprocess.run(
        ["pdftoppm", "-r", "180", "-png", "-singlefile",
         pdf_path, png_path.replace(".png", "")],
        check=True
    )
    print(f"Preview PNG: {png_path}")
