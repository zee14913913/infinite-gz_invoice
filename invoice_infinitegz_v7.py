#!/usr/bin/env python3
"""
INFINITE GZ SDN BHD Invoice Generator — v7
════════════════════════════════════════════
3 FIXES vs v5 (user review round 4):
  ✅ [FIX-D]  Tax label font size explicitly confirmed = SUBTOT_FONT_SZ (8pt)
              Tax label colour → DARK (same visual weight as SUBTOTAL label)
  ✅ [FIX-E]  Company name footer → scaled to 90% of page width
              (5% whitespace margin on each side, centred)
  ✅ [FIX-F]  Company name colour → LGRAY (#D5D5D5, light grey)
"""

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os

# ── Fonts ────────────────────────────────────────────────────────────
FONT_DIR = "/usr/share/fonts/truetype/liberation/"
pdfmetrics.registerFont(TTFont("IGZ_Reg",    FONT_DIR + "LiberationSans-Regular.ttf"))
pdfmetrics.registerFont(TTFont("IGZ_Bold",   FONT_DIR + "LiberationSans-Bold.ttf"))
pdfmetrics.registerFont(TTFont("IGZ_Serif",  FONT_DIR + "LiberationSerif-Regular.ttf"))
pdfmetrics.registerFont(TTFont("IGZ_SerifB", FONT_DIR + "LiberationSerif-Bold.ttf"))

# ── Colours ───────────────────────────────────────────────────────────
WHITE  = colors.HexColor("#FFFFFF")
BLACK  = colors.HexColor("#000000")
DARK   = colors.HexColor("#1A1A1A")
DGRAY  = colors.HexColor("#333333")
MGRAY  = colors.HexColor("#666666")
LGRAY  = colors.HexColor("#D5D5D5")
XLGRAY = colors.HexColor("#F4F4F4")

# ── Page geometry ─────────────────────────────────────────────────────
W, H  = A4
PAD_L = 18 * mm
PAD_R = 18 * mm
PAD_T = 18 * mm
PAD_B = 18 * mm
ML    = PAD_L
MR    = W - PAD_R
TW    = MR - ML     # ≈ 174 mm


# ─────────────────────────────────────────────────────────────────────
#  Helpers
# ─────────────────────────────────────────────────────────────────────
def _draw_spaced_text(c, x, y, text, font, size, char_space=0):
    """Draw text with per-character extra spacing (wide-spaced title)."""
    c.setFont(font, size)
    if char_space == 0:
        c.drawString(x, y, text)
        return
    cx = x
    for ch in text:
        c.drawString(cx, y, ch)
        cx += pdfmetrics.stringWidth(ch, font, size) + char_space


def _wrap_text(text, font, size, max_w):
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


def _fmt_unit(val):
    """Unit Price column: integer if whole (e.g. 100), else 2dp."""
    if val == int(val):
        return f"{int(val):,}"
    return f"{val:,.2f}"


def _fmt_amount(val):
    """Total / Subtotal / Grand-total column: always 2 dp (e.g. 100.00)."""
    return f"{val:,.2f}"


def _auto_font_size(text, font, target_width, size_min=8, size_max=200):
    """Binary-search the font size so that text width == target_width (pts)."""
    lo, hi = float(size_min), float(size_max)
    for _ in range(40):
        mid = (lo + hi) / 2
        w = pdfmetrics.stringWidth(text, font, mid)
        if w < target_width:
            lo = mid
        else:
            hi = mid
    return (lo + hi) / 2


# ─────────────────────────────────────────────────────────────────────
#  Main draw function
# ─────────────────────────────────────────────────────────────────────
def draw_igz_invoice(c, inv):
    """Draw INFINITE GZ invoice on canvas — v7 (2-fix revision)."""

    # White background
    c.setFillColor(WHITE)
    c.rect(0, 0, W, H, fill=1, stroke=0)

    y = H - PAD_T

    # ════════════════════════════════════════════════════════════════
    #  1. "INVOICE" title (large, extra letter-spaced, regular weight)
    # ════════════════════════════════════════════════════════════════
    INVOICE_SIZE    = 38
    INVOICE_CHAR_SP = 3.2

    c.setFillColor(BLACK)
    invoice_text = "INVOICE"

    spaced_w = sum(
        pdfmetrics.stringWidth(ch, "IGZ_Reg", INVOICE_SIZE) + INVOICE_CHAR_SP
        for ch in invoice_text
    ) - INVOICE_CHAR_SP

    title_y = y - 2 * mm
    _draw_spaced_text(c, ML, title_y, invoice_text, "IGZ_Reg", INVOICE_SIZE, INVOICE_CHAR_SP)

    # Thin horizontal rule from end of title to right margin
    rule_x = ML + spaced_w + 6 * mm
    rule_y = title_y + INVOICE_SIZE * 0.352778 * 0.45
    c.setStrokeColor(LGRAY)
    c.setLineWidth(0.8)
    c.line(rule_x, rule_y, MR, rule_y)

    y -= (INVOICE_SIZE * 0.352778 + 16 * mm)

    # ════════════════════════════════════════════════════════════════
    #  2. Info grid: ISSUED TO (left) | INVOICE NO (right)
    # ════════════════════════════════════════════════════════════════
    LEFT_X          = ML
    RIGHT_X         = ML + TW * 0.52
    INLINE_FONT_SZ  = 11
    FIELD_FONT_SZ   = 7.5
    FIELD_ROW       = 6.5 * mm
    VAL_INDENT      = 30 * mm

    c.setFont("IGZ_Bold", INLINE_FONT_SZ)
    c.setFillColor(BLACK)
    lbl_issued = "ISSUED TO :"
    c.drawString(LEFT_X, y, lbl_issued)
    lbl_issued_w = pdfmetrics.stringWidth(lbl_issued, "IGZ_Bold", INLINE_FONT_SZ)
    c.setFont("IGZ_Reg", INLINE_FONT_SZ)
    c.setFillColor(DARK)
    c.drawString(LEFT_X + lbl_issued_w + 2 * mm, y, str(inv.get("bill_to", "")))

    c.setFont("IGZ_Bold", INLINE_FONT_SZ)
    c.setFillColor(BLACK)
    lbl_invno = "INVOICE NO :"
    c.drawString(RIGHT_X, y, lbl_invno)
    lbl_invno_w = pdfmetrics.stringWidth(lbl_invno, "IGZ_Bold", INLINE_FONT_SZ)
    c.setFont("IGZ_Reg", INLINE_FONT_SZ)
    c.setFillColor(DARK)
    c.drawString(RIGHT_X + lbl_invno_w + 2 * mm, y, str(inv.get("inv_no", "")))
    y -= 9 * mm

    field_left = [
        ("COMPANY",    inv.get("company", "–")),
        ("APPROVAL",   inv.get("approval", "–")),
        ("RECEIPT NO", inv.get("receipt_no", "–")),
        ("REF NO",     inv.get("ref_no", "–")),
    ]
    field_right = [
        ("DATE",         inv.get("date", "–")),
        ("DUE DATE",     inv.get("due_date", inv.get("date", "–"))),
        ("PAYMENT TYPE", inv.get("payment_type", "–")),
        ("CARD NO",      inv.get("card_no", "–")),
    ]

    fy = y
    for label, value in field_left:
        c.setFont("IGZ_Reg", FIELD_FONT_SZ)
        c.setFillColor(MGRAY)
        c.drawString(LEFT_X, fy, label + ":")
        c.setFillColor(BLACK)
        c.drawString(LEFT_X + VAL_INDENT, fy, str(value))
        fy -= FIELD_ROW

    fy2 = y
    for label, value in field_right:
        c.setFont("IGZ_Reg", FIELD_FONT_SZ)
        c.setFillColor(MGRAY)
        c.drawString(RIGHT_X, fy2, label + ":")
        c.setFillColor(BLACK)
        c.drawString(RIGHT_X + VAL_INDENT, fy2, str(value))
        fy2 -= FIELD_ROW

    y = min(fy, fy2) - 8 * mm

    # ════════════════════════════════════════════════════════════════
    #  3. PAY TO section
    # ════════════════════════════════════════════════════════════════
    c.setFont("IGZ_Bold", 8.5)
    c.setFillColor(BLACK)
    c.drawString(LEFT_X, y, "PAY TO:")
    y -= 5.5 * mm

    pay_lines = [
        "Hong Leong Bank",
        "Account Name: INFINITE GZ SDN BHD",
        "Account No :  23600594645",
    ]
    for pl in pay_lines:
        c.setFont("IGZ_Reg", 8)
        c.setFillColor(DARK)
        c.drawString(LEFT_X, y, pl)
        y -= 5 * mm

    y -= 5 * mm

    # ════════════════════════════════════════════════════════════════
    #  4. Table
    # ════════════════════════════════════════════════════════════════
    C_DESC  = LEFT_X
    C_UP_R  = MR - 80 * mm
    C_QTY_C = MR - 47 * mm
    C_TOT_R = MR
    DESC_W  = C_UP_R - 32 * mm - C_DESC

    # Table header labels (no top border line — FIX-1 from v4 kept)
    y -= 7 * mm
    c.setFont("IGZ_Reg", 8)
    c.setFillColor(DARK)
    c.drawString(C_DESC, y, "DESCRIPTION")
    c.drawRightString(C_UP_R, y, "UNIT PRICE")
    c.drawCentredString(C_QTY_C, y, "QTY")
    c.drawRightString(C_TOT_R, y, "TOTAL")
    y -= 3 * mm

    # Single thin bottom-of-header line
    c.setStrokeColor(DARK)
    c.setLineWidth(0.8)
    c.line(ML, y, MR, y)
    y -= 7 * mm

    # ── Item rows ────────────────────────────────────────────────────
    ROW_LINE_H = 6.0 * mm

    subtotal = 0
    items = inv.get("items", [])
    for item in items:
        desc_lines = _wrap_text(item["desc"], "IGZ_Reg", 8, DESC_W)
        row_h = max(7 * mm, len(desc_lines) * ROW_LINE_H + 3 * mm)

        ty = y - 2 * mm
        c.setFont("IGZ_Reg", 8)
        c.setFillColor(DARK)
        for li, dl in enumerate(desc_lines):
            c.drawString(C_DESC, ty - li * ROW_LINE_H, dl)

        # ── [FIX-A] Unit Price: integer display (100)
        if item.get("unit_price") is not None:
            c.drawRightString(C_UP_R, ty, _fmt_unit(item["unit_price"]))

        if item.get("qty") is not None:
            c.drawCentredString(C_QTY_C, ty, str(item["qty"]))

        # ── [FIX-A] Total column: always 2 dp (100.00)
        if item.get("amount") is not None:
            c.drawRightString(C_TOT_R, ty, _fmt_amount(item["amount"]))
            subtotal += item["amount"]

        y -= row_h
        # No row separator lines (FIX-2 from v4 kept)

    y -= 3 * mm

    # Bottom table border
    c.setStrokeColor(LGRAY)
    c.setLineWidth(0.5)
    c.line(ML, y + 1.5 * mm, MR, y + 1.5 * mm)
    y -= 4 * mm

    # ════════════════════════════════════════════════════════════════
    #  5. Totals block
    # ════════════════════════════════════════════════════════════════
    SUBTOT_LABEL_X = C_DESC   # left-aligned (FIX-3 from v4 kept)
    SUBTOT_FONT_SZ = 8

    # ── [FIX-A] SUBTOTAL row: always RM xxx.xx
    c.setFont("IGZ_Reg", SUBTOT_FONT_SZ)
    c.setFillColor(DARK)
    c.drawString(SUBTOT_LABEL_X, y, "SUBTOTAL")
    c.drawRightString(MR, y, f"RM {_fmt_amount(subtotal)}")
    y -= 6.5 * mm

    # ── [FIX-D] Tax row — same font size as SUBTOTAL (SUBTOT_FONT_SZ = 8pt)
    #             Label colour → DARK (same visual weight as SUBTOTAL label)
    c.setFont("IGZ_Reg", SUBTOT_FONT_SZ)   # explicitly same as SUBTOTAL
    c.setFillColor(DARK)                    # DARK same as SUBTOTAL label
    c.drawString(SUBTOT_LABEL_X, y, "TAX :")
    c.setFillColor(DARK)
    c.drawRightString(MR, y, inv.get("tax", "-"))
    y -= 7 * mm

    # ════════════════════════════════════════════════════════════════
    #  [FIX-B] TOTAL black box — equal PAD_ALL=3mm on ALL 4 sides
    # ════════════════════════════════════════════════════════════════
    TOTAL_FONT_SZ = 13
    total         = inv.get("total", subtotal)

    # ── [FIX-A] Grand-total label: always RM xxx.xx
    total_str     = f"TOTAL : RM {_fmt_amount(total)}"

    # Measure exact text width (in ReportLab points)
    txt_w = pdfmetrics.stringWidth(total_str, "IGZ_Bold", TOTAL_FONT_SZ)

    # ── [FIX-B] Equal padding on all four sides
    PAD_ALL = 3 * mm   # same value for left / right / top / bottom

    # Approximate font metrics (Liberation Sans Bold, 13pt)
    # 1 pt = 0.352778 mm  →  ascent ≈ 72% of font size; descent ≈ 22%
    _pt   = 0.352778          # mm per point (for metric calcs only)
    _asc  = TOTAL_FONT_SZ * 0.72 * _pt * mm   # ≈ in ReportLab pts
    _desc = TOTAL_FONT_SZ * 0.22 * _pt * mm   # ≈ in ReportLab pts
    # (mm * mm cancels → keep in pts: 1mm = 1*mm in RL units)
    # actually _asc = 13 * 0.72 * 0.352778 mm converted to pts
    # = 13 * 0.72 * 0.352778 * 2.834645  pts ≈ 9.37 pts
    # Let's just use pts directly:
    _asc_pt  = TOTAL_FONT_SZ * 0.72   # pts
    _desc_pt = TOTAL_FONT_SZ * 0.22   # pts

    BOX_H = _asc_pt + _desc_pt + 2 * PAD_ALL   # total box height (pts)
    BOX_W = txt_w   + 2 * PAD_ALL               # total box width  (pts)
    BOX_X = MR - BOX_W                          # left edge, flush with right margin

    # Text baseline sits PAD_ALL below the box top
    # box top  = text_y + _asc_pt + PAD_ALL  →  text_y = box_top - _asc_pt - PAD_ALL
    # box bottom = box_top - BOX_H
    # We start from current y (top of "row") → define box_top = y
    box_top    = y
    box_bottom = box_top - BOX_H
    text_y     = box_top - PAD_ALL - _asc_pt   # baseline

    # Draw black rect (rounded corners — FIX-G)
    CORNER_R = 3 * mm   # 3mm corner radius
    c.setFillColor(BLACK)
    c.roundRect(BOX_X, box_bottom, BOX_W, BOX_H, CORNER_R, fill=1, stroke=0)

    # White text inside box
    c.setFont("IGZ_Bold", TOTAL_FONT_SZ)
    c.setFillColor(WHITE)
    c.drawString(BOX_X + PAD_ALL, text_y, total_str)

    # Advance y below the box
    y = box_bottom - 6 * mm

    # ════════════════════════════════════════════════════════════════
    #  6. Footer
    # ════════════════════════════════════════════════════════════════
    COMPANY_Y  = PAD_B + 16 * mm
    SUBTEXT_Y  = PAD_B + 8  * mm

    # ── [FIX-E] Company name scaled to 90% of page width
    # 5% whitespace margin on each side → text fills centre 90%
    # ── [FIX-F] Company name colour → LGRAY (light grey)
    COMPANY_TEXT    = "INFINITE GZ SDN BHD"
    TARGET_TXT_W    = W * 0.90   # 90% of A4 page width (5% margin each side)
    company_font_sz = _auto_font_size(COMPANY_TEXT, "IGZ_Bold", TARGET_TXT_W)

    c.setFont("IGZ_Bold", company_font_sz)
    c.setFillColor(LGRAY)        # light grey (#D5D5D5)
    c.drawCentredString(W / 2, COMPANY_Y, COMPANY_TEXT)

    # Registration + email (small, unchanged)
    c.setFont("IGZ_Reg", 7.5)
    c.setFillColor(MGRAY)
    c.drawCentredString(W / 2, SUBTEXT_Y,
        "202401019141 (1564990-X)   Email: business@infinite-gz.com")

    # Safety check
    if y < PAD_B + 32 * mm:
        print(f"  [WARN] {inv.get('inv_no', '')} content too long, may overlap footer!")
    else:
        print(f"  [OK]   IGZ {inv.get('inv_no', '')}  clearance = {(y - PAD_B - 32*mm)/mm:.1f}mm")


# ─────────────────────────────────────────────────────────────────────
#  Public API
# ─────────────────────────────────────────────────────────────────────
def make_igz(inv, path):
    """Generate one INFINITE GZ invoice PDF."""
    c = canvas.Canvas(path, pagesize=A4)
    c.setTitle(f"Invoice {inv['inv_no']}")
    draw_igz_invoice(c, inv)
    c.showPage()
    c.save()
    return path


# ─────────────────────────────────────────────────────────────────────
#  Self-test
# ─────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    import subprocess

    test_inv = {
        "inv_no":       "IGZ_0001",
        "date":         "30 March 2026",
        "due_date":     "30 March 2026",
        "bill_to":      "John Doe",
        "company":      "Example Sdn Bhd",
        "receipt_no":   "000001",
        "approval":     "ABC123",
        "ref_no":       "987654321012",
        "payment_type": "CREDIT",
        "card_no":      "4111 11** **** 1111",
        "tax":          "-",
        "total":        700.00,
        "items": [
            {"desc": "Brand consultation",     "unit_price": 100.00, "qty": 1, "amount": 100.00},
            {"desc": "Logo design",            "unit_price": 100.00, "qty": 1, "amount": 100.00},
            {"desc": "Website design",         "unit_price": 100.00, "qty": 1, "amount": 100.00},
            {"desc": "Social media templates", "unit_price": 100.00, "qty": 1, "amount": 100.00},
            {"desc": "Brand photography",      "unit_price": 100.00, "qty": 1, "amount": 100.00},
            {"desc": "Brand guide",            "unit_price": 100.00, "qty": 1, "amount": 100.00},
            {"desc": "Social media templates", "unit_price": 100.00, "qty": 1, "amount": 100.00},
        ],
    }

    OUT = "/mnt/user-data/outputs"
    os.makedirs(OUT, exist_ok=True)
    pdf_path = f"{OUT}/IGZ_v7_preview.pdf"
    png_path = f"{OUT}/preview_igz_v7.png"

    make_igz(test_inv, pdf_path)
    print(f"PDF saved: {pdf_path}")

    subprocess.run(
        ["pdftoppm", "-r", "180", "-png", "-singlefile",
         pdf_path, png_path.replace(".png", "")],
        check=True
    )
    print(f"Preview PNG: {png_path}")
