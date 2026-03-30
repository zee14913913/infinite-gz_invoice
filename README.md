# 🧾 INFINITE GZ & AI SMART TECH — Invoice Management System

> **Internal Invoice Generation & Record Management System**  
> INFINITE GZ SDN BHD (202401019141) · AI SMART TECH SDN BHD (202401043356)

---

## ✨ Features

| Feature | Description |
|---------|-------------|
| 🏢 **Dual Company** | Switch between INFINITE GZ and AI SMART TECH |
| 📸 **OCR Parsing** | Upload receipt image → auto-extract all fields |
| 📋 **Excel Summary** | Every invoice auto-saved to company Excel workbook |
| 📄 **PDF Invoice** | One-click A4 PDF generation matching company branding |
| 🔽 **Dropdowns** | Payment Type, Card Type, Product Item with 2026 price list |
| 🔗 **Link Tracking** | Receipt Link & Invoice Link stored per row |
| 🖥️ **Horizontal Scroll** | Wide Excel table with scrollable link columns |

---

## 📦 System Files

```
├── app.py                      ← Main Streamlit web app (v3)
├── invoice_infinitegz_v7.py    ← INFINITE GZ PDF engine (v7)
├── invoice_ast.py              ← AI SMART TECH PDF engine
├── requirements.txt            ← Python dependencies
├── font_setup.py               ← Font path auto-patcher (run once)
├── .streamlit/config.toml      ← Streamlit theme config
└── assets/
    └── ai_smart_tech_invoice.png   ← AST invoice style reference
```

---

## 🚀 Quick Start

### 1. Clone the repository

```bash
git clone https://github.com/YOUR_USERNAME/YOUR_REPO_NAME.git
cd YOUR_REPO_NAME
```

### 2. Install Python dependencies

```bash
pip install -r requirements.txt
```

### 3. Install Tesseract OCR

**Windows:**
- Download installer: https://github.com/UB-Mannheim/tesseract/wiki
- Install to default path `C:\Program Files\Tesseract-OCR\`

**macOS:**
```bash
brew install tesseract
```

**Linux (Ubuntu/Debian):**
```bash
sudo apt-get install tesseract-ocr fonts-liberation
```

### 4. Install Liberation fonts & auto-patch

```bash
python font_setup.py
```

This will:
- Detect your OS font directory
- Verify Liberation Sans/Serif fonts exist
- Auto-patch `FONT_DIR` in the invoice engines

### 5. Run the system

```bash
streamlit run app.py
```

Open browser at → **http://localhost:8501**

---

## 🧭 How to Use

```
[Launch] → [Select Company] → [Upload Receipt]
       → [OCR Auto-Parse] → [Review & Edit]
       → [Select Product Item] → [GENERATE INVOICE]
       → [Download PDF] → [View Excel Summary]
```

### Step-by-step:

1. **Select Company** — Choose INFINITE GZ (black) or AI SMART TECH (purple)
2. **Upload Receipt** — Drag & drop PNG/JPG receipt image
3. **OCR Parsing** — System auto-fills all fields (Invoice No, Date, Amount, etc.)
4. **Review Fields** — Only manual entry needed: **Company name**
5. **Select Product Item** — Pick from dropdown (IGZ: 51 types | AST: 37 services)
6. **Generate Invoice** — Click the big button → PDF created
7. **Download** — Download PDF + Excel summary updated automatically

---

## 📊 Excel Structure

Both companies have **22-column** Excel workbooks:

| Column | Type |
|--------|------|
| Invoice No | Auto-parsed (INV- prefix stripped) |
| Date / Due Date | Auto-parsed |
| Bill To | Auto-parsed |
| **Company** | ⭐ Manual entry only |
| Payment Type | Dropdown: Cash/Visa/Master/Amex/Transfer/E-Wallet |
| Card Type | Dropdown: Credit/Debit |
| Product Item | Dropdown: 51 types (IGZ) or 37 services (AST) |
| Qty | Number |
| Unit Price (RM) | Auto-filled from product catalogue |
| Total Amount (RM) | Auto-calculated = Qty × Unit Price |
| Subtotal / Tax / Total | Auto-parsed |
| Receipt Link | Stored file path / URL |
| Invoice Link | Generated PDF path |

---

## 🏷️ IGZ Product Catalogue (51 Transaction Types)

From `IGZ Transaction Reference v1.0`:

| Group | Category | Count |
|-------|----------|-------|
| **A** | Financing Arrangement | 8 |
| **B** | Record & Report Handling | 8 |
| **C** | Digital & Business Services | 8 |
| **D** | Outcome Sharing | 12 |
| **E** | Bank Account Operations | 14 |
| **U** | Uncertain / Pending | 1 |

## 🤖 AST Product Catalogue (37 AI Services)

2026 Malaysia market prices covering:
Website Dev · Mobile App · Software Dev · CRM · AI Automation · Hardware · Marketing · Technical Support

---

## 🎨 Invoice Design

| Company | Style | Colors |
|---------|-------|--------|
| INFINITE GZ | Black minimal, large INVOICE title, rounded TOTAL box | Black `#000000` + Dark `#1A1A1A` |
| AI SMART TECH | Purple corporate, numbered table, promo rebate row | Purple `#6B5B95` |

---

## ⚙️ Requirements

- Python 3.9+
- Tesseract OCR 5.x
- Liberation Sans/Serif fonts
- See `requirements.txt` for Python packages

---

## 📁 Output Directory

At runtime the system creates:
```
output/
├── invoices/     ← Generated PDF invoices
├── receipts/     ← Uploaded receipt images
├── IGZ_invoices.xlsx
└── AST_invoices.xlsx
```

---

*INFINITE GZ SDN BHD · Internal Use Only · Confidential*
