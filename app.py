r"""
سیستم تطبیق فیش‌های بانکی
===========================================
FastAPI backend + pdfplumber + pandas

نحوه اجرا:
  cd d:\Checkhesab\receipt-checker
  python app.py

سپس مرورگر را باز کنید:
  http://localhost:8765
"""

from __future__ import annotations
import io, re, json, logging
from pathlib import Path
from typing import Optional

import uvicorn
import pdfplumber
import pandas as pd
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from bs4 import BeautifulSoup

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("receipt_checker")

app = FastAPI(title="Receipt Checker")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

# ──────────────────────────────────────────────────────────────────────────────
# Number utilities (Persian/Arabic/English digits + comma removal)
# ──────────────────────────────────────────────────────────────────────────────
FA_DIGITS = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")

# Codes that identify the account holder themselves (not tracking codes)
# These are extracted from the PDF header and should be excluded from per-row matching
ACCOUNT_HOLDER_CODES: set[str] = set()

def to_num(s: str) -> Optional[float]:
    """Convert a possibly-Persian number string to float. Returns None on failure."""
    if not s:
        return None
    s_str = str(s).translate(FA_DIGITS).replace(",", "").replace("،", "").replace(" ", "").strip()
    if not s_str:
        return None
    try:
        return float(s_str)
    except ValueError:
        return None

def clean_str(s) -> str:
    if s is None:
        return ""
    return str(s).strip()

def nrm(s: str) -> str:
    """Normalize Persian/Arabic string for comparison."""
    if not s:
        return ""
    s = s.replace("\u200c", " ").replace("\u200d", " ")  # ZWNJ
    s = s.replace("ك", "ک").replace("ي", "ی")             # Arabic variants
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s

# ──────────────────────────────────────────────────────────────────────────────
# PDF Parser — pdfplumber table extraction
# ──────────────────────────────────────────────────────────────────────────────

# Column header patterns (right→left reading order in visual PDF)
# The PDF has these columns (left side in x-coordinate space → right side visually):
#   مانده ریالی | بستانکار مالی | بدهکار مالی | مانده طلایی | بستانکار طلا | بدهکار طلا
#   | شرح | شماره سند | تاریخ | حساب
COLUMN_PATTERNS = {
    "حساب":          re.compile(r"حساب"),
    "تاریخ":         re.compile(r"تاریخ"),
    "شماره_سند":     re.compile(r"شماره\s*سند"),
    "شرح":           re.compile(r"^شرح$"),
    "بدهکار_طلا":    re.compile(r"بدهکار\s*طلا"),
    "بستانکار_طلا":  re.compile(r"بستانکار\s*طلا"),
    "مانده_طلا":     re.compile(r"مانده\s*طلا"),
    "بدهکار_مالی":   re.compile(r"بدهکار\s*مالی"),
    "بستانکار_مالی": re.compile(r"بستانکار\s*مالی"),
    "مانده_ریالی":   re.compile(r"مانده\s*ریالی"),
}

def identify_col(header_text: str) -> str:
    """Map a header cell text to a column key."""
    t = clean_str(header_text)
    for key, pat in COLUMN_PATTERNS.items():
        if pat.search(t):
            return key
    return ""

import fitz  # PyMuPDF

def fix_rtl(text: str) -> str:
    """Reverse extracted RTL text and restore LTR layout for numbers and English words."""
    if not text: return ""
    rev = text[::-1]
    return re.sub(r'[A-Za-z0-9/\.,\-_\[\]\(\)]+', lambda m: m.group(0)[::-1], rev)

def parse_pdf(pdf_bytes: bytes) -> list[dict]:
    """Extract rows from the Tahesab account statement PDF using visual word assembly."""
    global ACCOUNT_HOLDER_CODES
    rows_out = []
    ACCOUNT_HOLDER_CODES = set()

    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    logger.info(f"PDF pages: {len(doc)}")
    
    for page_num, page in enumerate(doc, 1):
        words = page.get_text("words")
        
        # Extract account holder code from page 1 header (format: Name)CODE( AccountNumber in RTL)
        if page_num == 1:
            for w in words:
                text = w[4]
                # PyMuPDF often reverses RTL text: so (8181) appears as )8181(
                m = re.match(r"^\)(\d{4,5})\($", text)  # e.g. ")8181("
                if not m:
                    m = re.match(r"^\((\d{4,5})\)$", text)  # also check normal form
                if m:
                    ACCOUNT_HOLDER_CODES.add(m.group(1))
                    logger.info(f"Account holder code detected: {m.group(1)}")
        
        rows_map = {}
        
        # 1. Group words by visual Y-row
        for w in words:
            x0, y0, x1, y1, text, b, l, w_num = w
            y_mid = (y0 + y1) / 2
            y_key = round(y_mid / 5) * 5
            if y_key not in rows_map:
                rows_map[y_key] = []
            rows_map[y_key].append((x0, text))

        # 2. Assemble lines by sorting right-to-left (X descending)
        sorted_y = sorted(rows_map.keys())
        for y in sorted_y:
            # sort words by X right-to-left
            row_words = sorted(rows_map[y], key=lambda item: item[0], reverse=True)
            line = " ".join(item[1] for item in row_words).strip()
            
            if not line:
                continue

            # --- Apply Regex Logic to Assembled Line ---
            line_norm = line.translate(FA_DIGITS)
            line_no_spaces = re.sub(r"\s+", "", line_norm)
            
            # Identify valid row: must have a date-like string or a long dash-surrounded document number
            # the garbled dates look like 1/11//1/1 sometimes, but amounts are the real indicator
            
            # Amounts: match sequences of digits with standard commas/spaces
            # Note: PyMuPDF maps the zero glyph to '2' in this PDF font
            # So we look at the raw extracted digits and fix trailing 222... groups
            amounts = []
            for m in re.finditer(r"\b(\d{1,3}(?:[,،./]\d{3})+|\d{5,})\b", line_norm):
                raw = m.group(1).replace(",", "").replace("،", "").replace(".", "").replace("/", "")
                # Fix trailing run of 2s that is actually zeros (PDF font encoding bug)
                # Only replace '22...' at the trailing end if they form groups of 2+
                # Be conservative: only replace trailing 2s if the number has > 6 digits
                fixed = raw
                if len(raw) > 6 and raw.endswith("22"):
                    # Count trailing 2s
                    trailer = len(raw) - len(raw.rstrip("2"))
                    if trailer >= 2:
                        fixed = raw[:-trailer] + "0" * trailer
                amt = to_num(fixed)
                if amt and amt > 1000 and len(fixed) <= 9:  # exclude 10+ digit account numbers
                    amounts.append(amt)
            
            amounts.sort(reverse=True)
            
            # If no large amounts, it's not a transaction data row
            if not amounts:
                continue
                
            # Usually credit is the 2nd largest number, or 1st if there's only one
            credit = amounts[1] if len(amounts) >= 2 else amounts[0]
            
            codes, sender = parse_desc(line)
            
            desc_fixed = line[:200].strip()

            doc_m = re.search(r"\b(\d{6,8})\b", line_norm)
            doc_num = doc_m.group(1) if doc_m else ""
            
            # Since dates can be weird, we just match a naive pattern or leave empty
            date_m = re.search(r"1[34]\d\d[\D]\d{1,2}[\D]\d{1,2}", line_no_spaces)
            date_str = date_m.group(0).replace('-', '/').replace('_', '/') if date_m else ""

            if credit > 0:
                rows_out.append({
                    "page":        page_num,
                    "date":        date_str,
                    "doc_num":     doc_num,
                    "desc":        desc_fixed,
                    "credit":      credit,
                    "debit":       0,
                    "credit_raw":  str(credit),
                    "debit_raw":   "0",
                    "codes":       codes,
                    "sender":      sender,
                    "doc_type":    "بستانکار",
                    "amount":      credit,
                })

    doc.close()
    logger.info(f"Total PDF rows extracted: {len(rows_out)}")
    
    # Auto-detect and filter codes that appear in >30% of all rows
    # These are "owner codes" (like account branch codes) that appear everywhere
    if rows_out:
        code_freq: dict[str, int] = {}
        for r in rows_out:
            for c in r.get("codes", []):
                code_freq[c] = code_freq.get(c, 0) + 1
        threshold = len(rows_out) * 0.1
        auto_owner_codes = {c for c, cnt in code_freq.items() if cnt > threshold}
        if auto_owner_codes:
            logger.info(f"Auto-detected owner codes (filtered from tracking): {auto_owner_codes}")
            ACCOUNT_HOLDER_CODES.update(auto_owner_codes)
            for r in rows_out:
                r["codes"] = [c for c in r["codes"] if c not in ACCOUNT_HOLDER_CODES]

    return rows_out

def parse_html(html_bytes: bytes) -> list[dict]:
    """Parse HTML statement into the standard dict format."""
    
    # Try decoding as standard UTF-8, fallback to cp1256 (Windows-1256) used by older Iranian accounting tools
    try:
        html_text = html_bytes.decode('utf-8')
    except UnicodeDecodeError:
        html_text = html_bytes.decode('cp1256', errors='replace')
        
    soup = BeautifulSoup(html_text, "html.parser")
    rows_out = []
    seen_tx = set()
    
    
    # Try to find exactly which columns contain the desc/credit/debit
    # by looking at table headers
    desc_idx = credit_idx = debit_idx = date_idx = -1
    
    tables = soup.find_all("table")
    for table in tables:
        headers = table.find_all(["th", "td"])
        for i, h in enumerate(headers):
            text = h.get_text(strip=True)
            if "شرح" in text: desc_idx = i
            elif "بستانکار" in text: credit_idx = i
            elif "بدهکار" in text: debit_idx = i
            elif "تاریخ" in text: date_idx = i
            
        # If we found at least desc and credit in this table, process its rows
        if desc_idx != -1 and credit_idx != -1:
            trs = table.find_all("tr")
            for tr in trs:
                tds = tr.find_all("td")
                if len(tds) <= max(desc_idx, credit_idx):
                    continue
                    
                desc = " | ".join(td.get_text(strip=True) for td in tds if td.get_text(strip=True))
                # Ignore header rows that got caught as td
                if "شرح" in desc or "ردیف" in desc or not desc:
                    continue
                    
                credit_raw = tds[credit_idx].get_text(strip=True)
                debit_raw = tds[debit_idx].get_text(strip=True) if debit_idx != -1 and len(tds) > debit_idx else "0"
                date_str = tds[date_idx].get_text(strip=True) if date_idx != -1 and len(tds) > date_idx else ""
                
                credit = to_num(credit_raw) or 0
                debit = to_num(debit_raw) or 0
                
                if credit == 0 and debit == 0:
                    continue # Not a transaction row
                    
                tx_key = (date_str, desc[:100], credit, debit)
                if tx_key in seen_tx:
                    continue
                seen_tx.add(tx_key)
                    
                codes, sender = parse_desc(desc)
                
                rows_out.append({
                    "page":        1,
                    "date":        date_str,
                    "doc_num":     "",
                    "desc":        desc[:200],
                    "credit":      credit,
                    "debit":       debit,
                    "credit_raw":  credit_raw,
                    "debit_raw":   debit_raw,
                    "codes":       codes,
                    "sender":      sender,
                    "doc_type":    "بستانکار" if credit > 0 else "بدهکار",
                    "amount":      credit if credit > 0 else debit,
                })
            # Do not break here! Process all tables because HTML might be paginated
            # with multiple transaction tables.
            pass
                
    logger.info(f"Total HTML rows extracted: {len(rows_out)}")
    
    # Auto-detect owner codes as we do for PDF
    if rows_out:
        code_freq: dict[str, int] = {}
        for r in rows_out:
            for c in r.get("codes", []):
                code_freq[c] = code_freq.get(c, 0) + 1
        threshold = len(rows_out) * 0.1
        auto_owner_codes = {c for c, cnt in code_freq.items() if cnt > threshold}
        if auto_owner_codes:
            global ACCOUNT_HOLDER_CODES
            ACCOUNT_HOLDER_CODES.update(auto_owner_codes)
            for r in rows_out:
                r["codes"] = [c for c in r["codes"] if c not in ACCOUNT_HOLDER_CODES]

    return rows_out

def parse_desc(desc: str) -> tuple[list[str], str]:
    """
    Extract 4-digit tracking codes and sender name from شرح column.

    Examples:
      "واریز نقد به بانک (از مشتری) [صادرات منشادی،24454] خدمتی/2452"
      "واریز نقد به بانک (از مشتری) [100617] واحدشه/1264"
      "واریز نقد به بانک (از مشتری) [21102] برهام/1102"
      "[97830/3882] دادور/2951"
    """
    if not desc:
        return [], ""

    # Normalize Persian/Arabic digits to Latin
    desc_n = desc.translate(FA_DIGITS)

    codes: set[str] = set()

    # Pattern 1: فارسی/XXXX  — tracking code after sender name + slash
    # Pattern 1a: فارسی/XXXX — sender name + slash + tracking code
    for m in re.finditer(r"[\u0600-\u06FF]+\s*/\s*(\d{4,8})\b", desc_n):
        codes.add(m.group(1))

    # Pattern 1b: XXXX/فارسی  — tracking code before sender name + slash
    for m in re.finditer(r"\b(\d{4})\s*/\s*[\u0600-\u06FF]+", desc_n):
        codes.add(m.group(1))

    # Pattern 2: NNNN[اسم فارسی] — 4-digit code BEFORE a bracket (e.g. 8842[غفاری])
    for m in re.finditer(r"(?<!\d)(\d{4})(?!\d)\s*\[", desc_n):
        codes.add(m.group(1))

    # Pattern 3: inside brackets [بانک،NNNNN]
    for m in re.finditer(r"\[([^\]]+)\]", desc_n):
        inner = m.group(1)
        nums = re.findall(r"\d+", inner)
        for n in nums:
            if len(n) >= 4:
                codes.add(n)   # Keep full number

    # Pattern 4: code AFTER closing bracket ] (PyMuPDF reverses RTL brackets)
    # Format: "[ بانک منشادی ، ]3030," → code is 3030 (right after ])
    # Only match if the bracket content is long (bank name), NOT short person names
    for m in re.finditer(r"\[([^\]]{8,})\]\s*(\d{4,8})\b", desc_n):
        codes.add(m.group(2))

    # Pattern 5: Isolated 4 to 8 digit numbers surrounded by | spaces |
    for m in re.finditer(r"\|\s*(\d{4,8})\s*(?=\|)", desc_n):
        codes.add(m.group(1))

    # Pattern 6: Isolated 4 to 8 digit numbers at the end of the line
    for m in re.finditer(r"\|\s*(\d{4,8})\s*$", desc_n):
        codes.add(m.group(1))

    # Pattern 7: Multiple 4-8 digit numbers separated by slashes/dashes (like 0260/2502)
    for m in re.finditer(r"\b\d{4,8}(?:\s*[/,-]\s*\d{4,8})+\b", desc_n):
        match_str = m.group(0)
        # Exclude dates like 1404/12/04
        if not re.search(r"\b\d{4}/\d{1,2}/\d{1,2}\b", match_str): 
            for n in re.findall(r"\d{4,8}", match_str):
                codes.add(n)

    sender = ""
    # Simplify regex to prevent catastrophic backtracking (ReDoS)
    # Match 1+ Persian words before / 4-digit
    m = re.search(r"([\u0600-\u06FF\s]+)\s*/\s*\d{4}", desc)
    if m:
        sender = m.group(1).strip()
    else:
        m2 = re.search(r"\d{4}\s*/\s*([\u0600-\u06FF\s]+)", desc)
        if m2:
            sender = m2.group(1).strip()
        else:
            # Fallback: any Persian name after last bracket
            m = re.search(r"\]\s*([\u0600-\u06FF\s]+)", desc)
            if m:
                sender = m.group(1).strip()
            else:
                # Fallback: حواله به: Name
                m_havale = re.search(r"حواله\s*به:\s*\d*\s*([\u0600-\u06FF\s]+)", desc)
                if m_havale:
                    sender2 = m_havale.group(1).strip()
                    sender2 = re.sub(r'[^\u0600-\u06FF\s]+$', '', sender2).strip()
                    sender = sender2
                    
                    # Also try to extract the actual sender name appearing BEFORE (حواله به
                    m_before = re.search(r"\|\s*([^\|]+?)\s*\(حواله\s*به:", desc)
                    if m_before:
                        sender1 = m_before.group(1).strip()
                        sender1 = re.sub(r'[^\u0600-\u06FF\s]+', '', sender1).strip()
                        if sender1:
                            sender = f"{sender1} {sender2}"

    # Filter out codes that belong to the account holder (not tracking codes)
    codes -= ACCOUNT_HOLDER_CODES

    return sorted(codes), sender

# ──────────────────────────────────────────────────────────────────────────────
# Excel Parser
# ──────────────────────────────────────────────────────────────────────────────

def parse_excel(excel_bytes: bytes, filename: str) -> list[dict]:
    """
    Parse bank Excel statement. Detects header row automatically.
    Returns list of transactions with: ref, last4, amount, date, desc, sender.
    """
    # Detect engine from extension
    ext = Path(filename).suffix.lower()
    engine = "xlrd" if ext == ".xls" else "openpyxl"

    try:
        df_raw = pd.read_excel(io.BytesIO(excel_bytes), engine=engine, header=None,
                               dtype=str, na_filter=False)
    except Exception as e:
        raise HTTPException(400, f"خطا در خواندن Excel: {e}")

    logger.info(f"Excel raw shape: {df_raw.shape}")

    # Find header row
    header_kw = {
        "ref":    re.compile(r"مرجع|پیگیری|reference|شناسه|ارجاع|trace|شماره\s*سند", re.I),
        "credit": re.compile(r"بستانکار|واریز|credit|دریافتی", re.I),
        "debit":  re.compile(r"بدهکار|برداشت|debit|پرداختی", re.I),
        "amount": re.compile(r"مبلغ|amount", re.I),
        "date":   re.compile(r"تاریخ|date", re.I),
        "desc":   re.compile(r"شرح|توضیح|description|بابت|نوع", re.I),
        "sender": re.compile(r"نام|واریزکننده|صاحب|sender", re.I),
    }

    # Bank Melli often uses "واریز" for amount, "شماره سند" for ref, "شرح" for desc
    
    header_idx = -1
    col_map: dict[str, int] = {}

    for ri in range(min(20, len(df_raw))):
        row = df_raw.iloc[ri]
        hits, tmp = 0, {}
        for ci, cell in enumerate(row):
            c = str(cell).strip()
            # Melli specific columns
            if re.search(r"بستانکار|واریز|credit|دریافتی", c, re.I) and "credit" not in tmp:
                tmp["credit"] = ci; hits += 1
            elif re.search(r"بدهکار|برداشت|debit|پرداختی", c, re.I) and "debit" not in tmp:
                tmp["debit"] = ci; hits += 1
            elif re.search(r"مبلغ|amount", c, re.I) and "amount" not in tmp:
                tmp["amount"] = ci; hits += 1
            elif re.search(r"تاریخ|date", c, re.I) and "date" not in tmp:
                tmp["date"] = ci; hits += 1
            elif re.search(r"شرح|توضیح|description|بابت|نوع", c, re.I) and "desc" not in tmp:
                tmp["desc"] = ci; hits += 1
            elif re.search(r"شماره\s*سند|مرجع|پیگیری|reference|شناسه|ارجاع|trace", c, re.I) and "ref" not in tmp:
                tmp["ref"] = ci; hits += 1
                
        if hits >= 3:
            header_idx = ri; col_map = tmp
            logger.info(f"Excel header at row {ri}: {col_map}")
            break

    txns = []

    if header_idx >= 0:
        for ri in range(header_idx + 1, len(df_raw)):
            row = df_raw.iloc[ri]
            def g(k): return clean_str(row.iloc[col_map[k]]) if k in col_map else ""
            
            ref  = g("ref")
            date = g("date")
            desc = g("desc")
            
            # Extract sender from description for Melli bank statements
            # Typically something like "-0412060171037205-ملي-حسن-زارع  نريماني"
            sndr = ""
            m = re.search(r"[-\s]([^\-\s\d]{3,}(?:[\s\-][^\-\s\d]{3,})*)$", desc)
            if m:
                sndr = m.group(1).replace("-", " ").strip()
            else:
                sndr = desc
            
            amt = 0
            tx_type = "unknown"
            
            # Determine credit/debit
            credit_amt = to_num(g("credit")) or 0 if "credit" in col_map else 0
            debit_amt = to_num(g("debit")) or 0 if "debit" in col_map else 0
            general_amt = to_num(g("amount")) or 0 if "amount" in col_map else 0
            amt_raw = g("amount") or g("credit") or g("debit")
            
            
            if credit_amt > 0:
                amt = credit_amt
                tx_type = "deposit"
            elif debit_amt > 0:
                amt = debit_amt
                tx_type = "withdrawal"
            elif general_amt > 0:
                amt = general_amt
                # Guess based on description if it's a general amount column
                if re.search(r"واریز|بستانکار|دریافت|اعتبار", desc):
                    tx_type = "deposit"
                elif re.search(r"برداشت|بدهکار|پرداخت|خرید", desc):
                    tx_type = "withdrawal"
            
            if not ref and not amt and not date:
                continue
                
            ref_digits = re.sub(r"\D", "", ref)
            desc_digits = re.sub(r"\D", "", str(desc))
            
            # All possible 4+ digit codes from this bank row:
            bank_codes: set[str] = set()
            
            # (1) Ref string (full length if >= 4)
            if len(ref_digits) >= 4:
                bank_codes.add(ref_digits)
            
            # (2) Digits immediately before English letters (e.g. "2337GPPC")
            for m_eng in re.finditer(r"(\d{4,})[A-Za-z]", ref):
                bank_codes.add(m_eng.group(1))
            
            # (3) Codes inside description
            desc_n = str(desc).translate(FA_DIGITS)
            for m_desc in re.finditer(r"\d{10,}", desc_n):
                bank_codes.add(m_desc.group(0))
            for m_desc in re.finditer(r"\b(\d{4,8})\b", desc_n):
                if m_desc.group(1) != str(amt_raw).translate(FA_DIGITS):
                    bank_codes.add(m_desc.group(1))
            
            # Primary reference number
            m_eng = re.search(r"(\d{4,})[A-Za-z]", ref)
            last4 = m_eng.group(1) if m_eng else ref_digits
            
            # Check for duplicate lock
            raw_joined = " | ".join(str(r) for r in row if str(r).strip())
            is_locked = False
            lock_text = ""
            for cell in row:
                if "تطبیق شده" in str(cell):
                    is_locked = True
                    lock_text = str(cell).strip()
                    break
            
            txns.append({
                "row_num":   ri + 1,
                "ref":       ref,
                "last4":     last4,
                "all_codes": sorted(bank_codes),   # list, not set — JSON serializable
                "amount":    amt,
                "tx_type":   tx_type,
                "date":      date,
                "desc":      desc,
                "sender":    sndr,
                "raw":       raw_joined,
                "is_locked": is_locked,
                "lock_text": lock_text,
            })
    else:
        # Auto-detect: scan every row for big numbers
        logger.warning("No header found — auto-detecting Excel rows")
        for ri in range(len(df_raw)):
            row = df_raw.iloc[ri]
            joined = " ".join(str(c) for c in row)
            joined_n = joined.translate(FA_DIGITS)
            refs    = re.findall(r"\b(\d{10,24})\b", joined_n)
            amounts = [float(m) for m in re.findall(r"\b(\d{6,})\b", joined_n) if float(m) > 10000]
            if not refs and not amounts:
                continue
            ref = sorted(refs, key=len, reverse=True)[0] if refs else ""
            amt = sorted(amounts, reverse=True)[0] if amounts else None
            # find Persian name
            sndr_m = re.search(r"[\u0600-\u06FF]{4,}(?:\s[\u0600-\u06FF]{3,})*", joined)
            sndr = sndr_m.group(0) if sndr_m else ""
            
            is_locked = False
            lock_text = ""
            for cell in row:
                if "تطبیق شده" in str(cell):
                    is_locked = True
                    lock_text = str(cell).strip()
                    break
            
            tx_type = "unknown"
            if re.search(r"واریز|بستانکار|دریافت", joined_n): tx_type = "deposit"
            elif re.search(r"برداشت|بدهکار|پرداخت", joined_n): tx_type = "withdrawal"
            
            txns.append({
                "row_num": ri + 1,
                "ref": ref, "last4": ref[-4:] if len(ref) >= 4 else ref,
                "all_codes": [ref] if ref else [],
                "amount": amt, "tx_type": tx_type, "date": "", "desc": joined[:80],
                "sender": sndr,
                "raw": joined[:120],
                "is_locked": is_locked,
                "lock_text": lock_text,
            })

    logger.info(f"Excel transactions: {len(txns)}")
    return txns

# ──────────────────────────────────────────────────────────────────────────────
# Matching Engine
# ──────────────────────────────────────────────────────────────────────────────

def match_receipts(
    pdf_rows: list[dict],
    bank_txns: list[dict],
    credit_only: bool = True,
    use_tracking: bool = True,
    use_name: bool = True,
    use_amount: bool = True,
    tx_type_filter: str = "all",
) -> list[dict]:
    """Match PDF receipt rows against bank transactions."""

    # Filter PDF rows
    rows = [r for r in pdf_rows if not credit_only or (r.get("credit") and r["credit"] > 0)]

    # ── Only keep banking transaction rows ──
    # Banking rows have: bank account prefix, bank keywords, or bracket references
    # Exclude: pure purchase/sale rows (e.g. "حسن زاده NAME[COMPANY] AMOUNT")
    BANK_KEYWORDS = re.compile(
        r"واریز|برداشت|خروج|بانک|حواله|آبشده|انتقال|متفرقه|پایا|ساتنا|شبا|اینترنت"
        r"|دستور|دریافت|پرداخت|نقد|فیش|رمزدار|چک|صادرات|اي.ران زمین"
        r"|چ\s+\d{10}|پ\s+\d{10}|ش\s+\d{10}|س\s+\d{10}|د\s+\d{10}|ي\s+\d{10}"
        r"|\[\s*(?:صادرات|ملي|ملت|تجارت|سامان|پارسيان|پاسارگاد|رسالت|قرض|اي.ران)"
        r"|\bGPPC\b|\bDRPA\b|\bGPAC\b|\bIMPT\b|\bSPAC\b",
        re.IGNORECASE
    )
    # We want to keep rows that either have a banking keyword OR contain an isolated 4+ digit number in the description
    rows = [r for r in rows if BANK_KEYWORDS.search(r.get("desc", "")) or re.search(r"\b\d{4,}\b", r.get("desc", ""))]
    logger.info(f"Banking-type PDF rows after filter: {len(rows)}")

    # Build bank lookup maps
    by_last4: dict[str, list] = {}
    by_amount: dict[str, list] = {}
    by_sender: dict[str, list] = {}

    for tx in bank_txns:
        # Filter by transaction type if specified
        if tx_type_filter != "all" and tx.get("tx_type") != "unknown" and tx.get("tx_type") != tx_type_filter:
            continue
            
        # Index ALL possible codes from this bank row
        all_codes = set(tx.get("all_codes", []))
        if tx.get("last4"):
            all_codes.add(tx["last4"])
            
        for code in all_codes:
            if code:
                by_last4.setdefault(code, []).append(tx)
                # Also index the last 4 digits of any longer code for fallback
                if len(code) > 4:
                    by_last4.setdefault(code[-4:], []).append(tx)
        if tx.get("amount"):
            k = str(int(tx["amount"]))
            by_amount.setdefault(k, []).append(tx)
        for src in [tx.get("sender", ""), tx.get("desc", "")]:
            k = nrm(src)
            if len(k) >= 3:
                by_sender.setdefault(k, []).append(tx)
                
    if bank_txns:
        logger.info(f"Sample bank txns last4s: {[t.get('last4') for t in bank_txns[:10]]}")

    results = []
    
    if rows:
        logger.info(f"Sample PDF row codes: {[r.get('codes') for r in rows[:10]]}")
        
    for r in rows:
        amount  = r.get("credit") or r.get("debit") or 0
        codes   = r.get("codes", [])
        sender  = r.get("sender", "")

        matched = None
        status  = "not_found" # "exact", "review", "not_found"
        method  = ""

        # ── 1. Match by tracking code (Full or Last 4) ──
        if use_tracking:
            for code in codes:
                if len(code) < 4:
                    continue
                
                # Try exact match first, then 4-digit fallback
                search_codes = [code]
                if len(code) > 4:
                    search_codes.append(code[-4:])
                    
                for scode in search_codes:
                    cands = by_last4.get(scode, [])
                    if cands:
                        # Find the best candidate. Prioritize locked candidates so duplicates are correctly flagged even if another unlocked row exists
                        if len(cands) == 1:
                            matched, method, status = cands[0], f"کد: {scode}", "exact"
                        else:
                            # Disambiguate by amount (10% tolerance)
                            amount_cands = [c for c in cands if c.get("amount") and amount and abs(c["amount"] - amount) < amount * 0.10]
                            working_cands = amount_cands if amount_cands else cands
                            
                            # If there is a locked candidate among valid options, choose it so we flag duplicate
                            locked_cand = next((c for c in working_cands if c.get("is_locked")), None)
                            if locked_cand:
                                matched, method, status = locked_cand, f"کد {scode} + مبلغ (تکراری)", "exact"
                            else:
                                matched, method, status = working_cands[0], f"کد {scode} (چندگانه)", "exact"
                        break
                
                if matched:
                    break

        # ── 2. Match by sender name AND amount (Needs Manual Review) ──
        # ONLY IF the PDF receipt actually had a tracking code (user requirement: don't match receipts without codes)
        if not matched and use_name and sender and amount and codes:
            skey = nrm(sender)
            cands = by_sender.get(skey, [])
            if cands:
                # Find the ones that match amount strictly
                amount_cands = [c for c in cands if c.get("amount") and abs(c["amount"] - amount) < amount * 0.05]
                if amount_cands:
                    locked_cand = next((c for c in amount_cands if c.get("is_locked")), None)
                    matched = locked_cand if locked_cand else amount_cands[0]
                    method, status = f"نام مشابه + مبلغ یکسان", "review"

            # Partial name match
            if not matched and len(skey) >= 3:
                for k, txs in by_sender.items():
                    if len(k) >= 3 and (k in skey or skey in k):
                        amount_cands = [c for c in txs if c.get("amount") and abs(c["amount"] - amount) < amount * 0.05]
                        if amount_cands:
                            locked_cand = next((c for c in amount_cands if c.get("is_locked")), None)
                            matched = locked_cand if locked_cand else amount_cands[0]
                            method, status = f"نام جزئی مشابه + مبلغ یکسان", "review"
                            break

        # ── 3. (Removed pure amount fallback to prevent false positives) ──
        
        # ── 4. Override status if this matched row is already locked ──
        if matched and matched.get("is_locked"):
            status = "duplicate"
            method = "تکراری (قبلاً تطبیق شده)"

        results.append({
            "idx":          len(results) + 1,
            "date":         r.get("date", ""),
            "doc_num":      r.get("doc_num", ""),
            "doc_type":     r.get("doc_type", ""),
            "amount":       int(amount) if amount else 0,
            "codes":        ", ".join(codes),
            "sender":       sender,
            "desc":         r.get("desc", ""),
            "found":        matched is not None,
            "status":       status,
            "match_method": method,
            "bank_ref":     matched.get("ref", "") if matched else "",
            "bank_date":    matched.get("date", "") if matched else "",
            "bank_row":     matched.get("row_num", "") if matched else "",
        })

    return results

# ──────────────────────────────────────────────────────────────────────────────
# FastAPI Routes
# ──────────────────────────────────────────────────────────────────────────────

@app.get("/", response_class=HTMLResponse)
async def index():
    html_path = Path(__file__).parent / "index.html"
    return html_path.read_text(encoding="utf-8")

@app.post("/analyze")
async def analyze(
    pdf_file:    UploadFile = File(...),
    excel_file:  UploadFile = File(...),
    credit_only: str        = Form("true"),
    use_tracking:str        = Form("true"),
    use_name:    str        = Form("true"),
    use_amount:  str        = Form("true"),
    tx_type_filter: str     = Form("all"),
):
    pdf_bytes   = await pdf_file.read()
    excel_bytes = await excel_file.read()

    filename = (pdf_file.filename or "").lower()
    is_html = filename.endswith(".html") or filename.endswith(".htm")

    try:
        if is_html:
            pdf_rows = parse_html(pdf_bytes)
        else:
            pdf_rows = parse_pdf(pdf_bytes)
    except Exception as e:
        logger.exception(f"Exception parsing {filename}")
        raise HTTPException(500, f"خطا در پردازش فایل: {e}")

    try:
        bank_txns = parse_excel(excel_bytes, excel_file.filename or "bank.xls")
    except Exception as e:
        logger.exception("Excel parse error")
        raise HTTPException(500, f"خطا در پردازش Excel: {e}")

    results = match_receipts(
        pdf_rows, bank_txns,
        credit_only    = credit_only.lower() == "true",
        use_tracking   = use_tracking.lower() == "true",
        use_name       = use_name.lower()     == "true",
        use_amount     = use_amount.lower()   == "true",
        tx_type_filter = tx_type_filter.lower(),
    )

    # ── Excel Locking Generation ──
    try:
        import pandas as pd
        import openpyxl
        from openpyxl.styles import PatternFill
        import io
        import os

        # Get 1-based row numbers from matched results
        new_matched_rows = set(r["bank_row"] for r in results if r["status"] in ("exact", "review") and r.get("bank_row"))
        all_matched_rows = set(new_matched_rows)
        
        # Keep track of old lock texts to preserve them
        old_locks = {}
        
        # Also include previously locked rows from the bank txns so they don't lose their color
        # because pandas rebuilds the file without preserving original formatting.
        for t in bank_txns:
            if t.get("is_locked") and t.get("row_num"):
                # +1 because pandas to_excel without header/index makes 0-indexed rows 1-indexed in openpyxl,
                # but tx['row_num'] is already aligned to openpyxl 1-based indexing in parse_excel
                all_matched_rows.add(t["row_num"])
                if t.get("lock_text"):
                    old_locks[t["row_num"]] = t["lock_text"]

        if all_matched_rows:
            # 1. Read raw bytes with Pandas and convert to standard XLSX in memory
            # (We do this because python excel styling libs struggle with old formatting bounds)
            df = pd.read_excel(io.BytesIO(excel_bytes), header=None)
            
            # Write to a temporary bytes buffer as .xlsx
            xlsx_buffer = io.BytesIO()
            df.to_excel(xlsx_buffer, index=False, header=False)
            xlsx_buffer.seek(0)

            # 2. Open with openpyxl to apply colors
            wb = openpyxl.load_workbook(xlsx_buffer)
            ws = wb.active
            ws.sheet_view.rightToLeft = True
            yellow_fill = PatternFill(start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid")

            # Pandas drops the 1-based indexing, so row N in pandas is row N in openpyxl
            
            pdf_base_name = Path(pdf_file.filename or "ناشناس").stem
            new_lock_msg = f"تطبیق شده - {pdf_base_name}"
            
            for r_idx in all_matched_rows:
                # Append the locked string to the last column (e.g. column 15) to make it stateful
                msg = old_locks.get(r_idx, new_lock_msg)
                
                ws.cell(row=r_idx, column=15, value=msg) 
                # Color the row yellow
                for col in range(1, 16):
                    ws.cell(row=r_idx, column=col).fill = yellow_fill

            # 3. Save finalized colored workbook back to disk, replacing original if possible
            # Determine path (assuming script is running locally next to files)
            original_filename = excel_file.filename
            if not original_filename:
                original_filename = "bank.xls"
                
            # If the original was .xls, we save as .xlsx
            base_name, _ = os.path.splitext(original_filename)
            new_filename = f"{base_name}.xlsx"
            
            # Since the user runs this locally on Windows, save it to the checkhesab directory
            # Or ideally, same directory where they uploaded from if we knew it.
            # Assuming D:/Checkhesab/ is the working directory based on context.
            save_path = f"d:/Checkhesab/{new_filename}"
            wb.save(save_path)
            
            logger.info(f"Generated locked Excel with {len(matched_rows)} rows highlighted at {save_path}.")
    except Exception as e:
        logger.error(f"Failed to generate locked excel: {e}")

    found_count = sum(1 for r in results if r["status"] == "exact")
    review_count= sum(1 for r in results if r["status"] == "review")
    dupl_count  = sum(1 for r in results if r["status"] == "duplicate")
    
    # We map 'found' technically as exact matches, 'review' as review
    # The frontend uses status to categorize the UI boxes
    return JSONResponse({
        "ok":         True,
        "total":      len(results),
        "found":      found_count,
        "review":     review_count,
        "duplicate":  dupl_count,
        "not_found":  len(results) - found_count - review_count - dupl_count,
        "bank_total": len(bank_txns),
        "pdf_total":  len(pdf_rows),
        "results":    results,
        "debug": {
            "pdf_rows_sample":  pdf_rows[:5],
            "bank_txns_sample": bank_txns[:5],
        }
    })

@app.get("/health")
async def health():
    return {"ok": True}

if __name__ == "__main__":
    print("\n" + "="*60)
    print("  سیستم تطبیق فیش‌های بانکی")
    print("  آدرس: http://localhost:8765")
    print("="*60 + "\n")
    uvicorn.run(app, host="0.0.0.0", port=8765, log_level="info")
