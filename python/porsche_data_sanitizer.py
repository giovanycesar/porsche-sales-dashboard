# porsche_data_sanitizer.py
"""
Porsche Sales Data Sanitizer
=============================
Reads ``porsche_database_raw.xlsx``, applies every sanitization rule from
``porsche_sales_sanitization_schema.md``, and writes the result to
``porsche_database_cleaned.xlsx`` (the raw file is never modified).
"""

import pandas as pd
import re
import math
from datetime import datetime
from pathlib import Path
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from word2number import w2n

# ============================================================================
# Reference data
# ============================================================================

KNOWN_MODELS = [
    "911 Carrera", "911 Carrera S", "911 Carrera GTS",
    "911 Turbo", "911 Turbo S",
    "911 GT3", "911 GT3 RS", "911 Dakar",
    "911 Targa 4", "911 Targa 4S",
    "718 Cayman", "718 Cayman S", "718 Cayman GT4 RS",
    "718 Boxster", "718 Boxster GTS", "718 Spyder RS",
    "Cayenne", "Cayenne S", "Cayenne Coupe",
    "Cayenne E-Hybrid", "Cayenne Turbo", "Cayenne Turbo GT",
    "Macan", "Macan S", "Macan T", "Macan GTS", "Macan Electric",
    "Panamera", "Panamera 4", "Panamera 4S",
    "Panamera Turbo", "Panamera Turbo S", "Panamera 4 E-Hybrid",
    "Taycan", "Taycan 4S", "Taycan GTS",
    "Taycan Turbo", "Taycan Turbo S", "Taycan Cross Turismo",
]

# Build a lookup:  lower-cased canonical -> canonical
_MODEL_LOOKUP: dict[str, str] = {m.lower(): m for m in KNOWN_MODELS}

PAYMENT_METHODS = {
    "credit card": "Credit Card",
    "creditcard": "Credit Card",
    "credit": "Credit Card",
    "credit card payment": "Credit Card",
    "debit card": "Debit Card",
    "debitcard": "Debit Card",
    "bank transfer": "Bank Transfer",
    "bank_transfer": "Bank Transfer",
    "bank-transfer": "Bank Transfer",
    "banktransfer": "Bank Transfer",
    "wire transfer": "Wire Transfer",
    "wiretransfer": "Wire Transfer",
    "wire-transfer": "Wire Transfer",
    "wire": "Wire Transfer",
    "bank wire": "Wire Transfer",
    "financing": "Financing",
    "financing plan": "Financing",
    "finance": "Financing",
    "lease": "Lease",
    "leasing": "Lease",
    "lease plan": "Lease",
    "cash": "Cash",
    "cash payment": "Cash",
    "ach payment": "ACH Payment",
    "ach": "ACH Payment",
    "crypto payment": "Crypto Payment",
    "crypto": "Crypto Payment",
}

US_STATES = {
    "alabama": "AL", "alaska": "AK", "arizona": "AZ", "arkansas": "AR",
    "california": "CA", "colorado": "CO", "connecticut": "CT",
    "delaware": "DE", "florida": "FL", "georgia": "GA", "hawaii": "HI",
    "idaho": "ID", "illinois": "IL", "indiana": "IN", "iowa": "IA",
    "kansas": "KS", "kentucky": "KY", "louisiana": "LA", "maine": "ME",
    "maryland": "MD", "massachusetts": "MA", "michigan": "MI",
    "minnesota": "MN", "mississippi": "MS", "missouri": "MO",
    "montana": "MT", "nebraska": "NE", "nevada": "NV",
    "new hampshire": "NH", "new jersey": "NJ", "new mexico": "NM",
    "new york": "NY", "north carolina": "NC", "north dakota": "ND",
    "ohio": "OH", "oklahoma": "OK", "oregon": "OR", "pennsylvania": "PA",
    "rhode island": "RI", "south carolina": "SC", "south dakota": "SD",
    "tennessee": "TN", "texas": "TX", "utah": "UT", "vermont": "VT",
    "virginia": "VA", "washington": "WA", "west virginia": "WV",
    "wisconsin": "WI", "wyoming": "WY",
}

# Reverse lookup:  two-letter code -> itself  (for quick validation)
_VALID_CODES = set(US_STATES.values())

DELIVERY_STATUSES = {
    "delivered": "Delivered",
    "deliverd": "Delivered",          # common typo
    "pending": "Pending",
    "in transit": "In Transit",
    "intransit": "In Transit",
    "cancelled": "Cancelled",
    "canceled": "Cancelled",
    "awaiting delivery": "Awaiting Delivery",
    "awaiting pickup": "Awaiting Pickup",
    "pending approval": "Pending Approval",
    "pending review": "Pending Review",
    "shipped": "Shipped",
    "awaiting review": "Awaiting Review",
}

# ============================================================================
# Sanitization helpers
# ============================================================================

def _to_str(value) -> str:
    """Convert any value to a stripped string; handle datetime objects."""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, pd.Timestamp):
        return value.strftime("%Y-%m-%d")
    if pd.isna(value):
        return ""
    return str(value).strip()


# ---- DATES -----------------------------------------------------------------

def sanitize_date(raw) -> str:
    """Normalise to ISO YYYY-MM-DD or return INVALID."""

    # If pandas already parsed it as a Timestamp
    if isinstance(raw, (datetime, pd.Timestamp)):
        return raw.strftime("%Y-%m-%d")

    s = _to_str(raw)
    if not s:
        return "INVALID"

    # Remove ordinal suffixes: 1st, 2nd, 3rd, 4th ...
    s_clean = re.sub(r'(\d+)(st|nd|rd|th)\b', r'\1', s, flags=re.IGNORECASE)

    # Remove time component if present  (e.g. "2024-04-18 00:00:00")
    s_clean = re.sub(r'\s+\d{2}:\d{2}:\d{2}$', '', s_clean).strip()

    # Explicit format patterns (strict parsing)
    patterns = [
        ("%Y-%m-%d", s_clean),
        ("%Y/%m/%d", s_clean),
        ("%Y.%m.%d", s_clean),
        ("%m/%d/%Y", s_clean),
        ("%m/%d/%y", s_clean),
        ("%m-%d-%y", s_clean),
        ("%B %d, %Y", s_clean),
        ("%b %d, %Y", s_clean),
        ("%B %d %Y", s_clean),
        ("%b %d %Y", s_clean),
    ]
    for fmt, text in patterns:
        try:
            dt = datetime.strptime(text, fmt)
            # Extra calendar validation (strptime does NOT reject Feb-30, etc.)
            # Re-format and re-parse to confirm
            iso = dt.strftime("%Y-%m-%d")
            _ = datetime.strptime(iso, "%Y-%m-%d")
            return iso
        except ValueError:
            continue

    return "INVALID"


# ---- PORSCHE MODEL ---------------------------------------------------------

def sanitize_model(raw) -> str:
    s = _to_str(raw)
    if not s:
        return "INVALID"
    key = s.lower().strip()

    # Exact match
    if key in _MODEL_LOOKUP:
        return _MODEL_LOOKUP[key]

    # Longest-match search (try longer canonical names first)
    for canon in sorted(KNOWN_MODELS, key=len, reverse=True):
        if canon.lower() in key:
            return canon

    # Fallback: title-case
    return s.title()


# ---- MODEL YEAR ------------------------------------------------------------

_WORD_YEAR_MAP = {
    "twenty twenty": 2020,
    "twenty twenty one": 2021,
    "twenty twenty two": 2022,
    "twenty twenty three": 2023,
    "twenty twenty four": 2024,
    "twenty twenty five": 2025,
    "twenty twenty six": 2026,
    "twenty twenty seven": 2027,
    "two thousand twenty": 2020,
    "two thousand twenty one": 2021,
    "two thousand twenty two": 2022,
    "two thousand twenty three": 2023,
    "two thousand twenty four": 2024,
    "two thousand twenty five": 2025,
    "two thousand twenty six": 2026,
    "two thousand twenty seven": 2027,
}


def sanitize_year(raw) -> str:
    # Numeric already
    if isinstance(raw, (int, float)) and not math.isnan(raw):
        y = int(raw)
        return str(y) if 1990 <= y <= 2035 else "INVALID"

    s = _to_str(raw).lower().strip()
    if not s:
        return "INVALID"

    # Try word-year map (longest first)
    for phrase, year in sorted(_WORD_YEAR_MAP.items(), key=lambda x: -len(x[0])):
        if phrase in s:
            return str(year)

    # Try word2number
    try:
        y = w2n.word_to_num(s)
        if 1990 <= y <= 2035:
            return str(y)
    except Exception:
        pass

    # "20-24" / "20 24" -> "2024"
    m = re.match(r'^(\d{2})[\s\-](\d{2})$', s)
    if m:
        y = int(m.group(1) + m.group(2))
        return str(y) if 1990 <= y <= 2035 else "INVALID"

    # Plain 4-digit number
    m = re.search(r'(\d{4})', s)
    if m:
        y = int(m.group(1))
        return str(y) if 1990 <= y <= 2035 else "INVALID"

    # Plain 2-digit number
    m = re.search(r'(\d{2})', s)
    if m:
        y = 2000 + int(m.group(1))
        return str(y) if 1990 <= y <= 2035 else "INVALID"

    return "INVALID"


# ---- SALE PRICE ------------------------------------------------------------

def sanitize_price(raw) -> str:
    # Already numeric
    if isinstance(raw, (int, float)) and not math.isnan(raw):
        return f"{float(raw):.2f}"

    s = _to_str(raw)
    if not s:
        return "INVALID"

    original = s
    s = s.lower().strip()
    # Remove currency labels
    s = s.replace("usd", "").replace("dollars", "").replace("dollar", "").replace("$", "").strip()

    # Detect "k" multiplier  (e.g. "121k", "188k")
    m_k = re.match(r'^([\d.,]+)\s*k$', s)
    if m_k:
        num_str = m_k.group(1).replace(",", "")
        try:
            return f"{float(num_str) * 1000:.2f}"
        except ValueError:
            pass

    # European-style decimal:  "112.750" (no cents) or "89.750,00"
    # Heuristic: if there is exactly one dot and the fractional part is 3 digits,
    # it's a thousands separator (European), not a decimal point.
    # Also handle "89.750,00" -> swap comma/dot
    if "," in s and "." in s:
        # Could be "89.750,00" (European) or "$79,500.00" (US)
        # If comma comes AFTER the last dot -> European
        last_dot = s.rfind(".")
        last_comma = s.rfind(",")
        if last_comma > last_dot:
            # European: dots are thousands, comma is decimal  "89.750,00"
            s = s.replace(".", "").replace(",", ".")
        else:
            # US: commas are thousands, dot is decimal  "$79,500.00"
            s = s.replace(",", "")
    elif "." in s and "," not in s:
        # Check for European thousands separator:  "112.750" (3 digits after dot)
        m_eu = re.match(r'^([\d.]+)$', s.strip())
        if m_eu:
            parts = s.strip().split(".")
            # If every group after the first is exactly 3 digits -> thousands sep
            if len(parts) > 1 and all(len(p) == 3 for p in parts[1:]):
                s = s.replace(".", "")
    elif "," in s and "." not in s:
        # "104,500" -> remove comma (US thousands separator)
        # But "68,900" is also just US thousands
        s = s.replace(",", "")

    s = s.replace(" ", "").strip()

    try:
        val = float(s)
        return f"{val:.2f}"
    except ValueError:
        pass

    # Try word-to-number
    try:
        val = w2n.word_to_num(original)
        return f"{float(val):.2f}"
    except Exception:
        pass

    return "INVALID"


# ---- VEHICLE MILEAGE -------------------------------------------------------

def sanitize_mileage(raw) -> str:
    # Already numeric
    if isinstance(raw, (int, float)) and not math.isnan(raw):
        return str(int(raw))

    s = _to_str(raw)
    if not s:
        return "INVALID"

    low = s.lower().strip()

    # Zero / new
    if low in ("new", "new car", "0", "0 mi", "0 miles"):
        return "0"
    if "zero" in low or "new" in low:
        return "0"

    is_km = bool(re.search(r'\bkm\b', low, re.IGNORECASE))

    # Strip labels
    cleaned = re.sub(r'(miles|mile|mi\.?|km)\b', '', low, flags=re.IGNORECASE)
    cleaned = re.sub(r'^[\s:]+|[\s:]+$', '', cleaned)   # leading "Miles: "
    cleaned = cleaned.strip(": ")

    # Extract numeric part (handles "1.200", "9,800", "41,000", etc.)
    num_match = re.search(r'([\d.,]+)', cleaned)
    if num_match:
        num_str = num_match.group(1)

        # Decide whether dot is decimal or thousands separator
        if "." in num_str and "," not in num_str:
            parts = num_str.split(".")
            # If fraction part is exactly 3 digits -> thousands sep (e.g. "1.200")
            if len(parts) == 2 and len(parts[1]) == 3:
                num_str = num_str.replace(".", "")
            # Otherwise keep as-is (real decimal)
        elif "," in num_str and "." not in num_str:
            num_str = num_str.replace(",", "")
        elif "," in num_str and "." in num_str:
            # Mixed: treat comma as thousands  "41,000.5"
            num_str = num_str.replace(",", "")

        try:
            val = float(num_str)
            if is_km:
                val = val * 0.621371
            return str(int(round(val)))
        except ValueError:
            pass

    # Word-to-number  ("twelve thousand miles")
    try:
        val = w2n.word_to_num(low.replace("miles", "").replace("mi", "").strip())
        return str(int(val))
    except Exception:
        pass

    return "INVALID"


# ---- PAYMENT METHOD --------------------------------------------------------

def sanitize_payment(raw) -> str:
    s = _to_str(raw)
    if not s:
        return "INVALID"
    key = s.lower().replace("-", " ").replace("_", " ").strip()
    if key in PAYMENT_METHODS:
        return PAYMENT_METHODS[key]
    return s.title()


# ---- CITY ------------------------------------------------------------------

def sanitize_city(raw) -> str:
    s = _to_str(raw)
    if not s:
        return "INVALID"
    # Title-case but keep "St." etc.
    return s.title()


# ---- STATE -----------------------------------------------------------------

def sanitize_state(raw) -> str:
    s = _to_str(raw)
    if not s:
        return "INVALID"
    low = s.lower().strip()

    # Already a valid 2-letter code?
    if len(low) == 2 and low.upper() in _VALID_CODES:
        return low.upper()

    # Full name lookup
    if low in US_STATES:
        return US_STATES[low]

    return "INVALID"


# ---- DELIVERY STATUS -------------------------------------------------------

def sanitize_delivery(raw) -> str:
    s = _to_str(raw)
    if not s:
        return "INVALID"
    # Strip punctuation noise
    clean = re.sub(r'[!\.\*]+', '', s).lower().replace("-", " ").strip()

    if clean in DELIVERY_STATUSES:
        return DELIVERY_STATUSES[clean]

    # Compact comparison (remove spaces)
    compact = clean.replace(" ", "")
    for key, canon in DELIVERY_STATUSES.items():
        if compact == key.replace(" ", ""):
            return canon

    return s.title()


# ============================================================================
# Main pipeline
# ============================================================================

def main():
    root = Path(__file__).parent
    raw_path = root / "porsche_database_raw.xlsx"

    if not raw_path.is_file():
        print(f"[ERROR] Raw file not found: {raw_path}")
        return

    df = pd.read_excel(raw_path)
    print(f"Loaded {len(df)} rows from {raw_path.name}")

    # --- Insert sanitized columns right after each source column -------------
    def insert_after(dataframe, source_col, new_col, values):
        idx = list(dataframe.columns).index(source_col) + 1
        dataframe.insert(idx, new_col, values)

    insert_after(df, "sale_date",        "SaleDateSanitized",       df["sale_date"].apply(sanitize_date))
    insert_after(df, "porsche_model",    "PorscheModelSanitized",   df["porsche_model"].apply(sanitize_model))
    insert_after(df, "model_year",       "ModelYearSanitized",      df["model_year"].apply(sanitize_year))
    insert_after(df, "sale_price",       "SalesPriceSanitized",     df["sale_price"].apply(sanitize_price))
    insert_after(df, "vehicle_mileage",  "VehicleMileageSanitized", df["vehicle_mileage"].apply(sanitize_mileage))
    insert_after(df, "payment_method",   "PayMethodSanitized",      df["payment_method"].apply(sanitize_payment))
    insert_after(df, "city",             "CitySanitized",           df["city"].apply(sanitize_city))
    insert_after(df, "state",            "StateSanitized",          df["state"].apply(sanitize_state))
    insert_after(df, "delivery_status",  "DeliveryStatusSanitized", df["delivery_status"].apply(sanitize_delivery))

    # --- Quality checks ------------------------------------------------------
    sanitized_cols = [c for c in df.columns if c.endswith("Sanitized")]
    blanks = (df[sanitized_cols] == "").sum()
    if blanks.any():
        print("[WARN] Blank sanitized values found:")
        print(blanks[blanks > 0])

    invalids = (df[sanitized_cols] == "INVALID").sum()
    print(f"\nINVALID counts per sanitized column:\n{invalids.to_string()}")

    # --- Write output with color formatting -----------------------------------
    out_path = root / "porsche_database_cleaned.xlsx"
    df.to_excel(out_path, index=False, engine="openpyxl")

    # Re-open with openpyxl to apply styles
    from openpyxl import load_workbook

    wb = load_workbook(out_path)
    ws = wb.active

    # Define styles
    # Raw columns: blue tones
    raw_header_fill   = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")  # dark blue
    raw_header_font   = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
    raw_cell_fill     = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")  # light blue

    # Sanitized columns: green tones
    san_header_fill   = PatternFill(start_color="375623", end_color="375623", fill_type="solid")  # dark green
    san_header_font   = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
    san_cell_fill     = PatternFill(start_color="D9E9D2", end_color="D9E9D2", fill_type="solid")  # light green

    # INVALID cell style
    invalid_font      = Font(name="Calibri", bold=True, color="CC0000", size=11)
    invalid_fill      = PatternFill(start_color="FDE0DC", end_color="FDE0DC", fill_type="solid")  # light red

    # Normal cell font
    normal_font       = Font(name="Calibri", size=11)

    # Thin border
    thin_border = Border(
        left=Side(style="thin", color="BFBFBF"),
        right=Side(style="thin", color="BFBFBF"),
        top=Side(style="thin", color="BFBFBF"),
        bottom=Side(style="thin", color="BFBFBF"),
    )

    # Identify which columns are sanitized
    col_names = [cell.value for cell in ws[1]]
    sanitized_indices = set()
    for idx, name in enumerate(col_names, start=1):
        if name and name.endswith("Sanitized"):
            sanitized_indices.add(idx)

    # Apply header row styles
    for col_idx, cell in enumerate(ws[1], start=1):
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        if col_idx in sanitized_indices:
            cell.fill = san_header_fill
            cell.font = san_header_font
        else:
            cell.fill = raw_header_fill
            cell.font = raw_header_font

    # Apply data cell styles
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.border = thin_border
            cell.alignment = Alignment(vertical="center")
            col_idx = cell.column
            is_sanitized_col = col_idx in sanitized_indices

            if is_sanitized_col:
                if str(cell.value).strip() == "INVALID":
                    cell.font = invalid_font
                    cell.fill = invalid_fill
                else:
                    cell.font = normal_font
                    cell.fill = san_cell_fill
            else:
                cell.font = normal_font
                cell.fill = raw_cell_fill

    # Auto-adjust column widths
    for col_cells in ws.columns:
        max_len = 0
        col_letter = col_cells[0].column_letter
        for cell in col_cells:
            val = str(cell.value) if cell.value else ""
            max_len = max(max_len, len(val))
        ws.column_dimensions[col_letter].width = min(max_len + 4, 40)

    # Freeze header row
    ws.freeze_panes = "A2"

    wb.save(out_path)
    print(f"\n[OK] Cleaned file saved to: {out_path.name}")
    print(f"     Rows: {len(df)}  |  Columns: {len(df.columns)}")
    print(f"     Raw columns   -> Blue headers + light-blue cells")
    print(f"     Sanitized cols -> Green headers + light-green cells")
    print(f"     INVALID values -> Red font + light-red background")
    print(f"     Original columns preserved: {all(c in df.columns for c in ['sale_id','sale_date','customer_name','porsche_model','model_year','sale_price','vehicle_mileage','payment_method','city','state','salesperson','delivery_status'])}")


if __name__ == "__main__":
    main()

