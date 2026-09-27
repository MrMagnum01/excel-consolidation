"""Canonical column schema, the messy header spellings the generator uses,
and the normalisation helpers the consolidator uses to undo the mess.
Shared by both sides so the two can be tested against each other exactly.
"""

from __future__ import annotations

import datetime as _dt
import math
import re

# The canonical, normalised column names the master Data sheet uses.
CANONICAL_FIELDS = [
    "order_id",
    "order_date",
    "customer",
    "region",
    "product",
    "quantity",
    "unit_price",
    "amount",
    "sales_rep",
]

# Required for a row to be usable. `region` is not required at the row
# level because a missing region column falls back to the filename.
REQUIRED_FIELDS = ["order_id", "order_date", "product", "quantity", "amount"]

# Four different "house styles" of header row, each with its own spelling
# and its own column order. A real consolidation job sees exactly this:
# every regional office exported from a slightly different template.
HEADER_STYLES = [
    {
        "order_id": "Order ID",
        "order_date": "Order Date",
        "customer": "Customer Name",
        "region": "Region",
        "product": "Product",
        "quantity": "Qty",
        "unit_price": "Unit Price",
        "amount": "Amount",
        "sales_rep": "Sales Rep",
    },
    {
        "order_id": "OrderNo",
        "order_date": "Date",
        "customer": "Client",
        "region": "Territory",
        "product": "Item Description",
        "quantity": "Units",
        "unit_price": "Rate",
        "amount": "Line Total",
        "sales_rep": "Rep",
    },
    {
        "order_id": "Order #",
        "order_date": "Txn Date",
        "customer": "Cust. Name",
        "region": "Branch",
        "product": "Item",
        "quantity": "Quantity",
        "unit_price": "Unit Cost",
        "amount": "Total",
        "sales_rep": "Salesperson",
    },
    {
        "order_id": "Invoice No",
        "order_date": "Order Dt",
        "customer": "Buyer",
        "region": "Area",
        "product": "Product Name",
        "quantity": "Qty Sold",
        "unit_price": "Price",
        "amount": "Sales Amount",
        "sales_rep": "Agent",
    },
]

# Column order each style is laid out in (left to right). Deliberately not
# all the same order, and not always the canonical order.
HEADER_STYLE_ORDERS = [
    ["order_id", "order_date", "customer", "region", "product", "quantity", "unit_price", "amount", "sales_rep"],
    ["order_id", "order_date", "region", "customer", "product", "quantity", "unit_price", "amount", "sales_rep"],
    ["order_date", "order_id", "customer", "product", "region", "quantity", "unit_price", "amount", "sales_rep"],
    ["order_id", "order_date", "customer", "product", "quantity", "unit_price", "amount", "region", "sales_rep"],
]

# Extra substring keywords used as a fallback when a header doesn't exactly
# match one of the HEADER_STYLES labels above (e.g. a stray annotation).
FIELD_KEYWORDS = {
    "order_id": ["order id", "orderno", "order no", "order #", "invoice no", "order number"],
    "order_date": ["order date", "txn date", "order dt", "date"],
    "customer": ["customer", "client", "cust", "buyer"],
    "region": ["region", "territory", "branch", "area"],
    "product": ["product", "item"],
    "quantity": ["qty", "quantity", "units"],
    "unit_price": ["unit price", "unit cost", "rate", "price"],
    "amount": ["amount", "sales amount", "line total", "total"],
    "sales_rep": ["sales rep", "salesperson", "rep", "agent"],
}

# A row is a "totals row" embedded in the data if any of its text cells
# (in any column) matches one of these, case-insensitively.
TOTAL_ROW_MARKERS = {"total", "totals", "grand total", "subtotal", "sub-total", "sum"}

# Date formats the generator renders as *text* and the consolidator must
# parse back. Days are always kept <= 28 so no format is ambiguous.
DATE_FORMATS = [
    "%Y-%m-%d",       # 2026-03-14
    "%m/%d/%Y",       # 03/14/2026
    "%d-%b-%Y",       # 14-Mar-2026
    "%B %d, %Y",      # March 14, 2026
    "%d.%m.%Y",       # 14.03.2026
]

CURRENCY_SYMBOLS = "$€£¥"

# Maps an explicit currency symbol found in a money cell to its ISO code.
# Order matters only for iteration determinism; each symbol is distinct.
CURRENCY_SYMBOL_TO_CODE = {"$": "USD", "€": "EUR", "£": "GBP", "¥": "JPY"}

# Used when a money cell carries no explicit currency symbol at all. This
# is a declared, documented assumption (see README), not a silent guess:
# the consolidator records it in the change log every time it applies.
DEFAULT_CURRENCY = "USD"

_NORMALIZE_RE = re.compile(r"[^a-z0-9]+")


def normalize_header(text: object) -> str:
    """Lowercase, and collapse punctuation/whitespace to single spaces."""
    if text is None:
        return ""
    s = str(text).strip().lower()
    s = _NORMALIZE_RE.sub(" ", s).strip()
    return s


def _build_exact_lookup() -> dict[str, str]:
    lookup: dict[str, str] = {}
    for style in HEADER_STYLES:
        for field, label in style.items():
            lookup[normalize_header(label)] = field
    return lookup


_EXACT_LOOKUP = _build_exact_lookup()


def detect_field(raw_header: object) -> str | None:
    """Map one header cell's text to a canonical field name, or None."""
    norm = normalize_header(raw_header)
    if not norm:
        return None
    if norm in _EXACT_LOOKUP:
        return _EXACT_LOOKUP[norm]
    for field, keywords in FIELD_KEYWORDS.items():
        for kw in keywords:
            if kw in norm:
                return field
    return None


def parse_date(raw: object) -> _dt.date | None:
    """Parse a date rendered as text in one of DATE_FORMATS. Returns None
    if it doesn't match any known format (an invalid/unparsable date)."""
    if raw is None:
        return None
    if isinstance(raw, _dt.datetime):
        return raw.date()
    if isinstance(raw, _dt.date):
        return raw
    text = str(raw).strip()
    if not text:
        return None
    for fmt in DATE_FORMATS:
        try:
            return _dt.datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def parse_money(raw: object) -> float | None:
    """Parse a currency-symbol-and-thousands-separator number. Returns None
    if the cleaned text isn't a valid, finite, non-negative number. Rejects
    NaN/Infinity (Python's float() happily parses the literal text "nan" /
    "infinity", which is not a usable money value), booleans (a bool is a
    subclass of int in Python but is never a legitimate amount), and
    negative values (out of domain for an order amount/unit price in this
    schema)."""
    if raw is None:
        return None
    if isinstance(raw, bool):
        return None
    if isinstance(raw, (int, float)):
        value = float(raw)
        if not math.isfinite(value) or value < 0:
            return None
        return round(value, 2)
    text = str(raw).strip()
    if not text:
        return None
    for sym in CURRENCY_SYMBOLS:
        text = text.replace(sym, "")
    text = text.replace(",", "").strip()
    if not text:
        return None
    try:
        value = float(text)
    except ValueError:
        return None
    if not math.isfinite(value) or value < 0:
        return None
    return round(value, 2)


def detect_currency_symbol(raw: object) -> str | None:
    """Return the explicit currency symbol present in a raw money cell's
    text (e.g. "$1,234.56" -> "$"), or None if the cell carries no symbol
    at all (a bare number). Does not attempt to parse the number."""
    if not isinstance(raw, str):
        return None
    for sym in CURRENCY_SYMBOLS:
        if sym in raw:
            return sym
    return None


def parse_quantity(raw: object) -> int | None:
    """Parse an integer quantity, tolerating '12', '12.0', ' 12 ', '1,200'.
    Returns None (and rejects <= 0) for anything not a positive integer
    quantity."""
    if raw is None:
        return None
    if isinstance(raw, bool):
        return None
    if isinstance(raw, int):
        return raw if raw > 0 else None
    if isinstance(raw, float):
        if raw.is_integer() and raw > 0:
            return int(raw)
        return None
    text = str(raw).strip().replace(",", "")
    if not text:
        return None
    try:
        value = float(text)
    except ValueError:
        return None
    if not value.is_integer() or value <= 0:
        return None
    return int(value)


def is_blank_value(value: object) -> bool:
    return value is None or (isinstance(value, str) and value.strip() == "")
