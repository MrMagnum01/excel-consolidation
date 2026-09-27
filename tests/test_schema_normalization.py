"""Unit tests on the normalisation helpers in isolation, no I/O."""

import pytest

from excel_consolidation.schema import (
    HEADER_STYLES,
    detect_field,
    parse_date,
    parse_money,
    parse_quantity,
)


@pytest.mark.parametrize("field,label", [
    (field, label)
    for style in HEADER_STYLES
    for field, label in style.items()
])
def test_detect_field_recognises_every_shipped_header_variant(field, label):
    assert detect_field(label) == field


def test_detect_field_returns_none_for_unrelated_header():
    assert detect_field("SKU") is None
    assert detect_field("Warehouse") is None
    assert detect_field("") is None
    assert detect_field(None) is None


@pytest.mark.parametrize("text,expected", [
    ("2026-03-14", (2026, 3, 14)),
    ("03/14/2026", (2026, 3, 14)),
    ("14-Mar-2026", (2026, 3, 14)),
    ("March 14, 2026", (2026, 3, 14)),
    ("14.03.2026", (2026, 3, 14)),
])
def test_parse_date_accepts_every_shipped_format(text, expected):
    d = parse_date(text)
    assert d is not None
    assert (d.year, d.month, d.day) == expected


@pytest.mark.parametrize("text", ["31/02/2026", "not a date", "00/00/0000", "", None])
def test_parse_date_rejects_invalid_text(text):
    assert parse_date(text) is None


@pytest.mark.parametrize("text,expected", [
    ("$1,234.56", 1234.56),
    ("€999.00", 999.00),
    ("£500", 500.0),
    ("1234.56", 1234.56),
    (" 500 ", 500.0),
    ("1,200.00", 1200.00),
])
def test_parse_money_strips_symbols_and_separators(text, expected):
    assert parse_money(text) == expected


@pytest.mark.parametrize("text", ["TBD", "call for price", "???", "", None])
def test_parse_money_rejects_non_numeric_text(text):
    assert parse_money(text) is None


@pytest.mark.parametrize("text,expected", [
    ("12", 12),
    ("12.0", 12),
    (" 12 ", 12),
    ("1,200", 1200),
    (12, 12),
])
def test_parse_quantity_accepts_valid_forms(text, expected):
    assert parse_quantity(text) == expected


@pytest.mark.parametrize("text", ["many", "n/a", "-", "0", "-5", "", None, "12.5"])
def test_parse_quantity_rejects_invalid_or_non_positive(text):
    assert parse_quantity(text) is None
