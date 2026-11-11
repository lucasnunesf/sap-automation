"""A value that is ambiguous or not in the supplier's configured format must
become a rejected line with a reason, never a wrong number."""

import pytest

import parsers as P
from intake import build_row, validate


# --------------------------------------------------------------------------
# parse_number: the text has to match the configured format
# --------------------------------------------------------------------------

@pytest.mark.parametrize("value, style", [
    ("1.234,56", "dot"),     # Brazilian number in an English-format file
    ("1234.56", "comma"),    # English number in a Brazilian-format file
    ("1e3", "dot"),          # letters in the middle are not a currency prefix
    ("1.23.4", "comma"),     # thousands separator in the wrong place
    ("12 abc 3", "dot"),
    ("abc", "dot"),          # text where a number should be
])
def test_number_outside_the_configured_format_is_an_error(value, style):
    with pytest.raises(ValueError, match="format"):
        P.parse_number(value, style)


@pytest.mark.parametrize("value, style, expected", [
    ("1.234", "comma", 1234.0),   # one thousands separator, as configured
    ("1.500", "comma", 1500.0),
    ("1,234", "dot", 1234.0),
    ("R$ 1.145,00", "comma", 1145.0),
])
def test_number_in_the_configured_format_is_read(value, style, expected):
    assert P.parse_number(value, style) == pytest.approx(expected)


# --------------------------------------------------------------------------
# parse_quantity_with_unit: uses the supplier's decimal style
# --------------------------------------------------------------------------

@pytest.mark.parametrize("value, expected", [
    ("1,5 kg", 1.5),
    ("1.000 pcs", 1000.0),
    ("36 pcs", 36.0),
])
def test_quantity_with_unit_follows_the_decimal_style(value, expected):
    assert P.parse_quantity_with_unit(value, "comma") == pytest.approx(expected)


@pytest.mark.parametrize("value", ["approx 36 pcs", "pcs"])
def test_quantity_with_unit_rejects_anything_else(value):
    with pytest.raises(ValueError, match="format"):
        P.parse_quantity_with_unit(value, "comma")


# --------------------------------------------------------------------------
# parse_combined_taxes: any separator, and no tax is dropped
# --------------------------------------------------------------------------

def test_combined_taxes_separated_by_comma():
    rates = P.parse_combined_taxes("IPI 5%, ICMS 18%")
    assert rates["ipi"] == pytest.approx(0.05)
    assert rates["icms"] == pytest.approx(0.18)


# --------------------------------------------------------------------------
# End to end: the bad cell becomes a rejected line with the reason
# --------------------------------------------------------------------------

def test_price_in_the_wrong_format_rejects_the_line():
    config = {"name": "Supplier X", "tax_style": "percent_text"}
    columns = {"part_number": "Part", "description": "Desc", "quantity": "Qty",
               "unit_price": "Price", "ipi": "IPI"}
    raw = {"Part": "MAT-1", "Desc": "Roller", "Qty": 10, "Price": "1234.56", "IPI": "5%"}
    row = build_row(raw, config, "comma", columns, {}, "supplier_x.xlsx", 7, "Offer")

    accepted, rejected = validate([row])

    assert accepted == []
    assert any("unit_price" in p and "format" in p for p in rejected[0]["problems"])


def test_supplier_5_services_carry_the_right_value_to_sap():
    # The sample file used to say "PIS/COFINS 9.25%" in a comma-format file.
    # The old parser read 925% and sent 162,740.00 to SAP instead of 18,051.50.
    import yaml
    from pathlib import Path
    from build_requisition import sap_value
    from intake import read_supplier

    root = Path(__file__).resolve().parents[1]
    config = yaml.safe_load(open(root / "shared" / "suppliers.yaml", encoding="utf-8"))
    config = config.get("suppliers", config)["supplier_5"]
    rows = read_supplier(root / "01-supplier-intake" / "sample_data" / "supplier_5.xlsx", config)

    assert all(not row["parse_errors"] for row in rows)
    services = [row for row in rows if row["operation_type"] == "SERVICE"]
    assert [sap_value(row) for row in services] == [18051.50, 14167.00, 14395.50]
