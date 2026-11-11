"""Unit tests for shared/parsers.py.

Each supplier writes numbers and taxes its own way. These tests pin down that
every variant ends up as the same value, so a change in one parser can't
silently break another supplier.
"""

import pytest

import parsers as P


# --------------------------------------------------------------------------
# parse_number
# --------------------------------------------------------------------------

@pytest.mark.parametrize("value, style, expected", [
    ("1.234,56", "comma", 1234.56),     # Brazilian format
    ("1,234.56", "dot", 1234.56),       # English format
    ("BRL 1.145,00", "comma", 1145.00), # currency in the cell
    (" 48,90 ", "comma", 48.90),        # padding
    (42, "dot", 42.0),                  # already a number
    (0.05, "dot", 0.05),
    ("-15,50", "comma", -15.50),        # negative stays negative
])
def test_parse_number_reads_every_format(value, style, expected):
    assert P.parse_number(value, style) == pytest.approx(expected)


@pytest.mark.parametrize("value", [None, "", "  ", "-", "N/A", "n/a"])
def test_parse_number_returns_none_for_empty_cells(value):
    assert P.parse_number(value) is None


# --------------------------------------------------------------------------
# parse_rate: the same 5% IPI, as each sample supplier sends it
# --------------------------------------------------------------------------

@pytest.mark.parametrize("value, tax_style, style, net, supplier", [
    ("5%", "percent_text", "comma", None, "supplier 1"),
    (5.0, "percent_number", "dot", None, "supplier 2"),
    (0.05, "decimal_fraction", "dot", None, "supplier 3"),
    (1267.50, "amount", "dot", 25350.00, "supplier 4"),
])
def test_parse_rate_turns_every_format_into_a_fraction(value, tax_style, style, net, supplier):
    assert P.parse_rate(value, tax_style, style, net) == pytest.approx(0.05), supplier


def test_parse_rate_with_comma_decimal():
    assert P.parse_rate("9,25%", "percent_text", "comma") == pytest.approx(0.0925)


def test_parse_rate_empty_cell_means_no_tax():
    assert P.parse_rate(None, "percent_number") == 0.0


def test_parse_rate_amount_without_net_value_does_not_divide_by_zero():
    assert P.parse_rate(100, "amount", net_amount=0) == 0.0
    assert P.parse_rate(100, "amount", net_amount=None) == 0.0


def test_parse_rate_unknown_style_returns_zero():
    assert P.parse_rate("5%", "something_new") == 0.0


# --------------------------------------------------------------------------
# parse_combined_taxes: supplier 5 puts every tax in one cell
# --------------------------------------------------------------------------

def test_combined_taxes_are_split_per_tax():
    rates = P.parse_combined_taxes("IPI 5%; ICMS 18%; PIS/COFINS 9.25%")
    assert rates == pytest.approx(
        {"ipi": 0.05, "icms": 0.18, "pis_cofins": 0.0925, "iss": 0.0})


def test_combined_taxes_with_comma_decimal_and_iss():
    rates = P.parse_combined_taxes("PIS/COFINS 9,25%; ISS 5%", "comma")
    assert rates["pis_cofins"] == pytest.approx(0.0925)
    assert rates["iss"] == pytest.approx(0.05)
    assert rates["ipi"] == 0.0


def test_combined_taxes_reject_an_unknown_tax():
    # a tax the system doesn't know would leave the SAP total wrong:
    # the buyer has to see it before it becomes an order
    with pytest.raises(ValueError, match="unknown tax 'VAT'"):
        P.parse_combined_taxes("IPI 5%; VAT 20%")


@pytest.mark.parametrize("value", [None, "", "no taxes"])
def test_combined_taxes_empty_means_all_zero(value):
    assert set(P.parse_combined_taxes(value).values()) == {0.0}


# --------------------------------------------------------------------------
# Packed cells and descriptions
# --------------------------------------------------------------------------

@pytest.mark.parametrize("value, expected", [
    ("36 pcs", 36.0),
    ("12 h", 12.0),
    (8, 8.0),
    (None, None),
])
def test_quantity_with_unit(value, expected):
    assert P.parse_quantity_with_unit(value) == expected


@pytest.mark.parametrize("value, expected", [
    ("[MAT] Conveyor roller", ("MATERIAL", "Conveyor roller")),
    ("[SRV] Installation", ("SERVICE", "Installation")),
    ("[serv]   Commissioning", ("SERVICE", "Commissioning")),
    ("Plain description", (None, "Plain description")),
    (None, (None, None)),
])
def test_strip_type_prefix(value, expected):
    assert P.strip_type_prefix(value) == expected


@pytest.mark.parametrize("value, expected", [
    ("Material", "MATERIAL"),
    ("MAT", "MATERIAL"),
    ("Service", "SERVICE"),
    ("srv", "SERVICE"),
    ("Other", None),
    (None, None),
])
def test_normalise_operation_type(value, expected):
    assert P.normalise_operation_type(value) == expected


def test_clean_text_collapses_spaces():
    assert P.clean_text("  Flange   coupling\n DN80 ") == "Flange coupling DN80"
    assert P.clean_text(None) is None
