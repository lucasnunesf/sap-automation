"""Rules added to validate(): the amounts have to agree with each other, and a
tax the supplier left blank is not the same as a tax that doesn't apply."""

import pytest

import parsers as P
from intake import validate
from test_business_rules import line


# --------------------------------------------------------------------------
# unit price x quantity has to match the net amount
# --------------------------------------------------------------------------

def test_net_amount_that_does_not_match_price_times_quantity_is_rejected():
    accepted, rejected = validate([line(unit_price=100.0, quantity=10.0, net_amount=1500.0)])
    assert accepted == []
    assert any("does not match unit price x quantity" in p for p in rejected[0]["problems"])


@pytest.mark.parametrize("unit_price, quantity, net_amount", [
    (100.0, 10.0, 1000.0),    # exact
    (3.33, 3.0, 10.00),       # 9.99: the unit price was rounded to the cent
    (0.33, 1500.0, 500.00),   # 495.00: half a cent per unit, 1,500 units
])
def test_rounding_of_the_unit_price_is_tolerated(unit_price, quantity, net_amount):
    accepted, _ = validate([line(unit_price=unit_price, quantity=quantity, net_amount=net_amount)])
    assert len(accepted) == 1


# --------------------------------------------------------------------------
# blank tax cell (missing) vs "-", 0 or 0% (does not apply)
# --------------------------------------------------------------------------

@pytest.mark.parametrize("value, style", [(None, "percent_number"), ("", "percent_text"),
                                          ("  ", "percent_text"), (None, "amount")])
def test_blank_tax_cell_is_missing(value, style):
    assert P.parse_rate(value, style, net_amount=1000.0) is None


@pytest.mark.parametrize("value, style", [("-", "percent_text"), ("N/A", "percent_text"),
                                          (0, "percent_number"), ("0%", "percent_text"),
                                          (0, "amount")])
def test_explicit_dash_or_zero_means_the_tax_does_not_apply(value, style):
    assert P.parse_rate(value, style, net_amount=1000.0) == 0.0


@pytest.mark.parametrize("value", [None, ""])
def test_blank_combined_tax_cell_is_missing(value):
    assert set(P.parse_combined_taxes(value).values()) == {None}


@pytest.mark.parametrize("operation_type, tax", [
    ("MATERIAL", "ipi"),                                       # material uses IPI
    ("SERVICE", "ipi"), ("SERVICE", "icms"), ("SERVICE", "pis_cofins"), ("SERVICE", "iss"),
])
def test_missing_tax_used_in_the_sap_value_rejects_the_line(operation_type, tax):
    accepted, rejected = validate([line(operation_type=operation_type, **{tax: None})])
    assert accepted == []
    assert any(f"{tax} is missing" in p for p in rejected[0]["problems"])


def test_missing_tax_not_used_in_the_sap_value_is_accepted():
    # material only takes net + IPI, so a blank ISS changes nothing
    accepted, _ = validate([line(operation_type="MATERIAL", iss=None)])
    assert len(accepted) == 1
