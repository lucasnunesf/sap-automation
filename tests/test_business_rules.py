"""The business rules the README promises, checked one by one."""

import pytest

from build_requisition import sap_value
from intake import validate
from po_handover import group_lines, single_supplier


def line(**overrides):
    """A valid normalised line; each test changes only what it is about."""
    row = {
        "part_number": "MAT-100441",
        "description": "Bearing housing",
        "quantity": 10.0,
        "net_amount": 1000.0,
        "operation_type": "MATERIAL",
        "ipi": 0.05, "icms": 0.18, "pis_cofins": 0.0925, "iss": 0.0,
        "source_file": "supplier_1.xlsx",
        "source_row": 10,
    }
    row.update(overrides)
    return row


# --------------------------------------------------------------------------
# Rule 1: the value sent to SAP depends on the operation type
# --------------------------------------------------------------------------

def test_tooling_takes_net_plus_ipi_only():
    assert sap_value(line(operation_type="MATERIAL")) == 1050.00


def test_service_takes_all_taxes():
    row = line(operation_type="SERVICE", ipi=0.0, icms=0.0, pis_cofins=0.0925, iss=0.05)
    assert sap_value(row) == 1142.50


def test_missing_tax_counts_as_zero():
    assert sap_value(line(operation_type="SERVICE", icms=None, pis_cofins=None)) == 1050.00


# --------------------------------------------------------------------------
# Never guess: bad lines are rejected with a reason
# --------------------------------------------------------------------------

def test_valid_line_is_accepted():
    accepted, rejected = validate([line()])
    assert len(accepted) == 1 and rejected == []


@pytest.mark.parametrize("change, reason", [
    ({"quantity": None}, "quantity is empty"),
    ({"quantity": 0.0}, "quantity must be positive"),
    ({"net_amount": 0.0}, "net amount is zero"),
    ({"net_amount": None}, "net amount could not be read"),
    ({"part_number": None}, "part number is empty"),
    ({"description": None}, "description is empty"),
])
def test_bad_line_is_rejected_with_reason(change, reason):
    accepted, rejected = validate([line(**change)])
    assert accepted == []
    assert any(reason in problem for problem in rejected[0]["problems"])


def test_duplicate_part_number_is_rejected_and_points_to_first_row():
    first = line(source_row=10)
    second = line(source_row=12, part_number="mat-100441")  # case does not matter
    accepted, rejected = validate([first, second])
    assert accepted == [first]
    assert rejected[0]["problems"] == ["duplicate of row 10"]


def test_same_part_number_in_different_files_is_not_a_duplicate():
    accepted, _ = validate([line(), line(source_file="supplier_2.xlsx")])
    assert len(accepted) == 2


# --------------------------------------------------------------------------
# Rules 2 and 3: one supplier per RC, one PO per operation type
# --------------------------------------------------------------------------

def rc_line(code, operation_type, number="1000012345"):
    return {"internal_code": code, "operation_type": operation_type,
            "requisition_number": number}


def test_lines_are_split_into_one_group_per_operation_type():
    lines = [rc_line("A", "MATERIAL"), rc_line("B", "MATERIAL"), rc_line("C", "SERVICE")]
    suppliers = {code: {"supplier_code": "SUP-1"} for code in "ABC"}
    groups, orphans = group_lines(lines, suppliers)
    assert sorted(groups) == ["MATERIAL", "SERVICE"]
    assert len(groups["MATERIAL"]) == 2 and orphans == []
    assert single_supplier(groups) == "SUP-1"


def test_mixed_suppliers_are_refused():
    lines = [rc_line("A", "MATERIAL"), rc_line("B", "MATERIAL")]
    suppliers = {"A": {"supplier_code": "SUP-1"}, "B": {"supplier_code": "SUP-2"}}
    groups, _ = group_lines(lines, suppliers)
    assert single_supplier(groups) is None


def test_line_without_rc_number_is_held_back():
    groups, orphans = group_lines([rc_line("A", "MATERIAL", number=None)],
                                  {"A": {"supplier_code": "SUP-1"}})
    assert groups == {}
    assert orphans[0]["problem"] == "sheet has no requisition number yet"


def test_line_without_known_supplier_is_held_back():
    groups, orphans = group_lines([rc_line("X", "MATERIAL")], {})
    assert groups == {}
    assert orphans[0]["problem"] == "no supplier found for this item"
