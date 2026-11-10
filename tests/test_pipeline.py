"""End-to-end runs of the Python stages on the sample data.

These run the scripts exactly as a user would, in a temporary folder, so the
results in the README are checked on every change.
"""

import subprocess
import sys
from pathlib import Path

import openpyxl
import pytest

ROOT = Path(__file__).resolve().parents[1]
SAMPLES = ROOT / "01-supplier-intake" / "sample_data"
INTAKE = ROOT / "01-supplier-intake" / "src" / "intake.py"
BUILDER = ROOT / "02-requisition-builder" / "src" / "build_requisition.py"
HANDOVER = ROOT / "04-po-handover" / "src" / "po_handover.py"


def run(script, *args, cwd):
    return subprocess.run([sys.executable, str(script), *map(str, args)],
                          cwd=cwd, capture_output=True, text=True)


def fake_sap(requisition):
    """Stand in for stage 03: write the RC numbers SAP would return."""
    workbook = openpyxl.load_workbook(requisition)
    workbook["Tooling"]["K2"] = "1000012345"
    workbook["Service"]["K2"] = "1000012346"
    workbook.save(requisition)


def run_until_handover(tmp_path, *intake_args):
    normalised = tmp_path / "normalised_items.xlsx"
    requisition = tmp_path / "Input_RC_automatic_filled.xlsx"
    step1 = run(INTAKE, *intake_args, "--output", normalised, cwd=tmp_path)
    step2 = run(BUILDER, "--input", normalised, "--output", requisition, cwd=tmp_path)
    assert step1.returncode == 0, step1.stderr
    assert step2.returncode == 0, step2.stderr
    fake_sap(requisition)
    step4 = run(HANDOVER, "--requisition", requisition, "--normalised", normalised,
                "--output-dir", tmp_path, cwd=tmp_path)
    return step1, step2, step4


def test_single_supplier_runs_end_to_end(tmp_path):
    step1, step2, step4 = run_until_handover(
        tmp_path, "--input", SAMPLES / "supplier_2.xlsx")

    assert "10 normalised" in step1.stdout
    assert "10 written (7 tooling, 3 service)" in step2.stdout
    assert step4.returncode == 0, step4.stdout + step4.stderr
    assert "2 purchase orders generated" in step4.stdout
    assert sorted(p.name for p in tmp_path.glob("PO_*.xlsx")) == [
        "PO_SUP-10045_MATERIAL.xlsx", "PO_SUP-10045_SERVICE.xlsx"]


@pytest.mark.parametrize("supplier", [1, 2, 3, 4, 5])
def test_every_sample_supplier_reaches_purchase_orders(tmp_path, supplier):
    _, _, step4 = run_until_handover(
        tmp_path, "--input", SAMPLES / f"supplier_{supplier}.xlsx")
    assert step4.returncode == 0, step4.stdout + step4.stderr
    assert list(tmp_path.glob("PO_*.xlsx")), step4.stdout


def test_all_suppliers_report_the_planted_problems(tmp_path):
    result = run(INTAKE, "--input", SAMPLES, "--all",
                 "--output", tmp_path / "out.xlsx", cwd=tmp_path)
    assert result.returncode == 0, result.stderr
    for expected in ("44 lines read", "41 normalised", "3 rejected",
                     "row 11: quantity is empty",
                     "row 12: duplicate of row 10",
                     "row 13: net amount is zero"):
        assert expected in result.stdout


def test_handover_refuses_a_requisition_with_mixed_suppliers(tmp_path):
    _, _, step4 = run_until_handover(tmp_path, "--input", SAMPLES, "--all")
    assert step4.returncode != 0
    assert "A requisition covers one supplier" in step4.stdout
    assert not list(tmp_path.glob("PO_*.xlsx"))
