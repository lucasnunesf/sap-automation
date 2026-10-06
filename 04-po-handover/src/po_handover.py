"""Stage 04 - Purchase order handover.

Reads the requisition workbook once SAP has returned the requisition numbers,
joins it back to the supplier data from stage 01, and produces the purchase
order input files.

This is the handover that used to be done by copy and paste between the
administrative team and the buyer.

Two rules drive what comes out. A requisition belongs to one supplier, so the
script checks that every line agrees on it and refuses to continue when they
do not. And a purchase order covers one operation type, so a requisition
produces at most two files: one for tooling, one for service.

Usage:
    python po_handover.py \
        --requisition ../../02-requisition-builder/src/Input_RC_automatic_filled.xlsx \
        --normalised  ../../01-supplier-intake/src/normalised_items.xlsx
"""

import argparse
import re
import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "shared"))
import parsers as P          # noqa: E402
import item_master           # noqa: E402

ROOT = Path(__file__).resolve().parents[2]
TEMPLATE_PATH = ROOT / "shared" / "Input_PO_automatic.xlsx"

RC_FIRST_ROW = 8
RC_SHEETS = {"Tooling": "MATERIAL", "Service": "SERVICE"}
RC_COL = {"code": 3, "quantity": 4, "value": 5, "description": 6}
RC_CELL_NUMBER = "K2"

PO_FIRST_ROW = 9
PO_COL = {"id": 1, "requisition": 2, "description": 3,
          "quantity": 4, "gross": 5, "net": 6}
PO_CELL = {"delivery_date": "E2", "purchasing_group": "E3",
           "company": "E4", "po_type": "E5", "supplier": "E6"}


# --------------------------------------------------------------------------
# Reading the inputs
# --------------------------------------------------------------------------

def read_requisitions(path):
    """Every line SAP accepted, with the requisition number of its sheet."""
    workbook = openpyxl.load_workbook(path, data_only=True)
    lines = []

    for sheet_name, operation_type in RC_SHEETS.items():
        if sheet_name not in workbook.sheetnames:
            continue
        sheet = workbook[sheet_name]
        number = P.clean_text(sheet[RC_CELL_NUMBER].value)

        for row in range(RC_FIRST_ROW, sheet.max_row + 1):
            code = P.clean_text(sheet.cell(row=row, column=RC_COL["code"]).value)
            if not code:
                continue
            lines.append({
                "internal_code": code,
                "quantity": sheet.cell(row=row, column=RC_COL["quantity"]).value,
                "sap_value": sheet.cell(row=row, column=RC_COL["value"]).value,
                "description": sheet.cell(row=row, column=RC_COL["description"]).value,
                "operation_type": operation_type,
                "requisition_number": number,
                "source_sheet": sheet_name,
                "source_row": row,
            })
    return lines


def read_suppliers(path, master):
    """Internal code -> supplier, taken from the stage 01 table."""
    sheet = openpyxl.load_workbook(path, data_only=True).active
    headers = [cell.value for cell in sheet[1]]
    index = {name: position for position, name in enumerate(headers) if name}

    suppliers = {}
    for values in sheet.iter_rows(min_row=2, values_only=True):
        part = P.clean_text(values[index["part_number"]])
        entry = master.get(part.upper()) if part else None
        if not entry or not entry.get("internal_code"):
            continue
        suppliers[str(entry["internal_code"]).upper()] = {
            "supplier_code": P.clean_text(values[index["supplier_code"]]),
            "supplier_name": P.clean_text(values[index["supplier_name"]]),
            "net_amount": values[index["net_amount"]],
        }
    return suppliers


# --------------------------------------------------------------------------
# Grouping
# --------------------------------------------------------------------------

def group_lines(lines, suppliers):
    """One purchase order per operation type. The requisition belongs to a
    single supplier, so the supplier is a property of the run, not of a group."""
    groups = defaultdict(list)
    orphans = []

    for line in lines:
        entry = suppliers.get(line["internal_code"].upper())
        if not entry or not entry.get("supplier_code"):
            line["problem"] = "no supplier found for this item"
            orphans.append(line)
            continue
        if not line["requisition_number"]:
            line["problem"] = "sheet has no requisition number yet"
            orphans.append(line)
            continue

        line.update(entry)
        groups[line["operation_type"]].append(line)

    return groups, orphans


def single_supplier(groups):
    """A requisition covers one supplier. Return it, or None when the lines
    disagree, which means the requisition was built from mixed sources."""
    found = {line["supplier_code"] for lines in groups.values() for line in lines}
    return found.pop() if len(found) == 1 else None


# --------------------------------------------------------------------------
# Writing one purchase order file
# --------------------------------------------------------------------------

def safe_name(text):
    return re.sub(r"[^A-Za-z0-9_-]+", "_", str(text)).strip("_")


def write_purchase_order(supplier_code, operation_type, lines, out_dir,
                         template=TEMPLATE_PATH):
    workbook = openpyxl.load_workbook(template)
    sheet = workbook.active

    sheet[PO_CELL["po_type"]] = operation_type
    sheet[PO_CELL["supplier"]] = supplier_code
    sheet[PO_CELL["delivery_date"]] = datetime.now().date()

    for position, line in enumerate(lines):
        row = PO_FIRST_ROW + position
        sheet.cell(row=row, column=PO_COL["requisition"],
                   value=line["requisition_number"])
        sheet.cell(row=row, column=PO_COL["description"],
                   value=line["description"])
        sheet.cell(row=row, column=PO_COL["quantity"],
                   value=line["quantity"])
        sheet.cell(row=row, column=PO_COL["gross"],
                   value=line["sap_value"])
        sheet.cell(row=row, column=PO_COL["net"],
                   value=line["net_amount"])

    out_path = Path(out_dir) / f"PO_{safe_name(supplier_code)}_{operation_type}.xlsx"
    workbook.save(out_path)
    return out_path


# --------------------------------------------------------------------------

def write_report(groups, orphans, written, report_path):
    total = sum(len(lines) for lines in groups.values()) + len(orphans)
    report = [
        "Stage 04 - purchase order handover",
        f"Run at {datetime.now():%Y-%m-%d %H:%M}",
        "",
        f"{total} requisition lines read",
        f"{len(groups)} purchase orders generated",
        f"{len(orphans)} lines held back",
    ]

    if written:
        report += ["", "Files written:"]
        report += [f"  {path.name} ({count} lines)" for path, count in written]

    if orphans:
        report += ["", "Lines held back:"]
        for line in orphans:
            where = f"{line['source_sheet']} row {line['source_row']}"
            report.append(f"  {where} ({line['internal_code']}): {line['problem']}")

    text = "\n".join(report)
    Path(report_path).write_text(text, encoding="utf-8")
    return text


def main():
    parser = argparse.ArgumentParser(description="Stage 04 - purchase order handover")
    parser.add_argument("--requisition", required=True,
                        help="requisition workbook with the numbers SAP returned")
    parser.add_argument("--normalised", required=True,
                        help="normalised table from stage 01")
    parser.add_argument("--output-dir", default=".")
    parser.add_argument("--report", default="handover_report.txt")
    args = parser.parse_args()

    master = item_master.load()
    lines = read_requisitions(Path(args.requisition))
    suppliers = read_suppliers(Path(args.normalised), master)
    print(f"  {len(lines)} requisition lines read")

    groups, orphans = group_lines(lines, suppliers)

    supplier_code = single_supplier(groups)
    if groups and not supplier_code:
        suppliers_found = sorted({line["supplier_code"]
                                  for lines in groups.values() for line in lines})
        print()
        print("A requisition covers one supplier, but these lines come from "
              f"{len(suppliers_found)}:")
        for code in suppliers_found:
            print(f"  {code}")
        print()
        print("Run the pipeline once per supplier quotation.")
        return

    written = []
    for operation_type, group in sorted(groups.items()):
        path = write_purchase_order(supplier_code, operation_type, group,
                                    args.output_dir)
        written.append((path, len(group)))

    print()
    print(write_report(groups, orphans, written, args.report))


if __name__ == "__main__":
    main()
