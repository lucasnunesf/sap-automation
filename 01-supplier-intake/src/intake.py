"""Stage 01 - Supplier intake.

Reads supplier quotations in whatever layout each supplier uses and writes one
normalised table. This stage knows about suppliers and nothing else: it does
not know the item master, the requisition layout, or anything about SAP.

Nothing about a supplier's layout lives in this file. Adding a supplier means
adding a block to shared/suppliers.yaml.

Usage:
    python intake.py --input ../sample_data/ --all
    python intake.py --input ../sample_data/supplier_1.xlsx
"""

import argparse
import sys
from datetime import datetime
from pathlib import Path

import openpyxl
import yaml

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "shared"))
import parsers as P  # noqa: E402

ROOT = Path(__file__).resolve().parents[2]
CONFIG_PATH = ROOT / "shared" / "suppliers.yaml"

COLUMNS = ["supplier_code", "supplier_name", "project_code", "part_number",
           "description", "operation_type", "quantity", "unit_price",
           "net_amount", "ipi", "icms", "pis_cofins", "iss",
           "source_file", "source_sheet", "source_row"]


def load_config(path=CONFIG_PATH):
    with open(path, encoding="utf-8") as handle:
        return yaml.safe_load(handle)["suppliers"]


# --------------------------------------------------------------------------
# Reading one supplier file
# --------------------------------------------------------------------------

def header_index(sheet, header_row):
    return {P.clean_text(c.value): c.column
            for c in sheet[header_row] if c.value is not None}


def should_skip(raw, rule):
    if not rule:
        return False
    value = P.clean_text(raw.get(rule["column"]))
    if not value:
        return False
    upper = value.upper()
    return any(word.upper() in upper for word in rule["contains"])


def is_blank_line(raw, columns):
    """A row with no part number, no description and no quantity is layout,
    not data: a spacer, a section heading, or a stray total."""
    identity = [raw.get(columns[f]) for f in ("part_number", "description", "quantity")
                if f in columns]
    return all(P.clean_text(v) in (None, "") for v in identity)


def read_supplier(path, config):
    workbook = openpyxl.load_workbook(path, data_only=True)
    decimal_style = config.get("decimal_style", "dot")
    columns = config["columns"]
    rows = []

    file_values = {}
    for field, ref in (config.get("header_cells") or {}).items():
        file_values[field] = P.clean_text(workbook[config["sheets"][0]][ref].value)

    for sheet_name in config["sheets"]:
        sheet = workbook[sheet_name]
        index = header_index(sheet, config["header_row"])

        for excel_row in range(config["header_row"] + 1, sheet.max_row + 1):
            raw = {}
            for field, title in columns.items():
                col = index.get(P.clean_text(title))
                raw[title] = sheet.cell(row=excel_row, column=col).value if col else None

            if all(v is None for v in raw.values()):
                continue
            if should_skip(raw, config.get("skip_rows_when")):
                continue
            if is_blank_line(raw, columns):
                continue

            rows.append(build_row(raw, config, decimal_style, columns,
                                  file_values, path.name, excel_row, sheet_name))
    return rows


def build_row(raw, config, decimal_style, columns, file_values, file_name,
              excel_row, sheet_name):
    composite = config.get("composite") or {}
    get = lambda field: raw.get(columns[field]) if field in columns else None

    description = P.clean_text(get("description"))
    if composite.get("type_prefix_in_description"):
        operation_type, description = P.strip_type_prefix(description)
    else:
        operation_type = P.normalise_operation_type(get("operation_type"))

    if composite.get("quantity_has_unit"):
        quantity = P.parse_quantity_with_unit(get("quantity"))
    else:
        quantity = P.parse_number(get("quantity"), decimal_style)

    unit_price = P.parse_number(get("unit_price"), decimal_style)
    net_amount = P.parse_number(get("net_amount"), decimal_style)
    if net_amount is None and unit_price is not None and quantity is not None:
        net_amount = round(unit_price * quantity, 2)

    tax_style = config["tax_style"]
    if tax_style == "combined_text":
        rates = P.parse_combined_taxes(get("taxes"), decimal_style)
    else:
        rates = {key: P.parse_rate(get(key), tax_style, decimal_style, net_amount)
                 for key in ("ipi", "icms", "pis_cofins", "iss")}

    return {
        "supplier_code": file_values.get("supplier_code") or P.clean_text(get("supplier_code")),
        "supplier_name": file_values.get("supplier_name") or P.clean_text(get("supplier_name")) or config["name"],
        "project_code": file_values.get("project_code") or P.clean_text(get("project_code")),
        "part_number": P.clean_text(get("part_number")),
        "description": description,
        "operation_type": operation_type,
        "quantity": quantity,
        "unit_price": unit_price,
        "net_amount": net_amount,
        **rates,
        "source_file": file_name,
        "source_sheet": sheet_name,
        "source_row": excel_row,
    }


# --------------------------------------------------------------------------
# Validation
# --------------------------------------------------------------------------

def validate(rows):
    """Check what a supplier quotation can be checked against on its own.
    Whether the item exists internally is stage 02's question, not this one's."""
    accepted, rejected = [], []
    seen = {}

    for row in rows:
        problems = []

        if not row["part_number"]:
            problems.append("part number is empty")
        if not row["description"]:
            problems.append("description is empty")
        if row["quantity"] is None:
            problems.append("quantity is empty")
        elif row["quantity"] <= 0:
            problems.append(f"quantity must be positive (got {row['quantity']:g})")
        if row["net_amount"] is None:
            problems.append("net amount could not be read")
        elif row["net_amount"] <= 0:
            problems.append("net amount is zero")

        if row["part_number"]:
            key = (row["source_file"], row["part_number"].upper())
            if key in seen:
                problems.append(f"duplicate of row {seen[key]}")
            else:
                seen[key] = row["source_row"]

        if problems:
            row["problems"] = problems
            rejected.append(row)
        else:
            accepted.append(row)

    return accepted, rejected


# --------------------------------------------------------------------------
# Output
# --------------------------------------------------------------------------

def write_table(rows, out_path):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Normalised"

    sheet.append(COLUMNS)
    for cell in sheet[1]:
        cell.font = openpyxl.styles.Font(name="Arial", size=9, bold=True)
    sheet.freeze_panes = "A2"

    for row in rows:
        sheet.append([row.get(name) for name in COLUMNS])

    for column, width in zip("ABCDEFGHIJKLMNOP",
                             [14, 22, 13, 16, 34, 15, 10, 12, 13, 9, 9, 12, 9, 18, 16, 11]):
        sheet.column_dimensions[column].width = width

    workbook.save(out_path)


def write_report(accepted, rejected, report_path):
    lines = [
        "Stage 01 - supplier intake",
        f"Run at {datetime.now():%Y-%m-%d %H:%M}",
        "",
        f"{len(accepted) + len(rejected)} lines read",
        f"{len(accepted)} normalised",
        f"{len(rejected)} rejected",
    ]
    if rejected:
        lines += ["", "Rejected lines:"]
        for row in sorted(rejected, key=lambda r: (r["source_file"], r["source_row"])):
            where = f"{row['source_file']} [{row['source_sheet']}] row {row['source_row']}"
            lines.append(f"  {where}: {'; '.join(row['problems'])}")

    text = "\n".join(lines)
    Path(report_path).write_text(text, encoding="utf-8")
    return text


def match_config(path, config):
    for key, supplier in config.items():
        if path.match(supplier.get("file_pattern", f"{key}*")):
            return supplier
    return None


def main():
    parser = argparse.ArgumentParser(description="Stage 01 - supplier intake")
    parser.add_argument("--input", required=True, help="supplier file or folder")
    parser.add_argument("--all", action="store_true")
    parser.add_argument("--output", default="normalised_items.xlsx")
    parser.add_argument("--report", default="intake_report.txt")
    args = parser.parse_args()

    config = load_config()
    source = Path(args.input)
    files = sorted(source.glob("*.xlsx")) if args.all or source.is_dir() else [source]

    rows = []
    for path in files:
        supplier = match_config(path, config)
        if not supplier:
            print(f"  skipped {path.name}: no mapping found")
            continue
        found = read_supplier(path, supplier)
        print(f"  {path.name}: {len(found)} lines read")
        rows.extend(found)

    accepted, rejected = validate(rows)
    write_table(accepted, args.output)
    print()
    print(write_report(accepted, rejected, args.report))
    print(f"\nWritten to {args.output}")


if __name__ == "__main__":
    main()
