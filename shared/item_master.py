"""Item master access.

The item master maps a supplier's part number to the internal values a
requisition and a purchase order need: internal code, item category, material
group, budget and project. Stages 02 and 04 both read it, so it lives here
rather than inside either of them.
"""

from pathlib import Path

import openpyxl

import parsers as P

DEFAULT_PATH = Path(__file__).resolve().parent / "master_data.xlsx"
SHEET_NAME = "Item Master"
HEADER_ROW = 4

FIELDS = {
    "internal_code": "Internal Code",
    "item_category": "Item Category",
    "material_group": "Material Group",
    "budget_code": "Budget Code",
    "budget_line": "Budget Line",
    "project_code": "Project Code",
    "operation_type": "Operation Type",
}


def load(path=DEFAULT_PATH):
    """Return the item master keyed by the part number the supplier uses."""
    sheet = openpyxl.load_workbook(path, data_only=True)[SHEET_NAME]
    headers = [cell.value for cell in sheet[HEADER_ROW]]
    index = {name: position for position, name in enumerate(headers) if name}

    master = {}
    for values in sheet.iter_rows(min_row=HEADER_ROW + 1, values_only=True):
        part = P.clean_text(values[index["Supplier Part Number"]])
        if not part:
            continue
        master[part.upper()] = {field: values[index[column]]
                                for field, column in FIELDS.items()}
    return master


def by_internal_code(master):
    """Reverse index, for stages that start from the internal code."""
    return {str(entry["internal_code"]).upper(): entry
            for entry in master.values() if entry.get("internal_code")}
