"""Value parsing helpers.

Every supplier writes numbers their own way. These functions turn whatever
arrives into the single representation the rest of the pipeline works with:
floats, and tax rates as decimal fractions (5% is 0.05).
"""

import re

TAX_NAMES = {
    "ipi": "ipi",
    "icms": "icms",
    "pis/cofins": "pis_cofins",
    "pis cofins": "pis_cofins",
    "piscofins": "pis_cofins",
    "iss": "iss",
}


def parse_number(value, decimal_style="dot"):
    """Return a float, or None when the cell holds no usable number.

    Handles '1.234,56', '1,234.56', 'BRL 1.145,00', ' 48,90 ' and plain floats.
    """
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    if text in ("", "-", "N/A", "n/a"):
        return None

    text = re.sub(r"[^\d,.\-]", "", text)
    if text in ("", "-"):
        return None

    if decimal_style == "comma":
        text = text.replace(".", "").replace(",", ".")
    else:
        text = text.replace(",", "")

    try:
        return float(text)
    except ValueError:
        return None


def parse_rate(value, tax_style, decimal_style="dot", net_amount=None):
    """Return a tax rate as a decimal fraction: 5% becomes 0.05."""
    if tax_style == "amount":
        amount = parse_number(value, decimal_style)
        if amount is None or not net_amount:
            return 0.0
        return round(amount / net_amount, 6)

    number = parse_number(value, decimal_style)
    if number is None:
        return 0.0

    if tax_style == "percent_text" or tax_style == "percent_number":
        return round(number / 100, 6)
    if tax_style == "decimal_fraction":
        return round(number, 6)
    return 0.0


def parse_combined_taxes(text, decimal_style="dot"):
    """Split 'IPI 5%; ICMS 18%; PIS/COFINS 9.25%' into a rate per tax."""
    rates = {"ipi": 0.0, "icms": 0.0, "pis_cofins": 0.0, "iss": 0.0}
    if not text:
        return rates

    for chunk in str(text).split(";"):
        match = re.match(r"\s*([A-Za-z/ ]+?)\s*([\d.,]+)\s*%", chunk)
        if not match:
            continue
        label = match.group(1).strip().lower()
        key = TAX_NAMES.get(label)
        if not key:
            continue
        value = parse_number(match.group(2), decimal_style)
        if value is not None:
            rates[key] = round(value / 100, 6)
    return rates


def parse_quantity_with_unit(value):
    """'36 pcs' becomes 36. A plain number passes through unchanged."""
    if isinstance(value, (int, float)):
        return float(value)
    if value is None:
        return None
    match = re.search(r"[\d.,]+", str(value))
    return parse_number(match.group(0), "dot") if match else None


def strip_type_prefix(description):
    """'[MAT] Conveyor roller' becomes ('MATERIAL', 'Conveyor roller')."""
    if not description:
        return None, description
    match = re.match(r"\s*\[(MAT|SRV|SERV)\]\s*(.*)", str(description), re.I)
    if not match:
        return None, str(description).strip()
    tag = match.group(1).upper()
    return ("MATERIAL" if tag == "MAT" else "SERVICE"), match.group(2).strip()


def normalise_operation_type(value):
    """Map whatever the supplier calls it to MATERIAL or SERVICE."""
    if not value:
        return None
    text = str(value).strip().upper()
    if text.startswith("MAT"):
        return "MATERIAL"
    if text.startswith("SERV") or text.startswith("SRV"):
        return "SERVICE"
    return None


def clean_text(value):
    """Trim padding and collapse repeated spaces."""
    if value is None:
        return None
    return re.sub(r"\s+", " ", str(value)).strip()
