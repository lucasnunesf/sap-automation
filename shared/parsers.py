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


class NumberFormatError(ValueError):
    """A cell has a value, but not in the format configured for the supplier."""


# What a number may look like in each style: optional minus, digits with
# thousands separators every three digits (or none), optional decimals.
NUMBER_FORMATS = {
    "comma": (re.compile(r"-?(\d{1,3}(\.\d{3})+|\d+)(,\d+)?"), "1.234,56"),
    "dot": (re.compile(r"-?(\d{1,3}(,\d{3})+|\d+)(\.\d+)?"), "1,234.56"),
}
CURRENCY_PREFIX = re.compile(r"^(BRL|R\$|USD|US\$|\$|EUR|€)\s*", re.I)


def parse_number(value, decimal_style="dot"):
    """Return a float, or None when the cell is empty.

    Reads '1.234,56' (comma style), '1,234.56' (dot style), 'BRL 1.145,00',
    ' 48,90 ' and plain numbers. Text that does not match the supplier's
    configured format raises NumberFormatError instead of being guessed.
    """
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    if text in ("", "-", "N/A", "n/a"):
        return None

    number = CURRENCY_PREFIX.sub("", text).strip()
    pattern, example = NUMBER_FORMATS[decimal_style]
    if not pattern.fullmatch(number):
        raise NumberFormatError(f"'{text}' is not in {decimal_style} format ({example})")

    if decimal_style == "comma":
        number = number.replace(".", "").replace(",", ".")
    else:
        number = number.replace(",", "")
    return float(number)


def is_blank(value):
    return value is None or (isinstance(value, str) and not value.strip())


def parse_rate(value, tax_style, decimal_style="dot", net_amount=None):
    """Return a tax rate as a decimal fraction: 5% becomes 0.05.

    A blank cell returns None (the supplier didn't say). "-", "N/A", 0 or 0%
    return 0.0 (the supplier said the tax doesn't apply)."""
    if is_blank(value):
        return None
    if tax_style == "amount":
        amount = parse_number(value, decimal_style)
        if amount is None or not net_amount:
            return 0.0
        return round(amount / net_amount, 6)

    if tax_style not in ("percent_text", "percent_number", "decimal_fraction"):
        return 0.0
    if tax_style == "percent_text" and isinstance(value, str):
        value = value.strip().removesuffix("%")
    number = parse_number(value, decimal_style)
    if number is None:
        return 0.0

    if tax_style == "percent_text" or tax_style == "percent_number":
        return round(number / 100, 6)
    if tax_style == "decimal_fraction":
        return round(number, 6)
    return 0.0


def parse_combined_taxes(text, decimal_style="dot"):
    """Split 'IPI 5%; ICMS 18%; PIS/COFINS 9.25%' into a rate per tax.

    A blank cell means every tax is missing. Otherwise the cell lists the taxes
    that apply, and a tax it doesn't mention is 0."""
    if is_blank(text):
        return {"ipi": None, "icms": None, "pis_cofins": None, "iss": None}
    rates = {"ipi": 0.0, "icms": 0.0, "pis_cofins": 0.0, "iss": 0.0}

    # Find every "NAME number%" in the cell, whatever separates them (";" or ",").
    for match in re.finditer(r"([A-Za-z][A-Za-z/ ]*?)\s*(\d[\d.,]*)\s*%", str(text)):
        label = match.group(1).strip()
        key = TAX_NAMES.get(label.lower())
        if not key:
            raise NumberFormatError(f"unknown tax '{label}'")
        value = parse_number(match.group(2), decimal_style)
        if value is not None:
            rates[key] = round(value / 100, 6)
    return rates


def parse_quantity_with_unit(value, decimal_style="dot"):
    """'36 pcs' becomes 36 and, for a comma supplier, '1,5 kg' becomes 1.5.
    A plain number passes through unchanged. Anything else is an error."""
    if isinstance(value, (int, float)):
        return float(value)
    if value is None or not str(value).strip():
        return None
    match = re.fullmatch(r"\s*(-?\d[\d.,]*)\s*[A-Za-z]*\.?\s*", str(value))
    if not match:
        raise NumberFormatError(f"'{value}' is not in quantity format (36 pcs)")
    return parse_number(match.group(1), decimal_style)


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
