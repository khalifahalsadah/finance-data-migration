"""Cross-match short invoice refs (00331-2025) with full ERP refs (ACC-SINV-2025-00331)."""

import re


def normalize_ref(ref):
    """Extract (year, sequence) from both short and full invoice refs.
    Returns (year, sequence) tuple or None if unparseable.
    """
    if not ref:
        return None
    ref = str(ref).strip()

    # Full format: ACC-SINV-2025-00331
    m = re.match(r'ACC-SINV-(\d{4})-(\d+)', ref)
    if m:
        return (m.group(1), m.group(2).lstrip('0') or '0')

    # Short format: 00331-2025
    m = re.match(r'(\d+)-(\d{4})$', ref)
    if m:
        return (m.group(2), m.group(1).lstrip('0') or '0')

    # Short format without leading zeros: 028-2026
    m = re.match(r'(\d+)-(\d{4})$', ref)
    if m:
        return (m.group(2), m.group(1).lstrip('0') or '0')

    return None


def match_invoice(sheet_ref, erp_invoices):
    """Find matching ERP invoice for a sheet reference.
    erp_invoices: list of dicts with 'name' key (e.g., ACC-SINV-2025-00331).
    Returns matched invoice dict or None.
    """
    sheet_norm = normalize_ref(sheet_ref)
    if not sheet_norm:
        return None

    for inv in erp_invoices:
        erp_norm = normalize_ref(inv['name'])
        if erp_norm and erp_norm == sheet_norm:
            return inv

    return None


def build_invoice_map(sheet_refs, erp_invoices):
    """Build a mapping of sheet refs to ERP invoices.
    Returns: {sheet_ref: erp_invoice_dict or None}
    """
    result = {}
    matched_erp = set()

    for ref in sheet_refs:
        matched = match_invoice(ref, erp_invoices)
        if matched:
            result[ref] = matched
            matched_erp.add(matched['name'])
        else:
            result[ref] = None

    # Track unmatched ERP invoices
    unmatched_erp = [inv for inv in erp_invoices if inv['name'] not in matched_erp]

    return result, unmatched_erp
