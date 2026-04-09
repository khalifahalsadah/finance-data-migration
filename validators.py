"""Validators for each sheet — compare sheet data against live ERP."""

import re
from config import VALID_WORKFLOW_STATES
from invoice_matcher import match_invoice, build_invoice_map


def _strip_html(text):
    """Strip HTML tags from a string."""
    if not text:
        return ''
    return re.sub(r'<[^>]+>', '', str(text)).strip()


def _erp_action(status):
    """Map validation status to ERP action."""
    return {
        'UPDATE': 'UPDATE',
        'CREATE': 'CREATE',
        'MATCH': 'SKIP',
        'NO_STATUS': 'SKIP',
        'MANDATORY_EMPTY': 'SKIP',
        'MISSING_IN_ERP': 'SKIP',
        'MISSING_IN_SHEET': 'SKIP',
        'BLOCKED': 'SKIP',
        'RELEASE_RESOURCE': 'UPDATE',
    }.get(status, 'SKIP')


SHEET_TO_MODULE = {
    '1. Project Info': 'Project',
    '3. Resources (Planned)': 'Quotation',
    '4. Resources (Actual)': 'Project Employee Distribution',
    '5. Third-Party (Planned)': 'Quotation',
    '6. Third-Party (Actual)': 'Purchase Invoice / Expense Claim / Journal Entry',
    '7. Deliverables & Invoices': 'Task / Sales Invoice',
}


# Map field labels to ERP API fields (for auto-filling erp_target_field)
FIELD_LABEL_TO_API = {
    # Sheet 1
    'Project Status': 'workflow_state',
    'Billability': 'custom_billability_type',
    'OD Reference': 'custom_od_reference_',
    'Project Name': 'project_name',
    'Client Name': 'customer_name',
    'End-Client Name': 'custom_end_client_name',
    'Awarding Date': 'custom_awarded_date',
    'Official Start Date': 'custom_official_start_date',
    'Official End Date': 'official_end_date',
    'Contractuals Link': 'custom_contractuals_link',
    'Partner': 'custom_partner_name_value',
    'Engagement Manager': 'custom_engagement_manager',
    # Sheet 3
    'ratio': 'custom_sg_resources_price.percent',
    'uom': 'custom_sg_resources_price.uom',
    'qty': 'custom_sg_resources_price.qty',
    'rate': 'custom_sg_resources_price.rate',
    'level': 'custom_sg_resources_price.resources_level',
    # Sheet 4
    'name': 'distribution_detail.employee_name',
    'role': 'distribution_detail.designation',
    'start': 'distribution_detail.from_date',
    'end': 'distribution_detail.to_date',
    # 'ratio' already mapped above — context-dependent
    # Sheet 5
    'entry_type': 'entry_type',
    'description': 'team_name',
    'total': 'total',
    # Sheet 6
    'vendor': 'child_table_value',
    'ref': 'reference_form',
    'amount': 'total',
    'date': 'form_date',
    'cost_type': 'reference_doctype',
    # Sheet 7
    'deliverable': 'quoted_deliverables.item',
    'deliverable amount': 'quoted_deliverables.total_amount',
    'invoice number': 'sales_invoice',
    'invoice amount': 'grand_total',
    'invoice status': 'status',
    'invoice date': 'posting_date',
    'invoicing status': 'invoicing_status',
    'planned date': 'planned_invoicing_date',
    'actual invoicing date': 'actual_invoicing_date',
    'payment date': 'payment_date',
    'row status': 'row_status',
}


def _infer_target_field(field_label):
    """Infer ERP target field from the field label."""
    # Extract the part after ' — ' if present
    if ' — ' in field_label:
        sub = field_label.split(' — ')[-1].strip()
        return FIELD_LABEL_TO_API.get(sub, sub)
    return FIELD_LABEL_TO_API.get(field_label, '')


def _result(project, sheet, field, erp_value, sheet_value, status, message='',
            target_doctype='', target_name='', target_field='', child_name='', row_num=''):
    # Auto-fill erp_module from target_doctype or sheet name
    erp_module = target_doctype or SHEET_TO_MODULE.get(sheet, '')
    erp_target_field = target_field or _infer_target_field(field)

    return {
        'project': project,
        'sheet': sheet,
        'field': field,
        'erp_value': erp_value or '',
        'sheet_value': sheet_value or '',
        'status': status,
        'erp_action': _erp_action(status),
        'erp_module': erp_module,
        'erp_target_field': erp_target_field,
        'message': message,
        'target_doctype': target_doctype,
        'target_name': target_name,
        'target_field': target_field,
        'child_name': child_name,
        'row_num': row_num,
    }


def _compare_field(project, sheet, field_label, erp_live, sheet_effective, sheet_status,
                   api_field='', target_doctype='', target_name=''):
    """Core comparison logic for a single field."""
    erp_str = str(erp_live).strip() if erp_live else ''
    sheet_str = str(sheet_effective).strip() if sheet_effective else ''

    if sheet_status == 'Corrected':
        if not sheet_str:
            return _result(project, sheet, field_label, erp_str, sheet_str,
                           'MANDATORY_EMPTY', 'Corrected but no value provided')
        if sheet_str != erp_str:
            return _result(project, sheet, field_label, erp_str, sheet_str,
                           'UPDATE', '', target_doctype, target_name, api_field)
        return _result(project, sheet, field_label, erp_str, sheet_str,
                       'MATCH', 'Already matches ERP')

    if sheet_status == 'Confirmed':
        if sheet_str and sheet_str != erp_str:
            return _result(project, sheet, field_label, erp_str, sheet_str,
                           'UPDATE', 'Confirmed value differs from ERP',
                           target_doctype, target_name, api_field)
        if not sheet_str and not erp_str:
            return _result(project, sheet, field_label, erp_str, sheet_str,
                           'MANDATORY_EMPTY', 'Confirmed but both empty')
        return _result(project, sheet, field_label, erp_str, sheet_str, 'MATCH')

    if sheet_status == 'Missing in ERP':
        if sheet_str:
            return _result(project, sheet, field_label, erp_str, sheet_str,
                           'UPDATE', 'Was missing, now filled',
                           target_doctype, target_name, api_field)
        return _result(project, sheet, field_label, erp_str, sheet_str,
                       'MANDATORY_EMPTY', 'Missing in ERP, no value provided')

    if sheet_status == 'Blocked':
        return _result(project, sheet, field_label, erp_str, sheet_str,
                       'BLOCKED', 'Marked as blocked by Finance')

    # No status
    if not erp_str and not sheet_str:
        return _result(project, sheet, field_label, erp_str, sheet_str,
                       'MANDATORY_EMPTY', 'No status, both empty')
    if erp_str and not sheet_str:
        return _result(project, sheet, field_label, erp_str, sheet_str,
                       'NO_STATUS', 'Not reviewed (ERP has value)')
    if sheet_str and not erp_str:
        return _result(project, sheet, field_label, erp_str, sheet_str,
                       'NO_STATUS', 'Not reviewed (sheet has value, ERP empty)')
    return _result(project, sheet, field_label, erp_str, sheet_str,
                   'NO_STATUS', 'Not reviewed')


# ── Sheet 1: Project Info ────────────────────────────────────────────────────

def validate_project_info(proj_id, sheet_fields, erp_project):
    results = []
    sheet_name = '1. Project Info'

    if not sheet_fields:
        results.append(_result(proj_id, sheet_name, '(all)', '', '', 'MANDATORY_EMPTY',
                               'Project not in sheet'))
        return results

    for api_field, field_data in sheet_fields.items():
        label = field_data['label']
        erp_live = erp_project.get(api_field, '')
        effective = field_data['effective_val']
        status = field_data['status']

        # Special validation for workflow_state
        if api_field == 'workflow_state' and effective and effective not in VALID_WORKFLOW_STATES:
            results.append(_result(
                proj_id, sheet_name, label, erp_live or '', effective,
                'BLOCKED', f'"{effective}" is not a valid workflow state',
            ))
            continue

        r = _compare_field(
            proj_id, sheet_name, label, erp_live, effective, status,
            api_field=api_field, target_doctype='Project', target_name=proj_id,
        )
        results.append(r)

    return results


# ── Sheet 3: Resources Planned ───────────────────────────────────────────────

def validate_resources_planned(proj_id, sheet_rows, quotation, quot_note=''):
    results = []
    sheet_name = '3. Resources (Planned)'

    # Filter out empty/placeholder rows
    data_rows = [r for r in sheet_rows if r.get('level') and r['level'] != 'not found']

    if not data_rows:
        results.append(_result(proj_id, sheet_name, '(all)', '', '',
                               'MANDATORY_EMPTY', 'No planned resource data'))
        return results

    quot_ref = data_rows[0].get('quotation')

    # Check quotation lookup notes
    if quot_note and quot_note.startswith('WRONG_PROJECT:'):
        results.append(_result(proj_id, sheet_name, 'Quotation', '', quot_ref,
                               'BLOCKED', f'{quot_note} — Quotation belongs to another project',
                               target_doctype='Quotation'))
        # Skip all field comparisons — wrong quotation
        for i, row in enumerate(data_rows, 1):
            _check_planned_row_mandatory(results, proj_id, sheet_name, i, row)
        return results

    if quot_note and quot_note.startswith('NOT_LINKED:'):
        results.append(_result(proj_id, sheet_name, 'Quotation link', quot_ref or '', quot_note,
                               'MATCH', f'Quotation found but {quot_note}',
                               target_doctype='Quotation'))

    if quot_ref and not quotation:
        results.append(_result(proj_id, sheet_name, 'Quotation', '', quot_ref,
                               'MISSING_IN_ERP', f'Quotation {quot_ref} not found in ERP'))
        # Still show all sheet rows with their values (no ERP to compare against)
        for i, row in enumerate(data_rows, 1):
            row_label = f'Row {i} ({row.get("level", "?")})'
            _check_planned_row_mandatory(results, proj_id, sheet_name, i, row)
            for field_name in ['ratio', 'uom', 'qty', 'rate']:
                s_corr = row.get(f'{field_name}_corr') or ''
                s_erp_col = row.get(f'{field_name}_erp') or ''
                sheet_val = s_corr or s_erp_col or ''
                if sheet_val and sheet_val != 'not found':
                    results.append(_result(
                        proj_id, sheet_name, f'{row_label} — {field_name}',
                        '(no Quotation)', sheet_val,
                        'MISSING_IN_ERP', 'Quotation not found — cannot compare',
                    ))
            if not row.get('status'):
                results.append(_result(proj_id, sheet_name, f'{row_label} — status',
                                       '', '', 'NO_STATUS', 'Row not reviewed'))
        return results

    # Get ERP resource pricing from Quotation child table
    erp_resources = []
    if quotation:
        erp_resources = quotation.get('custom_sg_resources_price', [])

    # Match sheet rows to ERP rows by level name
    matched_erp = set()
    matched_sheet = set()

    for si, sr in enumerate(data_rows):
        s_level = (sr.get('level') or '').lower().strip()
        best = None

        for ei, er in enumerate(erp_resources):
            if ei in matched_erp:
                continue
            e_level = (er.get('resources_level') or '').lower().strip()
            if s_level == e_level:
                best = (ei, er)
                break

        row_label = f'Row {si + 1} ({sr["level"]})'

        # Mandatory field check
        _check_planned_row_mandatory(results, proj_id, sheet_name, si + 1, sr)

        if best:
            ei, er = best
            matched_erp.add(ei)
            matched_sheet.add(si)

            # Compare each field
            field_pairs = [
                ('ratio', sr.get('ratio'), er.get('percent'), 'percent'),
                ('uom', sr.get('uom'), er.get('uom'), 'uom'),
                ('qty', sr.get('qty'), er.get('qty'), 'qty'),
                ('rate', sr.get('rate'), er.get('rate'), 'rate'),
            ]

            row_status = sr.get('status')

            for field_name, s_val, e_val, erp_field in field_pairs:
                e_str = str(e_val).strip() if e_val else ''
                s_corr = str(sr.get(f'{field_name}_corr') or '').strip()
                s_erp_col = str(sr.get(f'{field_name}_erp') or '').strip()

                # Determine what the sheet is saying
                if s_corr:
                    # Finance filled in a corrected value — compare against live ERP
                    sheet_val = s_corr
                    _compare_planned_field(results, proj_id, sheet_name, row_label,
                                           field_name, e_str, sheet_val, row_status,
                                           quot_ref, erp_field)
                elif row_status == 'Confirmed':
                    # No correction, Finance says ERP is correct — verify it matches
                    sheet_val = s_erp_col or ''
                    _compare_planned_field(results, proj_id, sheet_name, row_label,
                                           field_name, e_str, sheet_val, row_status,
                                           quot_ref, erp_field)
                else:
                    # Corrected column empty AND no Confirmed status — not reviewed
                    if s_erp_col and s_erp_col != 'not found':
                        results.append(_result(
                            proj_id, sheet_name, f'{row_label} — {field_name}',
                            e_str or '(empty)', f'(not reviewed) ERP col={s_erp_col}',
                            'NO_STATUS', 'Corrected column empty, no status set',
                        ))
                    elif not e_str:
                        results.append(_result(
                            proj_id, sheet_name, f'{row_label} — {field_name}',
                            '(empty)', '(empty)',
                            'MANDATORY_EMPTY', 'Both ERP and sheet empty',
                        ))
                    else:
                        results.append(_result(
                            proj_id, sheet_name, f'{row_label} — {field_name}',
                            e_str, '(not reviewed)',
                            'NO_STATUS', 'Corrected column empty, no status set',
                        ))

            # Row status check
            if not sr.get('status'):
                results.append(_result(proj_id, sheet_name, f'{row_label} — status',
                                       '', '', 'NO_STATUS', 'Row not reviewed'))
        else:
            # Sheet row not found in ERP — show each field with CREATE status
            for field_name in ['ratio', 'uom', 'qty', 'rate']:
                s_corr = sr.get(f'{field_name}_corr') or ''
                s_erp_col = sr.get(f'{field_name}_erp') or ''
                sheet_val = s_corr or s_erp_col or ''
                if sheet_val and sheet_val != 'not found':
                    results.append(_result(
                        proj_id, sheet_name, f'{row_label} — {field_name}',
                        '(not in ERP)', sheet_val,
                        'CREATE', 'New resource — not in Quotation',
                        'Quotation', quot_ref or '', f'custom_sg_resources_price.{field_name}',
                    ))
                elif not sheet_val or sheet_val == 'not found':
                    results.append(_result(
                        proj_id, sheet_name, f'{row_label} — {field_name}',
                        '(not in ERP)', '(empty)',
                        'MANDATORY_EMPTY', f'{field_name} missing for new resource',
                    ))

            if not sr.get('status'):
                results.append(_result(proj_id, sheet_name, f'{row_label} — status',
                                       '', '', 'NO_STATUS', 'Row not reviewed'))

    # ERP rows not in sheet
    for ei, er in enumerate(erp_resources):
        if ei not in matched_erp:
            results.append(_result(
                proj_id, sheet_name,
                f'{er.get("resources_level")} ({er.get("percent")}% qty={er.get("qty")} rate={er.get("rate")})',
                f'In Quotation', '',
                'MISSING_IN_SHEET', 'ERP resource not in sheet',
            ))

    return results


def _compare_planned_field(results, proj_id, sheet_name, row_label,
                           field_name, erp_live, sheet_val, row_status,
                           quot_ref, erp_field):
    """Compare a single planned resource field against live ERP Quotation."""
    e_str = erp_live or ''
    s_str = sheet_val or ''

    # Numeric comparison
    try:
        s_num = float(s_str) if s_str else None
        e_num = float(e_str) if e_str else None
        if s_num is not None and e_num is not None:
            if abs(s_num - e_num) > 0.01:
                results.append(_result(
                    proj_id, sheet_name, f'{row_label} — {field_name}',
                    e_str, s_str, 'UPDATE', '',
                    'Quotation', quot_ref or '', f'custom_sg_resources_price.{erp_field}',
                ))
            else:
                results.append(_result(
                    proj_id, sheet_name, f'{row_label} — {field_name}',
                    e_str, s_str, 'MATCH',
                ))
            return
    except (ValueError, TypeError):
        pass

    # String comparison
    if s_str and e_str and s_str.lower().strip() != e_str.lower().strip():
        results.append(_result(
            proj_id, sheet_name, f'{row_label} — {field_name}',
            e_str, s_str, 'UPDATE', '',
            'Quotation', quot_ref or '', f'custom_sg_resources_price.{erp_field}',
        ))
    elif s_str or e_str:
        results.append(_result(
            proj_id, sheet_name, f'{row_label} — {field_name}',
            e_str or '(empty)', s_str or '(empty)', 'MATCH',
        ))


def _check_planned_row_mandatory(results, proj_id, sheet_name, row_num, row):
    """Check mandatory fields for a planned resource row."""
    row_label = f'Row {row_num} ({row.get("level", "?")})'
    for field in ['level', 'ratio', 'uom', 'qty', 'rate']:
        if not row.get(field) or row[field] == 'not found':
            results.append(_result(proj_id, sheet_name, f'{row_label} — {field}',
                                   '', '', 'MANDATORY_EMPTY', f'{field} is missing'))


# ── Sheet 4: Resources Actual (PED) ─────────────────────────────────────────

def validate_resources_actual(proj_id, sheet_rows, ped_details):
    results = []
    sheet_name = '4. Resources (Actual)'

    data_rows = [r for r in sheet_rows
                 if r.get('name') and r['name'] not in ('not found', 'Cancelled')]
    cancelled_rows = [r for r in sheet_rows if r.get('name') == 'Cancelled' or r.get('end') == 'Cancelled']

    if not data_rows and not cancelled_rows:
        if not ped_details:
            results.append(_result(proj_id, sheet_name, '(all)', '', '',
                                   'MANDATORY_EMPTY', 'No resource data in sheet or PED'))
        else:
            results.append(_result(proj_id, sheet_name, '(all)', f'{len(ped_details)} PED rows', '',
                                   'MANDATORY_EMPTY', 'Sheet empty but PED has data'))
        return results

    if not ped_details and data_rows:
        results.append(_result(proj_id, sheet_name, 'PED', '', '',
                               'MISSING_IN_ERP', 'No PED record found for this project'))

    # Match sheet rows to PED rows
    matched_ped = set()

    for si, sr in enumerate(data_rows):
        best = None
        s_name = (sr['name'] or '').lower().split()[0] if sr.get('name') else ''

        for pi, pd in enumerate(ped_details):
            if pi in matched_ped:
                continue
            p_name = (pd.get('employee_name') or '').lower().split()[0]
            if s_name and p_name and s_name == p_name and sr.get('start') == pd.get('from_date'):
                best = (pi, pd)
                break

        row_label = f'{sr["name"]} ({sr.get("start", "?")})'
        row_status = sr.get('status')

        if best:
            pi, pd = best
            matched_ped.add(pi)
            ped_id = pd.get('_ped_id', '')

            # Field-by-field comparison: name, role, start, end, ratio
            # PED fields: employee_name, designation, from_date, to_date, ratio_
            ped_field_map = [
                ('name', 'employee_name', 'employee_name'),
                ('role', 'designation', 'designation'),
                ('start', 'from_date', 'from_date'),
                ('end', 'to_date', 'to_date'),
                ('ratio', 'ratio_', 'ratio_'),
            ]

            for field_name, ped_field, ped_api_field in ped_field_map:
                e_val = str(pd.get(ped_field) or '').strip()
                s_corr = str(sr.get(f'{field_name}_corr') or '').strip()
                s_erp_col = str(sr.get(f'{field_name}_erp') or '').strip()

                if s_corr:
                    # Finance provided a corrected value
                    _compare_actual_field(
                        results, proj_id, sheet_name, row_label,
                        field_name, e_val, s_corr, row_status,
                        ped_id, ped_api_field, pd['name'], si + 1,
                    )
                elif row_status == 'Confirmed':
                    # No correction, Finance says ERP is correct
                    _compare_actual_field(
                        results, proj_id, sheet_name, row_label,
                        field_name, e_val, s_erp_col, row_status,
                        ped_id, ped_api_field, pd['name'], si + 1,
                    )
                else:
                    # Not reviewed — show what we have
                    if s_erp_col and s_erp_col != 'not found':
                        results.append(_result(
                            proj_id, sheet_name, f'{row_label} — {field_name}',
                            e_val or '(empty)', f'(not reviewed) ERP col={s_erp_col}',
                            'NO_STATUS', 'Corrected column empty, no status set',
                            row_num=si + 1,
                        ))
                    elif not e_val:
                        results.append(_result(
                            proj_id, sheet_name, f'{row_label} — {field_name}',
                            '(empty)', '(empty)',
                            'MANDATORY_EMPTY', 'Both PED and sheet empty',
                            row_num=si + 1,
                        ))
                    else:
                        results.append(_result(
                            proj_id, sheet_name, f'{row_label} — {field_name}',
                            e_val, '(not reviewed)',
                            'NO_STATUS', 'Corrected column empty, no status set',
                            row_num=si + 1,
                        ))

            # Row status check
            if not row_status:
                results.append(_result(proj_id, sheet_name,
                                       f'{row_label} — row status',
                                       '', '', 'NO_STATUS', 'Row not reviewed',
                                       row_num=si + 1))

        else:
            # Sheet row not found in PED — show each field with CREATE
            for field_name in ['name', 'role', 'start', 'end', 'ratio']:
                s_corr = sr.get(f'{field_name}_corr') or ''
                s_erp_col = sr.get(f'{field_name}_erp') or ''
                sheet_val = s_corr or s_erp_col or ''
                if sheet_val and sheet_val != 'not found':
                    results.append(_result(
                        proj_id, sheet_name, f'{row_label} — {field_name}',
                        '(not in PED)', sheet_val,
                        'CREATE', 'New resource — not in PED',
                        'Project Employee Distribution', '', field_name,
                        row_num=si + 1,
                    ))
                elif not sheet_val or sheet_val == 'not found':
                    results.append(_result(
                        proj_id, sheet_name, f'{row_label} — {field_name}',
                        '(not in PED)', '(empty)',
                        'MANDATORY_EMPTY', f'{field_name} missing for new resource',
                        row_num=si + 1,
                    ))

            if not row_status:
                results.append(_result(proj_id, sheet_name, f'{row_label} — row status',
                                       '', '', 'NO_STATUS', 'Row not reviewed',
                                       row_num=si + 1))

    # PED rows not in sheet — show each field
    for pi, pd in enumerate(ped_details):
        if pi not in matched_ped:
            ped_label = f'{pd["employee_name"]} ({pd["from_date"]})'
            for field_name, ped_field in [('name', 'employee_name'), ('role', 'designation'),
                                           ('start', 'from_date'), ('end', 'to_date'),
                                           ('ratio', 'ratio_')]:
                e_val = str(pd.get(ped_field) or '').strip()
                if e_val:
                    results.append(_result(
                        proj_id, sheet_name, f'{ped_label} — {field_name}',
                        e_val, '(not in sheet)',
                        'MISSING_IN_SHEET', 'PED row not in sheet',
                    ))

    # Cancelled rows
    for cr in cancelled_rows:
        results.append(_result(proj_id, sheet_name, 'Cancelled row',
                               '', 'Cancelled', 'RELEASE_RESOURCE',
                               'Resource marked for removal'))

    return results


def _compare_actual_field(results, proj_id, sheet_name, row_label,
                          field_name, erp_live, sheet_val, row_status,
                          ped_id, ped_api_field, child_name, row_num):
    """Compare a single actual resource field against live PED."""
    e_str = erp_live or ''
    s_str = sheet_val or ''

    # Numeric comparison
    try:
        s_num = float(s_str) if s_str else None
        e_num = float(e_str) if e_str else None
        if s_num is not None and e_num is not None:
            if abs(s_num - e_num) > 0.01:
                results.append(_result(
                    proj_id, sheet_name, f'{row_label} — {field_name}',
                    e_str, s_str, 'UPDATE', '',
                    'Project Employee Distribution', ped_id, ped_api_field, child_name,
                    row_num=row_num,
                ))
            else:
                results.append(_result(
                    proj_id, sheet_name, f'{row_label} — {field_name}',
                    e_str, s_str, 'MATCH', row_num=row_num,
                ))
            return
    except (ValueError, TypeError):
        pass

    # String comparison
    if s_str and e_str and s_str.lower().strip() != e_str.lower().strip():
        results.append(_result(
            proj_id, sheet_name, f'{row_label} — {field_name}',
            e_str, s_str, 'UPDATE', '',
            'Project Employee Distribution', ped_id, ped_api_field, child_name,
            row_num=row_num,
        ))
    elif s_str or e_str:
        results.append(_result(
            proj_id, sheet_name, f'{row_label} — {field_name}',
            e_str or '(empty)', s_str or '(empty)', 'MATCH', row_num=row_num,
        ))


# ── Sheet 5: Third-Party Planned ─────────────────────────────────────────────

def validate_thirdparty_planned(proj_id, sheet_rows, quotation, quot_note=''):
    results = []
    sheet_name = '5. Third-Party (Planned)'

    data_rows = [r for r in sheet_rows if r.get('entry_type') and r['entry_type'] != 'not found']

    if not data_rows:
        results.append(_result(proj_id, sheet_name, '(all)', '', '',
                               'MANDATORY_EMPTY', 'No third-party planned data'))
        return results

    # Check quotation lookup notes
    if quot_note and quot_note.startswith('WRONG_PROJECT:'):
        results.append(_result(proj_id, sheet_name, 'Quotation', '', '',
                               'BLOCKED', f'{quot_note} — Quotation belongs to another project',
                               target_doctype='Quotation'))
        return results

    if quot_note and quot_note.startswith('NOT_LINKED:'):
        quot_ref = data_rows[0].get('quotation') if data_rows else ''
        results.append(_result(proj_id, sheet_name, 'Quotation link', quot_ref or '', quot_note,
                               'MATCH', f'Quotation found but {quot_note}',
                               target_doctype='Quotation'))

    # Build ERP third-party data from Quotation child tables
    erp_3p = []
    if quotation:
        for table_key, entry_type_prefix in [
            ('custom_other', 'Third Party Cost - Saudi'),
            ('custom_tpc__osa', 'Third Party Cost - Outside Saudi'),
            ('custom_travel_cost__egypt', 'Travel Cost - Egypt Team'),
            ('custom_travel_cost__local', 'Travel Cost - Local Team'),
        ]:
            for row in quotation.get(table_key, []):
                erp_3p.append({
                    'entry_type': entry_type_prefix,
                    'description': row.get('team_name', ''),
                    'rate': row.get('rate', 0) or row.get('currency_rate', 0) or 0,
                    'qty': row.get('months', 0) or 0,
                    'total': row.get('total', 0) or row.get('amount', 0) or 0,
                    '_table': table_key,
                })
        for row in quotation.get('taxes', []):
            erp_3p.append({
                'entry_type': 'Sales Tax & Charges',
                'description': row.get('description', 'Tax'),
                'rate': row.get('rate', 0) or 0,
                'qty': 0,
                'total': row.get('tax_amount', 0) or 0,
                '_table': 'taxes',
            })

    quot_ref = data_rows[0].get('quotation') if data_rows else None

    # Match sheet rows to ERP rows by entry_type + description
    matched_erp = set()

    for i, sr in enumerate(data_rows, 1):
        s_type = (sr.get('entry_type') or '').strip()
        s_desc = (sr.get('description') or '').strip()
        row_label = f'Row {i} ({s_type} / {s_desc})'
        row_status = sr.get('status')

        # Find matching ERP row
        best = None
        for ei, er in enumerate(erp_3p):
            if ei in matched_erp:
                continue
            if er['entry_type'] == s_type and er['description'].lower().strip() == s_desc.lower().strip():
                best = (ei, er)
                break

        if best:
            ei, er = best
            matched_erp.add(ei)

            # Field-by-field comparison
            # Sheet fields: entry_type, description, rate, qty, uom, total
            # ERP fields: entry_type, description, rate, qty(months), total
            field_pairs = [
                ('entry_type', 'entry_type', 'entry_type'),
                ('description', 'description', 'team_name'),
                ('rate', 'rate', 'rate'),
                ('qty', 'qty', 'months'),
                ('total', 'total', 'total'),
            ]

            for field_name, erp_key, erp_api_field in field_pairs:
                e_val = str(er.get(erp_key) or '').strip()
                s_corr = str(sr.get(f'{field_name}_corr') or '').strip()
                s_erp_col = str(sr.get(f'{field_name}_erp') or '').strip()

                if s_corr:
                    _compare_3p_planned_field(
                        results, proj_id, sheet_name, row_label,
                        field_name, e_val, s_corr, row_status,
                        quot_ref, er['_table'], erp_api_field,
                    )
                elif row_status == 'Confirmed':
                    _compare_3p_planned_field(
                        results, proj_id, sheet_name, row_label,
                        field_name, e_val, s_erp_col, row_status,
                        quot_ref, er['_table'], erp_api_field,
                    )
                else:
                    if s_erp_col and s_erp_col != 'not found':
                        results.append(_result(
                            proj_id, sheet_name, f'{row_label} — {field_name}',
                            e_val or '(empty)', f'(not reviewed) ERP col={s_erp_col}',
                            'NO_STATUS', 'Corrected column empty, no status set',
                        ))
                    elif not e_val or e_val == '0':
                        pass  # both zero/empty, skip
                    else:
                        results.append(_result(
                            proj_id, sheet_name, f'{row_label} — {field_name}',
                            e_val, '(not reviewed)',
                            'NO_STATUS', 'Corrected column empty, no status set',
                        ))

            # UOM (sheet-only field, no ERP equivalent)
            s_uom_corr = sr.get('uom_corr') or ''
            s_uom_erp = sr.get('uom_erp') or ''
            uom_val = s_uom_corr or s_uom_erp or ''
            if uom_val and uom_val != 'not found':
                results.append(_result(
                    proj_id, sheet_name, f'{row_label} — uom',
                    '(no ERP field)', uom_val, 'MATCH', 'UOM is sheet-only',
                ))

        else:
            # Not in ERP — show each field with CREATE
            for field_name in ['entry_type', 'description', 'rate', 'qty', 'uom', 'total']:
                s_corr = sr.get(f'{field_name}_corr') or ''
                s_erp_col = sr.get(f'{field_name}_erp') or ''
                sheet_val = s_corr or s_erp_col or ''
                if sheet_val and sheet_val != 'not found' and sheet_val != '0':
                    results.append(_result(
                        proj_id, sheet_name, f'{row_label} — {field_name}',
                        '(not in ERP)', sheet_val,
                        'CREATE', 'New entry — not in Quotation',
                        'Quotation', quot_ref or '', field_name,
                    ))
                elif not sheet_val or sheet_val in ('not found', '0'):
                    pass  # zero/empty for a missing row is fine

        # Row status check
        if not row_status:
            results.append(_result(proj_id, sheet_name, f'{row_label} — row status',
                                   '', '', 'NO_STATUS', 'Row not reviewed'))

    # ERP rows not in sheet — show each field
    for ei, er in enumerate(erp_3p):
        if ei not in matched_erp:
            erp_label = f'{er["entry_type"]} / {er["description"]}'
            has_value = False
            for field_name, erp_key in [('entry_type', 'entry_type'), ('description', 'description'),
                                         ('rate', 'rate'), ('qty', 'qty'), ('total', 'total')]:
                e_val = er.get(erp_key)
                if e_val and float(e_val or 0) != 0 if isinstance(e_val, (int, float)) else e_val:
                    results.append(_result(
                        proj_id, sheet_name, f'{erp_label} — {field_name}',
                        str(e_val), '(not in sheet)',
                        'MISSING_IN_SHEET', 'Quotation entry not in sheet',
                    ))
                    has_value = True

    return results


def _compare_3p_planned_field(results, proj_id, sheet_name, row_label,
                               field_name, erp_live, sheet_val, row_status,
                               quot_ref, table_key, erp_api_field):
    """Compare a single third-party planned field against live ERP Quotation."""
    e_str = erp_live or ''
    s_str = sheet_val or ''

    # Numeric comparison
    try:
        s_num = float(s_str) if s_str else None
        e_num = float(e_str) if e_str else None
        if s_num is not None and e_num is not None:
            if abs(s_num - e_num) > 0.01:
                results.append(_result(
                    proj_id, sheet_name, f'{row_label} — {field_name}',
                    e_str, s_str, 'UPDATE', '',
                    'Quotation', quot_ref or '', f'{table_key}.{erp_api_field}',
                ))
            else:
                results.append(_result(
                    proj_id, sheet_name, f'{row_label} — {field_name}',
                    e_str, s_str, 'MATCH',
                ))
            return
    except (ValueError, TypeError):
        pass

    # String comparison
    if s_str and e_str and s_str.lower().strip() != e_str.lower().strip():
        results.append(_result(
            proj_id, sheet_name, f'{row_label} — {field_name}',
            e_str, s_str, 'UPDATE', '',
            'Quotation', quot_ref or '', f'{table_key}.{erp_api_field}',
        ))
    elif s_str or e_str:
        results.append(_result(
            proj_id, sheet_name, f'{row_label} — {field_name}',
            e_str or '(empty)', s_str or '(empty)', 'MATCH',
        ))


def _resolve_expense_doctype(ref):
    """Determine the ERP doctype from an expense reference prefix."""
    if not ref:
        return ''
    ref = str(ref).strip()
    if ref.startswith('ACC-PINV'):
        return 'Purchase Invoice'
    if ref.startswith('HR-EXP'):
        return 'Expense Claim'
    if ref.startswith('ACC-JV'):
        return 'Journal Entry'
    return ''


# ── Sheet 6: Third-Party Actual ──────────────────────────────────────────────

def validate_thirdparty_actual(proj_id, sheet_rows, ppi_expenses):
    results = []
    sheet_name = '6. Third-Party (Actual)'

    data_rows = [r for r in sheet_rows
                 if r.get('vendor') or (r.get('amount') and r['amount'] not in ('0', '0.0'))]
    empty_rows = [r for r in sheet_rows if not r.get('vendor') and not r.get('amount')]

    # Build ref maps — match by invoice/expense reference
    erp_refs = {}
    for e in ppi_expenses:
        erp_refs[e['reference_form']] = e

    # Match sheet rows to ERP by reference
    matched_erp = set()

    for i, sr in enumerate(data_rows, 1):
        ref = sr.get('ref')
        row_status = sr.get('status')
        vendor_display = (sr.get('vendor') or '?')[:30]
        row_label = f'Row {i} ({ref or "no ref"} / {vendor_display})'

        er = erp_refs.get(ref) if ref else None
        if er:
            matched_erp.add(ref)

        # Resolve the actual ERP doctype from the reference prefix or ERP data
        if er:
            row_doctype = er.get('reference_doctype', '') or _resolve_expense_doctype(ref)
        else:
            row_doctype = _resolve_expense_doctype(ref)

        # Field-by-field comparison
        field_pairs = [
            ('cost_type', None, 'reference_doctype'),
            ('vendor', er.get('child_table_value', '') if er else '', 'description'),
            ('ref', er.get('reference_form', '') if er else '', 'name'),
            ('amount', str(er.get('total', '')) if er else '', 'total'),
            ('date', er.get('form_date', '') if er else '', 'posting_date'),
        ]

        for field_name, e_val, erp_api_field in field_pairs:
            if field_name == 'cost_type':
                ct = sr.get('cost_type') or ''
                if er:
                    e_doctype = er.get('reference_doctype', '')
                    results.append(_result(
                        proj_id, sheet_name, f'{row_label} — cost_type',
                        e_doctype, ct, 'MATCH' if ct else 'NO_STATUS',
                        target_doctype=row_doctype, target_field='reference_doctype',
                    ))
                elif ct:
                    results.append(_result(
                        proj_id, sheet_name, f'{row_label} — cost_type',
                        '(not in ERP)', ct, 'CREATE' if ref else 'MISSING_IN_ERP',
                        target_doctype=row_doctype, target_field='reference_doctype',
                    ))
                continue

            e_str = str(e_val).strip() if e_val else ''
            s_corr = str(sr.get(f'{field_name}_corr') or '').strip()
            s_erp_col = str(sr.get(f'{field_name}_erp') or '').strip()

            if s_corr:
                _compare_3p_actual_field(
                    results, proj_id, sheet_name, row_label,
                    field_name, e_str, s_corr, row_status, erp_api_field,
                    row_doctype,
                )
            elif row_status == 'Confirmed':
                _compare_3p_actual_field(
                    results, proj_id, sheet_name, row_label,
                    field_name, e_str, s_erp_col, row_status, erp_api_field,
                    row_doctype,
                )
            else:
                if s_erp_col and s_erp_col != 'not found':
                    results.append(_result(
                        proj_id, sheet_name, f'{row_label} — {field_name}',
                        e_str or '(empty)', f'(not reviewed) ERP col={s_erp_col}',
                        'NO_STATUS', 'Corrected column empty, no status set',
                        target_doctype=row_doctype, target_field=erp_api_field,
                    ))
                elif e_str:
                    results.append(_result(
                        proj_id, sheet_name, f'{row_label} — {field_name}',
                        e_str, '(not reviewed)',
                        'NO_STATUS', 'Corrected column empty, no status set',
                        target_doctype=row_doctype, target_field=erp_api_field,
                    ))
                elif not er:
                    results.append(_result(
                        proj_id, sheet_name, f'{row_label} — {field_name}',
                        '(not in ERP)', '(empty)',
                        'MANDATORY_EMPTY', f'{field_name} missing',
                        target_doctype=row_doctype, target_field=erp_api_field,
                    ))

        # Row status check
        if not row_status:
            results.append(_result(proj_id, sheet_name, f'{row_label} — row status',
                                   '', '', 'NO_STATUS', 'Row not reviewed',
                                   target_doctype=row_doctype))

    # ERP expenses not in sheet — show each field with resolved doctype
    for ref, er in erp_refs.items():
        if ref not in matched_erp:
            erp_label = f'{ref}'
            erp_doctype = er.get('reference_doctype', '') or _resolve_expense_doctype(ref)
            for field_name, erp_key, target_f in [
                ('ref', 'reference_form', 'name'),
                ('doctype', 'reference_doctype', 'reference_doctype'),
                ('amount', 'total', 'total'),
                ('date', 'form_date', 'posting_date'),
            ]:
                e_val = er.get(erp_key)
                if e_val:
                    e_display = f'{e_val:,.2f}' if isinstance(e_val, (int, float)) else str(e_val)
                    results.append(_result(
                        proj_id, sheet_name, f'{erp_label} — {field_name}',
                        e_display, '(not in sheet)',
                        'MISSING_IN_SHEET', 'ERP expense not in sheet',
                        target_doctype=erp_doctype, target_field=target_f,
                    ))

    # Empty rows handling
    if not data_rows and not ppi_expenses:
        confirmed_empty = [r for r in empty_rows if r.get('status') == 'Confirmed']
        if confirmed_empty:
            results.append(_result(proj_id, sheet_name, '(all)', '', '',
                                   'MATCH', 'No third-party costs (confirmed)'))
        elif empty_rows:
            results.append(_result(proj_id, sheet_name, '(all)', '', '',
                                   'NO_STATUS', 'Empty rows not reviewed'))
        else:
            results.append(_result(proj_id, sheet_name, '(all)', '', '',
                                   'MANDATORY_EMPTY', 'No data in sheet or ERP'))
    elif not data_rows and ppi_expenses:
        for ref, er in erp_refs.items():
            erp_label = f'{ref}'
            erp_doctype = er.get('reference_doctype', '') or _resolve_expense_doctype(ref)
            for field_name, erp_key, target_f in [
                ('ref', 'reference_form', 'name'),
                ('amount', 'total', 'total'),
                ('date', 'form_date', 'posting_date'),
            ]:
                e_val = er.get(erp_key)
                if e_val:
                    e_display = f'{e_val:,.2f}' if isinstance(e_val, (int, float)) else str(e_val)
                    results.append(_result(
                        proj_id, sheet_name, f'{erp_label} — {field_name}',
                        e_display, '(not in sheet)',
                        'MISSING_IN_SHEET', 'ERP expense not in sheet',
                        target_doctype=erp_doctype, target_field=target_f,
                    ))

    return results


def _compare_3p_actual_field(results, proj_id, sheet_name, row_label,
                              field_name, erp_live, sheet_val, row_status, erp_api_field,
                              target_doctype=''):
    """Compare a single third-party actual field against PPI expense."""
    e_str = erp_live or ''
    s_str = sheet_val or ''

    try:
        s_num = float(s_str) if s_str else None
        e_num = float(e_str) if e_str else None
        if s_num is not None and e_num is not None:
            if abs(s_num - e_num) > 0.01:
                results.append(_result(
                    proj_id, sheet_name, f'{row_label} — {field_name}',
                    e_str, s_str, 'UPDATE', '',
                    target_doctype=target_doctype, target_field=erp_api_field,
                ))
            else:
                results.append(_result(
                    proj_id, sheet_name, f'{row_label} — {field_name}',
                    e_str, s_str, 'MATCH',
                    target_doctype=target_doctype, target_field=erp_api_field,
                ))
            return
    except (ValueError, TypeError):
        pass

    if s_str and e_str and s_str.lower().strip() != e_str.lower().strip():
        results.append(_result(
            proj_id, sheet_name, f'{row_label} — {field_name}',
            e_str, s_str, 'UPDATE', '',
            target_doctype=target_doctype, target_field=erp_api_field,
        ))
    elif s_str or e_str:
        results.append(_result(
            proj_id, sheet_name, f'{row_label} — {field_name}',
            e_str or '(empty)', s_str or '(empty)', 'MATCH',
            target_doctype=target_doctype, target_field=erp_api_field,
        ))


# ── Sheet 7: Deliverables & Invoices ─────────────────────────────────────────

def _match_task(deliverable_name, erp_tasks):
    """Find a Task matching the deliverable name."""
    if not deliverable_name or not erp_tasks:
        return None
    d_lower = deliverable_name.lower().strip()
    for t in erp_tasks:
        subj = (t.get('subject') or '').lower().strip()
        custom_name = (t.get('custom_deliverable_name') or '').lower().strip()
        if d_lower == subj or d_lower == custom_name:
            return t
        # Partial match
        if d_lower[:20] == subj[:20] or d_lower[:20] == custom_name[:20]:
            return t
    return None


def validate_deliverables(proj_id, sheet_rows, erp_invoices, erp_tasks, erp_client=None):
    results = []
    sheet_name = '7. Deliverables & Invoices'
    TASK = 'Task'
    SI = 'Sales Invoice'

    data_rows = [r for r in sheet_rows if r.get('deliverable')]
    empty_rows = [r for r in sheet_rows if not r.get('deliverable')]
    has_tasks = bool(erp_tasks)

    if not data_rows:
        if erp_invoices:
            for inv in erp_invoices:
                for fn, ek in [('invoice', 'name'), ('amount', 'grand_total'),
                               ('status', 'status'), ('date', 'posting_date')]:
                    e_val = inv.get(ek)
                    if e_val:
                        e_d = f'{e_val:,.0f}' if isinstance(e_val, (int, float)) else str(e_val)
                        results.append(_result(
                            proj_id, sheet_name, f'{inv["name"]} — {fn}',
                            e_d, '(not in sheet)', 'MISSING_IN_SHEET',
                            target_doctype=SI,
                        ))
        elif empty_rows:
            confirmed = [r for r in empty_rows if r.get('status') == 'Confirmed']
            if confirmed:
                results.append(_result(proj_id, sheet_name, '(all)', '', '', 'MATCH',
                                       'No deliverables (confirmed)'))
            else:
                results.append(_result(proj_id, sheet_name, '(all)', '', '',
                                       'NO_STATUS', 'Empty rows not reviewed'))
        else:
            results.append(_result(proj_id, sheet_name, '(all)', '', '',
                                   'MANDATORY_EMPTY', 'No data'))
        return results

    # Build invoice map
    sheet_inv_refs = [r['invoice_num'] for r in data_rows if r.get('invoice_num')]
    inv_map, unmatched_erp = build_invoice_map(sheet_inv_refs, erp_invoices)

    # Group sheet deliverables by invoice ref
    from collections import defaultdict
    by_invoice = defaultdict(list)
    for row in data_rows:
        ref = row.get('invoice_num') or 'no_invoice'
        by_invoice[ref].append(row)

    # Cache fetched SI details
    si_detail_cache = {}
    si_items_by_invoice = {}

    # ── VALIDATION 1: Per-invoice count + sum vs SI items ──
    for inv_ref, sheet_delivs in by_invoice.items():
        if inv_ref == 'no_invoice':
            continue

        matched_inv = inv_map.get(inv_ref)
        if not matched_inv:
            results.append(_result(
                proj_id, sheet_name, f'Invoice {inv_ref}',
                '(not in ERP)', inv_ref, 'MISSING_IN_ERP',
                f'Invoice {inv_ref} not found',
                target_doctype=SI, target_field='name',
            ))
            continue

        # Fetch SI detail with items
        si_name = matched_inv['name']
        si_items = []
        if erp_client and si_name not in si_detail_cache:
            try:
                si_detail = erp_client.get_sales_invoice_detail(si_name)
                si_detail_cache[si_name] = si_detail
                si_items = si_detail.get('items', [])
                si_items_by_invoice[si_name] = si_items
            except Exception:
                si_items = []
        elif si_name in si_detail_cache:
            si_items = si_items_by_invoice.get(si_name, [])

        inv_label = f'Invoice {inv_ref} ({si_name})'

        # Count validation
        sheet_count = len(sheet_delivs)
        si_count = len(si_items)
        if sheet_count == si_count:
            results.append(_result(
                proj_id, sheet_name, f'{inv_label} — deliverable count',
                str(si_count), str(sheet_count), 'MATCH',
                target_doctype=SI, target_field='items',
            ))
        else:
            results.append(_result(
                proj_id, sheet_name, f'{inv_label} — deliverable count',
                str(si_count), str(sheet_count), 'UPDATE',
                f'Sheet has {sheet_count}, SI has {si_count}',
                target_doctype=SI, target_field='items',
            ))

        # Sum validation (sheet invoice_amount sum vs SI items amount sum)
        sheet_sum = sum(float(d.get('invoice_amount') or d.get('amount') or 0) for d in sheet_delivs)
        si_sum = sum(item.get('amount', 0) for item in si_items)
        if abs(sheet_sum - si_sum) < 1:
            results.append(_result(
                proj_id, sheet_name, f'{inv_label} — amount sum',
                f'{si_sum:,.0f}', f'{sheet_sum:,.0f}', 'MATCH',
                target_doctype=SI, target_field='items.amount',
            ))
        else:
            results.append(_result(
                proj_id, sheet_name, f'{inv_label} — amount sum',
                f'{si_sum:,.0f}', f'{sheet_sum:,.0f}', 'UPDATE',
                f'Sum mismatch: SI items={si_sum:,.0f}, Sheet={sheet_sum:,.0f}',
                target_doctype=SI, target_field='items.amount',
            ))

        # Per-deliverable vs SI item comparison
        matched_si_items = set()
        for sd in sheet_delivs:
            sd_name = (sd.get('deliverable') or '').strip()
            sd_amt = float(sd.get('invoice_amount') or sd.get('amount') or 0)
            sd_label = f'{sd_name[:35]}'

            # Match by name in SI item description (strip HTML)
            matched_item = None
            for idx, si_item in enumerate(si_items):
                if idx in matched_si_items:
                    continue
                si_desc = _strip_html(si_item.get('description', ''))
                if sd_name and si_desc and sd_name.strip()[:20] == si_desc.strip()[:20]:
                    matched_item = (idx, si_item)
                    break

            if matched_item:
                idx, si_item = matched_item
                matched_si_items.add(idx)
                si_desc = _strip_html(si_item.get('description', ''))
                si_amt = si_item.get('amount', 0)

                results.append(_result(
                    proj_id, sheet_name, f'{sd_label} — vs SI item name',
                    si_desc[:40], sd_name[:40], 'MATCH',
                    target_doctype=SI, target_field='items.description',
                ))
                if abs(sd_amt - si_amt) < 1:
                    results.append(_result(
                        proj_id, sheet_name, f'{sd_label} — vs SI item amount',
                        f'{si_amt:,.0f}', f'{sd_amt:,.0f}', 'MATCH',
                        target_doctype=SI, target_field='items.amount',
                    ))
                else:
                    results.append(_result(
                        proj_id, sheet_name, f'{sd_label} — vs SI item amount',
                        f'{si_amt:,.0f}', f'{sd_amt:,.0f}', 'UPDATE',
                        target_doctype=SI, target_field='items.amount',
                    ))
            else:
                results.append(_result(
                    proj_id, sheet_name, f'{sd_label} — vs SI item',
                    '(no matching SI item)', sd_name[:40], 'MISSING_IN_ERP',
                    'Deliverable not found in SI items',
                    target_doctype=SI, target_field='items.description',
                ))

        # SI items not matched
        for idx, si_item in enumerate(si_items):
            if idx not in matched_si_items:
                si_desc = _strip_html(si_item.get('description', ''))
                results.append(_result(
                    proj_id, sheet_name, f'{si_desc[:35]} — SI item not in sheet',
                    f'{si_item.get("amount", 0):,.0f}', '(not in sheet)', 'MISSING_IN_SHEET',
                    target_doctype=SI, target_field='items.description',
                ))

    # ── VALIDATION 2: Per-deliverable vs Task ──
    for i, row in enumerate(data_rows, 1):
        deliv_name = (row.get('deliverable') or '')[:35]
        row_label = f'Row {i} ({deliv_name})'
        row_status = row.get('status')

        # Deliverable name vs Task
        if row.get('deliverable'):
            if has_tasks:
                matched_task = _match_task(row['deliverable'], erp_tasks)
                if matched_task:
                    results.append(_result(
                        proj_id, sheet_name, f'{row_label} — vs Task name',
                        matched_task.get('subject', ''), row['deliverable'], 'MATCH',
                        target_doctype=TASK, target_field='subject',
                        target_name=matched_task['name'],
                    ))
                    # Compare amount
                    task_amt = matched_task.get('custom_deliverable_amount') or matched_task.get('custom_amount_') or 0
                    sheet_amt = float(row.get('amount') or 0)
                    if task_amt and abs(float(task_amt) - sheet_amt) > 1:
                        results.append(_result(
                            proj_id, sheet_name, f'{row_label} — vs Task amount',
                            f'{float(task_amt):,.0f}', f'{sheet_amt:,.0f}', 'UPDATE',
                            target_doctype=TASK, target_field='custom_deliverable_amount',
                            target_name=matched_task['name'],
                        ))
                    elif task_amt:
                        results.append(_result(
                            proj_id, sheet_name, f'{row_label} — vs Task amount',
                            f'{float(task_amt):,.0f}', f'{sheet_amt:,.0f}', 'MATCH',
                            target_doctype=TASK, target_field='custom_deliverable_amount',
                        ))
                else:
                    results.append(_result(
                        proj_id, sheet_name, f'{row_label} — vs Task name',
                        '(no matching Task)', row['deliverable'], 'CREATE',
                        'Task not found',
                        target_doctype=TASK, target_field='subject',
                    ))
                    if row.get('amount'):
                        results.append(_result(
                            proj_id, sheet_name, f'{row_label} — vs Task amount',
                            '(no Task)', row['amount'], 'CREATE',
                            target_doctype=TASK, target_field='custom_deliverable_amount',
                        ))
            else:
                results.append(_result(
                    proj_id, sheet_name, f'{row_label} — vs Task name',
                    '(no Tasks)', row['deliverable'], 'CREATE',
                    'No Tasks for this project',
                    target_doctype=TASK, target_field='subject',
                ))
                if row.get('amount'):
                    results.append(_result(
                        proj_id, sheet_name, f'{row_label} — vs Task amount',
                        '(no Tasks)', row['amount'], 'CREATE',
                        target_doctype=TASK, target_field='custom_deliverable_amount',
                    ))

        # Other deliverable fields
        if row.get('planned_date'):
            results.append(_result(
                proj_id, sheet_name, f'{row_label} — planned date',
                '', row['planned_date'], 'MATCH' if row_status else 'NO_STATUS',
                target_doctype=TASK, target_field='custom_estimated_invoice_date',
            ))

        if row.get('invoicing_status'):
            results.append(_result(
                proj_id, sheet_name, f'{row_label} — invoicing status',
                '', row['invoicing_status'], 'MATCH' if row_status else 'NO_STATUS',
                target_doctype=TASK, target_field='custom_invoice_payment_status',
            ))
        else:
            results.append(_result(
                proj_id, sheet_name, f'{row_label} — invoicing status',
                '', '(empty)', 'MANDATORY_EMPTY',
                target_doctype=TASK, target_field='custom_invoice_payment_status',
            ))

        if row.get('payment_date'):
            results.append(_result(
                proj_id, sheet_name, f'{row_label} — payment date',
                '', row['payment_date'], 'MATCH' if row_status else 'NO_STATUS',
                target_doctype='Payment Entry', target_field='posting_date',
            ))

        if not row_status:
            results.append(_result(proj_id, sheet_name, f'{row_label} — row status',
                                   '', '', 'NO_STATUS', 'Row not reviewed',
                                   target_doctype=TASK))

    # ERP invoices not in sheet
    for inv in unmatched_erp:
        erp_label = f'{inv["name"]}'
        for fn, ek, tf in [('invoice', 'name', 'name'), ('amount', 'grand_total', 'grand_total'),
                            ('status', 'status', 'status'), ('date', 'posting_date', 'posting_date')]:
            e_val = inv.get(ek)
            if e_val:
                e_d = f'{e_val:,.0f}' if isinstance(e_val, (int, float)) else str(e_val)
                results.append(_result(
                    proj_id, sheet_name, f'{erp_label} — {fn}',
                    e_d, '(not in sheet)', 'MISSING_IN_SHEET',
                    'ERP invoice not referenced in sheet',
                    target_doctype=SI, target_field=tf,
                ))

    return results
