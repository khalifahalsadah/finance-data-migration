"""Finance Data Migration Tool — validate and sync Excel corrections to ERPNext."""

import argparse
import sys
import os
from collections import defaultdict

from config import TOTALS_FIELDS
from excel_parser import parse_workbook
from erp_client import ERPClient
from validators import (
    validate_project_info,
    validate_resources_planned,
    validate_resources_actual,
    validate_thirdparty_planned,
    validate_thirdparty_actual,
    validate_deliverables,
)
from report import print_terminal_report, write_csv_report
from backup import create_backup, restore_backup
from sharepoint import download_workbook


def get_workbook_path(args):
    if hasattr(args, 'file') and args.file:
        return args.file
    default = os.path.join(os.path.dirname(__file__), 'workbook.xlsx')
    if os.path.exists(default):
        return default
    print('No workbook found. Run `python migrate.py download` first or use --file.')
    sys.exit(1)


def get_projects_for_batch(wb_data, batch_num, project_filter=None):
    batch = wb_data['batches'].get(batch_num, [])
    if not batch:
        print(f'Batch {batch_num} not found. Available: {sorted(wb_data["batches"].keys())}')
        sys.exit(1)
    projects = [p['id'] for p in batch]
    if project_filter:
        if project_filter not in projects:
            print(f'{project_filter} is not in Batch {batch_num}. Projects: {projects}')
            sys.exit(1)
        projects = [project_filter]
    return projects


def run_validation(wb_data, projects, erp):
    """Run all validators for the given projects. Returns list of results."""
    all_results = []

    for proj_id in projects:
        print(f'\nValidating {proj_id}...', flush=True)

        # Fetch ERP data
        try:
            erp_project = erp.get_project(proj_id)
        except Exception as e:
            all_results.append({
                'project': proj_id, 'sheet': '(all)', 'field': 'Project',
                'erp_value': '', 'sheet_value': '', 'status': 'BLOCKED',
                'message': f'Cannot fetch project: {e}',
                'target_doctype': '', 'target_name': '', 'target_field': '',
                'child_name': '', 'row_num': '',
            })
            continue

        try:
            ped_data = erp.get_ped_for_project(proj_id)
        except Exception:
            ped_data = []

        # Get quotation ref from sheet if available
        sheet_quot_ref = None
        planned_rows = wb_data['resources_planned'].get(proj_id, [])
        if planned_rows:
            sheet_quot_ref = planned_rows[0].get('quotation')
        if not sheet_quot_ref:
            tp_rows = wb_data['thirdparty_planned'].get(proj_id, [])
            if tp_rows:
                sheet_quot_ref = tp_rows[0].get('quotation')

        try:
            quotation, quot_note = erp.get_quotation_for_project(proj_id, sheet_quot_ref)
        except Exception:
            quotation, quot_note = None, ''

        try:
            invoices = erp.get_sales_invoices(proj_id)
        except Exception:
            invoices = []

        try:
            tasks = erp.get_tasks(proj_id)
        except Exception:
            tasks = []

        # Run validators
        all_results += validate_project_info(
            proj_id,
            wb_data['project_info'].get(proj_id, {}),
            erp_project,
            erp,
        )
        all_results += validate_resources_planned(
            proj_id,
            wb_data['resources_planned'].get(proj_id, []),
            quotation,
            quot_note,
        )
        all_results += validate_resources_actual(
            proj_id,
            wb_data['resources_actual'].get(proj_id, []),
            ped_data,
        )
        all_results += validate_thirdparty_planned(
            proj_id,
            wb_data['thirdparty_planned'].get(proj_id, []),
            quotation,
            quot_note,
        )
        try:
            gl_expenses = erp.get_gl_expenses(proj_id)
        except Exception:
            gl_expenses = []

        all_results += validate_thirdparty_actual(
            proj_id,
            wb_data['thirdparty_actual'].get(proj_id, []),
            gl_expenses,
        )
        all_results += validate_deliverables(
            proj_id,
            wb_data['deliverables'].get(proj_id, []),
            invoices,
            tasks,
            erp,
        )

    return all_results


def compute_totals_reconciliation(wb_data, projects):
    """Compare Sheet 2 totals against computed values from raw data."""
    reconciliation = {}

    for proj_id in projects:
        sheet2 = wb_data['totals'].get(proj_id, {})
        planned_res = wb_data['resources_planned'].get(proj_id, [])
        actual_res = wb_data['resources_actual'].get(proj_id, [])
        planned_3p = wb_data['thirdparty_planned'].get(proj_id, [])
        actual_3p = wb_data['thirdparty_actual'].get(proj_id, [])
        deliverables = wb_data['deliverables'].get(proj_id, [])

        # Compute planned cost from resources + 3P
        planned_res_cost = 0
        for r in planned_res:
            try:
                rate = float(r.get('rate', 0) or 0)
                qty = float(r.get('qty', 0) or 0)
                ratio = float(r.get('ratio', 0) or 0)
                planned_res_cost += rate * qty * (ratio / 100)
            except (ValueError, TypeError):
                pass

        planned_3p_cost = 0
        for r in planned_3p:
            if r.get('entry_type') == 'Sales Tax & Charges':
                continue
            try:
                planned_3p_cost += float(r.get('total', 0) or 0)
            except (ValueError, TypeError):
                pass

        # Compute actual cost from 3P actuals (resource cost comes from PED buying_price, not in sheet)
        actual_3p_cost = 0
        for r in actual_3p:
            try:
                actual_3p_cost += float(r.get('amount', 0) or 0)
            except (ValueError, TypeError):
                pass

        # Revenue from deliverables
        planned_revenue = 0
        actual_revenue = 0
        for r in deliverables:
            try:
                planned_revenue += float(r.get('amount', 0) or 0)
            except (ValueError, TypeError):
                pass
            try:
                actual_revenue += float(r.get('invoice_amount', 0) or 0)
            except (ValueError, TypeError):
                pass

        def _s2(label):
            f = sheet2.get(label, {})
            return f.get('effective_val')

        def _fmt(v):
            if v is None:
                return '—'
            try:
                return f'{float(v):,.0f}'
            except (ValueError, TypeError):
                return str(v)

        totals = []
        computed_planned_cost = planned_res_cost + planned_3p_cost

        totals.append({
            'field': 'Planned Cost (resources + 3P)',
            'sheet2': _fmt(_s2('Planned Cost (till date)')),
            'computed': _fmt(computed_planned_cost) if computed_planned_cost else '—',
            'erp': '—',
            'discrepancy': '',
        })

        totals.append({
            'field': 'Actual 3P Cost',
            'sheet2': '—',
            'computed': _fmt(actual_3p_cost) if actual_3p_cost else '—',
            'erp': '—',
            'discrepancy': '',
        })

        totals.append({
            'field': 'Planned Revenue (deliverables)',
            'sheet2': _fmt(_s2('Planned Revenue (till completion)')),
            'computed': _fmt(planned_revenue) if planned_revenue else '—',
            'erp': '—',
            'discrepancy': '',
        })

        totals.append({
            'field': 'Actual Revenue (invoiced)',
            'sheet2': _fmt(_s2('Actual Revenue')),
            'computed': _fmt(actual_revenue) if actual_revenue else '—',
            'erp': '—',
            'discrepancy': '',
        })

        totals.append({
            'field': 'Planned Duration (months)',
            'sheet2': _fmt(_s2('Planned Duration (months)')),
            'computed': '—',
            'erp': '—',
            'discrepancy': '',
        })

        totals.append({
            'field': 'Actual Duration (months)',
            'sheet2': _fmt(_s2('Actual Duration (months)')),
            'computed': '—',
            'erp': '—',
            'discrepancy': '',
        })

        reconciliation[proj_id] = totals

    return reconciliation


# ── CLI Commands ─────────────────────────────────────────────────────────────

def cmd_download(args):
    output = os.path.join(os.path.dirname(__file__), 'workbook.xlsx')
    download_workbook(output)


def cmd_validate(args):
    wb_path = get_workbook_path(args)
    wb_data = parse_workbook(wb_path)
    projects = get_projects_for_batch(wb_data, args.batch, args.project)

    print(f'Validating Batch {args.batch}: {len(projects)} projects')
    print(f'Workbook: {wb_path}')

    erp = ERPClient()
    results = run_validation(wb_data, projects, erp)
    totals_recon = compute_totals_reconciliation(wb_data, projects)

    print_terminal_report(results, totals_recon)

    if args.csv:
        csv_path = os.path.join(os.path.dirname(__file__), 'reports',
                                f'batch{args.batch}_validation.csv')
        write_csv_report(results, csv_path)

    # Print overall summary
    counts = defaultdict(int)
    action_counts = defaultdict(int)
    for r in results:
        counts[r['status']] += 1
        action_counts[r.get('erp_action', 'SKIP')] += 1
    print(f'\n{"=" * 70}')
    print(f'OVERALL: {len(results)} checks across {len(projects)} projects')
    print(f'\nBy Status:')
    for s in ['UPDATE', 'CREATE', 'TO_DELETE', 'MATCH', 'MANDATORY_EMPTY', 'MISSING_IN_ERP',
              'MISSING_IN_SHEET', 'NO_STATUS', 'BLOCKED', 'RELEASE_RESOURCE']:
        if counts[s]:
            print(f'  {s}: {counts[s]}')
    print(f'\nERP Actions if applied:')
    for a in ['UPDATE', 'CREATE', 'TO_DELETE', 'SKIP']:
        if action_counts[a]:
            print(f'  {a}: {action_counts[a]}')


def cmd_apply(args):
    wb_path = get_workbook_path(args)
    wb_data = parse_workbook(wb_path)
    projects = get_projects_for_batch(wb_data, args.batch, args.project)

    erp = ERPClient()
    results = run_validation(wb_data, projects, erp)

    updates = [r for r in results if r['status'] == 'UPDATE'
               and r['target_doctype'] and r['target_field']]

    if not updates:
        print('No updates to apply.')
        return

    if not args.execute:
        print(f'[DRY RUN] Would apply {len(updates)} changes:\n')
        for u in updates:
            print(f'  {u["project"]} | {u["sheet"]} | {u["field"]}')
            print(f'    {u["target_doctype"]}.{u["target_name"]}.{u["target_field"]}')
            print(f'    {u["erp_value"]} → {u["sheet_value"]}')
        print(f'\nTo apply, add --execute')
        return

    # Create backup first
    backup_path = create_backup(args.batch, updates, erp,
                                os.path.join(os.path.dirname(__file__), 'backups'))

    success = 0
    errors = 0
    for u in updates:
        try:
            if u['target_doctype'] == 'Project':
                erp.update_project(u['target_name'], {u['target_field']: u['sheet_value']})
            elif u['target_doctype'] == 'Project Employee Distribution':
                erp.update_ped_row(u['target_name'], u['child_name'],
                                   {u['target_field']: u['sheet_value']})
            elif u['target_doctype'] == 'Quotation':
                erp.update_quotation(u['target_name'], {u['target_field']: u['sheet_value']})
            else:
                print(f'  SKIP: Unknown doctype {u["target_doctype"]}')
                continue
            print(f'  OK: {u["project"]} {u["field"]}: {u["erp_value"]} → {u["sheet_value"]}')
            success += 1
        except Exception as e:
            print(f'  ERROR: {u["project"]} {u["field"]}: {e}')
            errors += 1

    print(f'\nApplied: {success}, Errors: {errors}')
    print(f'Backup: {backup_path} (use `python migrate.py restore {backup_path}` to revert)')


def cmd_restore(args):
    erp = ERPClient()
    restore_backup(args.backup, erp, args.project)


# ── Main ─────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description='Finance Data Migration Tool')
    sub = parser.add_subparsers(dest='command')

    # download
    p_dl = sub.add_parser('download', help='Download latest workbook from SharePoint')

    # validate
    p_val = sub.add_parser('validate', help='Validate batch against ERP')
    p_val.add_argument('--batch', type=int, required=True, help='Batch number (1-7)')
    p_val.add_argument('--project', type=str, help='Single project ID (e.g., PROJ-1214)')
    p_val.add_argument('--file', type=str, help='Path to local Excel file')
    p_val.add_argument('--csv', action='store_true', help='Also write CSV report')

    # apply
    p_app = sub.add_parser('apply', help='Apply corrections to ERP')
    p_app.add_argument('--batch', type=int, required=True, help='Batch number')
    p_app.add_argument('--project', type=str, help='Single project ID')
    p_app.add_argument('--file', type=str, help='Path to local Excel file')
    p_app.add_argument('--execute', action='store_true', help='Actually apply (default is dry-run)')

    # restore
    p_res = sub.add_parser('restore', help='Restore from backup')
    p_res.add_argument('backup', type=str, help='Path to backup JSON file')
    p_res.add_argument('--project', type=str, help='Restore only one project')

    args = parser.parse_args()

    if args.command == 'download':
        cmd_download(args)
    elif args.command == 'validate':
        cmd_validate(args)
    elif args.command == 'apply':
        cmd_apply(args)
    elif args.command == 'restore':
        cmd_restore(args)
    else:
        parser.print_help()


if __name__ == '__main__':
    main()
