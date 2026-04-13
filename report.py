"""Generate validation reports — terminal output and CSV."""

import csv
import os
from collections import defaultdict


STATUS_SYMBOLS = {
    'MATCH': 'MATCH',
    'UPDATE': 'UPDATE',
    'CREATE': 'CREATE',
    'MISSING_IN_ERP': 'MISSING_IN_ERP',
    'MISSING_IN_SHEET': 'MISSING_IN_SHEET',
    'NO_STATUS': 'NO_STATUS',
    'MANDATORY_EMPTY': 'MANDATORY_EMPTY',
    'BLOCKED': 'BLOCKED',
    'RELEASE_RESOURCE': 'RELEASE_RESOURCE',
}

SHEET_ORDER = [
    '1. Project Info',
    '3. Resources (Planned)',
    '4. Resources (Actual)',
    '5. Third-Party (Planned)',
    '6. Third-Party (Actual)',
    '7. Deliverables & Invoices',
]


def print_terminal_report(results, totals_reconciliation=None):
    """Print validation report grouped by project → sheet → field."""
    # Group by project
    by_project = defaultdict(list)
    for r in results:
        by_project[r['project']].append(r)

    for proj_id in sorted(by_project.keys()):
        proj_results = by_project[proj_id]

        # Summary counts
        counts = defaultdict(int)
        for r in proj_results:
            counts[r['status']] += 1

        # Count ERP actions
        action_counts = defaultdict(int)
        for r in proj_results:
            action_counts[r.get('erp_action', 'SKIP')] += 1

        summary_parts = []
        for s in ['UPDATE', 'CREATE', 'TO_DELETE', 'MATCH', 'MANDATORY_EMPTY', 'MISSING_IN_ERP',
                   'MISSING_IN_SHEET', 'NO_STATUS', 'BLOCKED', 'RELEASE_RESOURCE']:
            if counts[s]:
                summary_parts.append(f'{counts[s]} {s}')

        print(f'\n{"=" * 70}')
        print(f'{proj_id}')
        print(f'{"=" * 70}')

        # Group by sheet
        by_sheet = defaultdict(list)
        for r in proj_results:
            by_sheet[r['sheet']].append(r)

        for sheet in SHEET_ORDER:
            if sheet not in by_sheet:
                continue
            sheet_results = by_sheet[sheet]

            print(f'\n  {sheet}')
            print(f'  {"─" * 66}')

            if sheet == '1. Project Info':
                _print_project_info_table(sheet_results)
            elif sheet in ('3. Resources (Planned)', '5. Third-Party (Planned)'):
                _print_planned_table(sheet_results)
            elif sheet == '4. Resources (Actual)':
                _print_ped_table(sheet_results)
            elif sheet == '6. Third-Party (Actual)':
                _print_thirdparty_actual_table(sheet_results)
            elif sheet == '7. Deliverables & Invoices':
                _print_deliverables_table(sheet_results)

        # Totals reconciliation
        if totals_reconciliation and proj_id in totals_reconciliation:
            _print_totals_reconciliation(totals_reconciliation[proj_id])

        # Project summary
        action_summary = ' | '.join(f'{v} {k}' for k, v in sorted(action_counts.items()) if v)
        print(f'\n  Summary: {" | ".join(summary_parts)}')
        print(f'  ERP Actions: {action_summary}')


def _print_project_info_table(results):
    print(f'  {"Field":<22} {"ERP Value":<22} {"Sheet Value":<22} {"Status":<20} {"Action":<8} {"Module":<10} {"Target Field":<25}')
    print(f'  {"─" * 22} {"─" * 22} {"─" * 22} {"─" * 20} {"─" * 8} {"─" * 10} {"─" * 25}')
    for r in results:
        erp = (r['erp_value'] or '(empty)')[:20]
        sheet = (r['sheet_value'] or '(empty)')[:20]
        status = r['status'][:18]
        action = r.get('erp_action', 'SKIP')
        module = (r.get('erp_module') or '')[:10]
        target = (r.get('erp_target_field') or '')[:25]
        print(f'  {r["field"]:<22} {erp:<22} {sheet:<22} {status:<20} {action:<8} {module:<10} {target}')


def _print_detail_row(r):
    """Print a single detail row with ERP, Sheet, Status, Action, Module, Target."""
    erp = (r['erp_value'] or '(empty)')[:25]
    sheet = (r['sheet_value'] or '(empty)')[:25]
    status = r['status']
    action = r.get('erp_action', 'SKIP')
    module = r.get('erp_module') or ''
    target = r.get('erp_target_field') or ''
    msg = f' ({r["message"][:20]})' if r['message'] else ''
    print(f'  {r["field"]:<45} {erp:<25} → {sheet:<25} {status}{msg}')
    print(f'  {"":>45} [{action}] {module}{"." + target if target else ""}')


def _print_planned_table(results):
    for r in results:
        _print_detail_row(r)


def _print_ped_table(results):
    for r in results:
        _print_detail_row(r)


def _print_thirdparty_actual_table(results):
    for r in results:
        _print_detail_row(r)


def _print_deliverables_table(results):
    for r in results:
        _print_detail_row(r)


def _print_totals_reconciliation(totals):
    print(f'\n  Totals Reconciliation (Sheet 2 vs Computed)')
    print(f'  {"─" * 66}')
    print(f'  {"Field":<35} {"Sheet 2":>12} {"Computed":>12} {"ERP":>12} {"Discrepancy":<15}')
    print(f'  {"─" * 35} {"─" * 12} {"─" * 12} {"─" * 12} {"─" * 15}')
    for t in totals:
        s2 = t.get('sheet2', '—')
        comp = t.get('computed', '—')
        erp = t.get('erp', '—')
        disc = t.get('discrepancy', '—')
        print(f'  {t["field"]:<35} {str(s2):>12} {str(comp):>12} {str(erp):>12} {disc:<15}')


def write_csv_report(results, output_path, totals_reconciliation=None):
    """Write validation results to CSV."""
    os.makedirs(os.path.dirname(output_path) or '.', exist_ok=True)
    with open(output_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=[
            'project', 'sheet', 'field', 'erp_value', 'sheet_value',
            'status', 'erp_action', 'erp_module', 'erp_target_field',
            'message', 'target_name', 'child_name',
        ])
        writer.writeheader()
        for r in results:
            writer.writerow({k: r.get(k, '') for k in writer.fieldnames})

    print(f'\nCSV report written to: {output_path}')
