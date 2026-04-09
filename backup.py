"""Backup and restore ERP data for reversibility."""

import json
import os
from datetime import datetime


def create_backup(batch_num, updates, erp_client, backup_dir='backups'):
    """Snapshot current ERP values for all fields being changed.

    Args:
        batch_num: Batch number for naming
        updates: List of validation results with status='UPDATE'
        erp_client: ERPClient instance to fetch current values
        backup_dir: Directory to store backups

    Returns: Backup file path
    """
    os.makedirs(backup_dir, exist_ok=True)
    ts = datetime.now().strftime('%Y-%m-%d_%H%M%S')
    backup_path = os.path.join(backup_dir, f'{ts}_batch{batch_num}_pre.json')

    changes = []
    for u in updates:
        change = {
            'doctype': u['target_doctype'],
            'name': u['target_name'],
            'field': u['target_field'],
            'old_value': u['erp_value'],
            'new_value': u['sheet_value'],
            'project': u['project'],
            'sheet': u['sheet'],
        }
        if u.get('child_name'):
            change['child_name'] = u['child_name']
        changes.append(change)

    backup = {
        'created': datetime.now().isoformat(),
        'batch': batch_num,
        'total_changes': len(changes),
        'changes': changes,
    }

    with open(backup_path, 'w', encoding='utf-8') as f:
        json.dump(backup, f, indent=2, ensure_ascii=False)

    print(f'Backup saved: {backup_path} ({len(changes)} changes)')
    return backup_path


def restore_backup(backup_path, erp_client, project_filter=None):
    """Restore ERP to pre-change state from backup.

    Args:
        backup_path: Path to backup JSON file
        erp_client: ERPClient instance
        project_filter: Optional project ID to restore only one project

    Returns: Dict with success/error counts
    """
    with open(backup_path, 'r', encoding='utf-8') as f:
        backup = json.load(f)

    changes = backup['changes']
    if project_filter:
        changes = [c for c in changes if c['project'] == project_filter]

    print(f'Restoring {len(changes)} changes from {backup_path}')
    if project_filter:
        print(f'  Filtered to project: {project_filter}')

    success = 0
    errors = 0

    for c in changes:
        try:
            if c['doctype'] == 'Project':
                erp_client.update_project(c['name'], {c['field']: c['old_value']})
                print(f'  Restored {c["name"]}.{c["field"]} → {c["old_value"]}')
                success += 1
            elif c['doctype'] == 'Project Employee Distribution':
                erp_client.update_ped_row(c['name'], c['child_name'], {c['field']: c['old_value']})
                print(f'  Restored PED {c["name"]}.{c["child_name"]}.{c["field"]} → {c["old_value"]}')
                success += 1
            elif c['doctype'] == 'Quotation':
                erp_client.update_quotation(c['name'], {c['field']: c['old_value']})
                print(f'  Restored {c["name"]}.{c["field"]} → {c["old_value"]}')
                success += 1
            else:
                print(f'  SKIP: Unknown doctype {c["doctype"]}')
                errors += 1
        except Exception as e:
            print(f'  ERROR restoring {c["name"]}.{c["field"]}: {e}')
            errors += 1

    print(f'\nRestore complete. Success: {success}, Errors: {errors}')
    return {'success': success, 'errors': errors}
