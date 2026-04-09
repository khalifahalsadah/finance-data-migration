"""ERPNext API client for reading and writing project data."""

import requests
import json
from config import ERP_BASE, ERP_TOKEN


class ERPClient:
    def __init__(self, base=ERP_BASE, token=ERP_TOKEN):
        self.base = base
        self.headers = {
            'Authorization': f'token {token}',
            'Content-Type': 'application/json',
        }

    def _get(self, endpoint, params=None):
        r = requests.get(f'{self.base}{endpoint}', headers=self.headers, params=params)
        r.raise_for_status()
        return r.json()

    def _put(self, endpoint, data):
        r = requests.put(f'{self.base}{endpoint}', headers=self.headers, json=data)
        r.raise_for_status()
        return r.json()

    def _get_list(self, doctype, filters=None, fields=None, limit=100):
        params = {
            'doctype': doctype,
            'limit_page_length': limit,
        }
        if filters:
            params['filters'] = json.dumps(filters)
        if fields:
            params['fields'] = json.dumps(fields)
        r = self._get('/api/method/frappe.client.get_list', params=params)
        return r.get('message', [])

    # ── Project ──────────────────────────────────────────────────

    def get_project(self, proj_id):
        return self._get(f'/api/resource/Project/{proj_id}')['data']

    # ── PED (Project Employee Distribution) ──────────────────────

    def get_ped_for_project(self, proj_id):
        """Get all PED distribution_detail rows for a project.
        Queries PED directly by project (not via PPI).
        Returns: list of detail rows with _ped_id added.
        """
        # Query PED directly by project
        ped_list = self._get_list(
            'Project Employee Distribution',
            filters=[['project', '=', proj_id]],
            fields=['name'],
            limit=20,
        )

        results = []
        for ped_rec in ped_list:
            ped_id = ped_rec['name']
            try:
                ped = self._get(f'/api/resource/Project Employee Distribution/{ped_id}')['data']
                for detail in ped.get('distribution_detail', []):
                    detail['_ped_id'] = ped_id
                    results.append(detail)
            except Exception:
                pass
        return results

    def get_ped_record(self, ped_id):
        return self._get(f'/api/resource/Project Employee Distribution/{ped_id}')['data']

    # ── Quotation ────────────────────────────────────────────────

    def get_quotation(self, qtn_name):
        return self._get(f'/api/resource/Quotation/{qtn_name}')['data']

    def get_quotation_for_project(self, proj_id, sheet_quot_ref=None):
        """Find the Quotation linked to a project.
        Tries: 1) Opportunity link, 2) Direct sheet reference.
        Returns: (quotation_data, lookup_note) tuple.
        lookup_note is '' if found via opportunity, or a warning message.
        """
        project = self.get_project(proj_id)
        opp = project.get('opportunity')

        # Try via Opportunity link
        if opp:
            quotations = self._get_list(
                'Quotation',
                filters=[['opportunity', '=', opp]],
                fields=['name', 'docstatus', 'grand_total'],
                limit=10,
            )
            if quotations:
                submitted = [q for q in quotations if q['docstatus'] == 1]
                best = submitted[0] if submitted else quotations[0]
                return self.get_quotation(best['name']), ''

        # Fallback: direct lookup by sheet reference
        if sheet_quot_ref:
            try:
                quot = self.get_quotation(sheet_quot_ref)
                # Check if this quotation is linked to a different project
                quot_opp = quot.get('opportunity', '')
                if quot_opp and quot_opp != opp:
                    # Check if the opportunity's project is different
                    try:
                        opp_data = self._get(f'/api/resource/Opportunity/{quot_opp}')['data']
                        opp_project = opp_data.get('external_project_reference', '')
                        if opp_project and opp_project != proj_id:
                            return quot, f'WRONG_PROJECT: Quotation linked to {opp_project} via opportunity {quot_opp}'
                    except Exception:
                        pass
                return quot, f'NOT_LINKED: Quotation not linked to opportunity {opp or "(none)"}'
            except Exception:
                pass

        return None, ''

    # ── Sales Invoice ────────────────────────────────────────────

    def get_sales_invoices(self, proj_id):
        return self._get_list(
            'Sales Invoice',
            filters=[['project', '=', proj_id], ['docstatus', '=', 1]],
            fields=['name', 'grand_total', 'outstanding_amount', 'status', 'posting_date'],
            limit=100,
        )

    # ── Expenses (combined from source doctypes + PPI) ─────────

    def get_expenses_for_project(self, proj_id):
        """Get all expenses from source doctypes + PPI, merged by reference.
        Returns dict keyed by reference form name.
        """
        expenses = {}

        # 1. PPI project_expenses (has the most detail including child doc refs)
        try:
            ppi = self.get_ppi(proj_id)
            for e in ppi.get('project_expenses', []):
                ref = e.get('reference_form', '')
                if ref:
                    expenses[ref] = {
                        'reference_form': ref,
                        'reference_doctype': e.get('reference_doctype', ''),
                        'total': e.get('total', 0),
                        'form_date': e.get('form_date', ''),
                        'child_table_value': e.get('child_table_value', ''),
                        'source': 'PPI',
                    }
        except Exception:
            pass

        # 2. Expense Claims (submitted)
        try:
            ecs = self._get_list(
                'Expense Claim',
                filters=[['project', '=', proj_id], ['docstatus', '=', 1]],
                fields=['name', 'total_claimed_amount', 'posting_date', 'employee_name'],
                limit=100,
            )
            for ec in ecs:
                ref = ec['name']
                if ref not in expenses:
                    expenses[ref] = {
                        'reference_form': ref,
                        'reference_doctype': 'Expense Claim',
                        'total': ec.get('total_claimed_amount', 0),
                        'form_date': ec.get('posting_date', ''),
                        'child_table_value': ec.get('employee_name', ''),
                        'source': 'Expense Claim',
                    }
        except Exception:
            pass

        # 3. Purchase Invoices (submitted, linked to project)
        try:
            pis = self._get_list(
                'Purchase Invoice',
                filters=[['project', '=', proj_id], ['docstatus', '=', 1]],
                fields=['name', 'grand_total', 'posting_date', 'supplier_name'],
                limit=100,
            )
            for pi in pis:
                ref = pi['name']
                if ref not in expenses:
                    expenses[ref] = {
                        'reference_form': ref,
                        'reference_doctype': 'Purchase Invoice',
                        'total': pi.get('grand_total', 0),
                        'form_date': pi.get('posting_date', ''),
                        'child_table_value': pi.get('supplier_name', ''),
                        'source': 'Purchase Invoice',
                    }
        except Exception:
            pass

        # 4. Journal Entries (via Journal Entry Account child table)
        try:
            je_accounts = self._get_list(
                'Journal Entry Account',
                filters=[['project', '=', proj_id]],
                fields=['parent', 'debit_in_account_currency', 'account'],
                limit=100,
            )
            for ja in je_accounts:
                ref = ja['parent']
                if ref not in expenses and ja.get('debit_in_account_currency', 0) > 0:
                    expenses[ref] = {
                        'reference_form': ref,
                        'reference_doctype': 'Journal Entry',
                        'total': ja.get('debit_in_account_currency', 0),
                        'form_date': '',
                        'child_table_value': ja.get('account', ''),
                        'source': 'Journal Entry',
                    }
        except Exception:
            pass

        return expenses

    # ── Task (deliverables) ─────────────────────────────────────

    def get_tasks(self, proj_id):
        return self._get_list(
            'Task',
            filters=[['project', '=', proj_id]],
            fields=['name', 'subject', 'status', 'custom_deliverable_name',
                    'custom_deliverable_amount', 'custom_amount_',
                    'custom_estimated_invoice_date', 'custom_actual_invoice_date',
                    'custom_invoice_no', 'custom_invoice_payment_status',
                    'custom_actual_payment_date', 'custom_collection_status'],
            limit=50,
        )

    # ── PPI (read-only for expenses and deliverables) ────────────

    def get_ppi(self, proj_id):
        return self._get(f'/api/resource/Project Profitably Input/{proj_id}')['data']

    # ── Write operations ─────────────────────────────────────────

    def update_project(self, proj_id, fields):
        return self._put(f'/api/resource/Project/{proj_id}', fields)

    def update_ped_row(self, ped_id, child_name, fields):
        """Update a specific child row in PED distribution_detail."""
        ped = self.get_ped_record(ped_id)
        details = ped.get('distribution_detail', [])
        for d in details:
            if d['name'] == child_name:
                d.update(fields)
                break
        return self._put(
            f'/api/resource/Project Employee Distribution/{ped_id}',
            {'distribution_detail': details},
        )

    def update_quotation(self, qtn_name, fields):
        return self._put(f'/api/resource/Quotation/{qtn_name}', fields)
