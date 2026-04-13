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

    # ── Expenses from GL Entry ──────────────────────────────────

    def get_gl_expenses(self, proj_id):
        """Get project expenses from General Ledger, grouped by account.
        GL Entry is the single source of truth — aggregates from all source doctypes.
        Returns list of dicts with account, net amount, voucher details.
        """
        entries = self._get_list(
            'GL Entry',
            filters=[['project', '=', proj_id]],
            fields=['account', 'debit', 'credit', 'voucher_type', 'voucher_no', 'posting_date'],
            limit=500,
        )

        # Group by account
        from collections import defaultdict
        by_account = defaultdict(lambda: {'debit': 0, 'credit': 0, 'vouchers': []})
        for e in entries:
            acc = e['account']
            by_account[acc]['debit'] += e.get('debit', 0)
            by_account[acc]['credit'] += e.get('credit', 0)
            by_account[acc]['vouchers'].append({
                'voucher_type': e.get('voucher_type', ''),
                'voucher_no': e.get('voucher_no', ''),
                'posting_date': e.get('posting_date', ''),
                'debit': e.get('debit', 0),
                'credit': e.get('credit', 0),
            })

        # Return expense accounts (5xxx) with net amounts
        results = []
        for acc, vals in sorted(by_account.items()):
            net = vals['debit'] - vals['credit']
            results.append({
                'account': acc,
                'net_amount': net,
                'debit': vals['debit'],
                'credit': vals['credit'],
                'vouchers': vals['vouchers'],
            })
        return results

    # ── Customer lookup ──────────────────────────────────────────

    def find_customer(self, name_query):
        """Find closest Customer record by name. Returns (customer_id, customer_name) or None."""
        customers = self._get_list(
            'Customer',
            filters=[['customer_name', 'like', f'%{name_query}%']],
            fields=['name', 'customer_name'],
            limit=5,
        )
        if customers:
            return customers[0]['name'], customers[0]['customer_name']
        return None, None

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

    # ── Sales Invoice detail (with items child table) ──────────

    def get_sales_invoice_detail(self, sinv_name):
        """Fetch full Sales Invoice including items child table."""
        return self._get(f'/api/resource/Sales Invoice/{sinv_name}')['data']

    # ── PPI (read-only for expenses) ─────────────────────────────

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
