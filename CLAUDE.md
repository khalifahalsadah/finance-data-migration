# Finance Data Migration Tool

## Purpose
Validates and syncs project financial data between an Excel workbook (maintained by Finance team) and ERPNext. Compares every field in every sheet against live ERP data, generates a detailed report, and can apply corrections with full backup/restore capability.

## Current Status (2026-04-14)
- **Environment**: UAT (`https://strategicgears-uat.frappe.cloud`)
- **Batch 1**: Validated and approved by Yazeed (v4). Ready for apply on UAT.
- **Batch 2**: Validated, not yet reviewed.
- **Apply logic**: Project fields (Sheet 1) and PED (Sheet 4) working. Task creation (Sheet 7), Quotation updates (Sheet 3/5), and tax handling still need implementation.

## Source Data
- **Excel workbook**: `Project_Data_Completion_Templates_v3.xlsx` on Yazeed Alsughayyir's OneDrive
- **File ID**: `8B8607E9-7CA6-4384-957F-29E09FCFBE43`
- **User**: `yazeed.su@strategicgears.com`
- **43 projects** across 7 batches, 7 data sheets

## Commands
```bash
cd /Users/khalifah/Projects/sg/finance-data-migration

# Download latest workbook from SharePoint
/Users/khalifah/Projects/sg/venv/bin/python3 migrate.py download

# Validate a batch
/Users/khalifah/Projects/sg/venv/bin/python3 migrate.py validate --batch 1
/Users/khalifah/Projects/sg/venv/bin/python3 migrate.py validate --batch 1 --project PROJ-1214
/Users/khalifah/Projects/sg/venv/bin/python3 migrate.py validate --batch 1 --csv

# Dry run apply (shows what would change)
/Users/khalifah/Projects/sg/venv/bin/python3 migrate.py apply --batch 1

# Apply (Claude must NEVER run this — provide command to user)
/Users/khalifah/Projects/sg/venv/bin/python3 migrate.py apply --batch 1 --execute

# Restore from backup
/Users/khalifah/Projects/sg/venv/bin/python3 migrate.py restore backups/<file>.json
```

## CRITICAL RULE
**Claude must NEVER execute ERP write commands.** When asked to apply, respond with:
"Here is the command: `python migrate.py apply --batch X --execute` — as per instructions I'm not allowed to execute this."

## Sheet → ERP Module Mapping

| Sheet | ERP Module | Target Fields |
|---|---|---|
| 1. Project Info | `Project` | `workflow_state`, `custom_billability_type`, `custom_od_reference_`, `project_name`, `customer_name`, `custom_awarded_date`, `custom_official_start_date`, `official_end_date`, `custom_partner_name_value`, `custom_engagement_manager`, `custom_end_client_name`, `custom_contractuals_link` |
| 2. Totals | (derived — never update directly) | Reconciliation check only |
| 3. Resources (Planned) | `Quotation` → `custom_sg_resources_price` child table | `resources_level`, `percent`, `uom`, `qty`, `rate` |
| 4. Resources (Actual) | `Project Employee Distribution` (PED) → `distribution_detail` child table | `employee_name`, `designation`, `from_date`, `to_date`, `ratio_` (note underscore) |
| 5. Third-Party (Planned) | `Quotation` → `custom_other`, `custom_tpc__osa`, `custom_travel_cost__egypt`, `custom_travel_cost__local`, `taxes` | `team_name`, `rate`, `months`, `total` |
| 6. Third-Party (Actual) | Resolved per row: `Purchase Invoice` (ACC-PINV-*), `Expense Claim` (HR-EXP-*), `Journal Entry` (ACC-JV-*) | `total`, `posting_date`, `description` |
| 7. Deliverables & Invoices | `Task` (deliverable milestones) + `Sales Invoice` (invoice data) + `Payment Entry` (payment dates) | Task: `subject`, `custom_deliverable_amount`, `custom_estimated_invoice_date`, `custom_invoice_payment_status`, `custom_actual_invoice_date`, `custom_invoice_no`. SI: `grand_total`, `status`, `posting_date` |

## PED API Note
- **API doctype name**: `Project Employee Distribution` (NOT `CRM Project Employee Distribution` — that one returns 500)
- **Ratio field**: `ratio_` (with underscore)
- PED IDs are found via PPI → `project_resources_allocation` → `ped` field

## Validation Logic

Every field is checked using this logic:

| Sheet Column | Status Column | Action |
|---|---|---|
| Corrected col has value | Corrected/Missing in ERP | Compare corrected value vs live ERP → UPDATE if different, MATCH if same |
| Corrected col empty | Confirmed | Compare sheet ERP col vs live ERP → MATCH if same, UPDATE if different |
| Corrected col empty | No status | → NO_STATUS (Finance hasn't reviewed) |
| Both sheet and ERP empty | — | → MANDATORY_EMPTY |
| Sheet has data | Not in ERP | → CREATE (new record needed) |
| ERP has data | Not in sheet | → MISSING_IN_SHEET |
| Invalid value | — | → BLOCKED (e.g., "Signed" not a valid workflow state) |
| PED row "Cancelled" | — | → RELEASE_RESOURCE |

## ERP Action Column
Each result has an `erp_action` field:
- **UPDATE**: Will modify an existing ERP record
- **CREATE**: Will add a new record to ERP
- **SKIP**: No ERP change (match, not reviewed, blocked, etc.)

## Report Output
- **Terminal**: Project → 7 sheets → field-by-field with ERP value, Sheet value, Status, Action, Module, Target Field
- **CSV**: `reports/batchN_validation.csv` — one row per field check, columns: project, sheet, field, erp_value, sheet_value, status, erp_action, erp_module, erp_target_field, message, target_name, child_name
- **Totals Reconciliation**: At the end of each project, compares Sheet 2 totals against computed values from raw data

## File Structure
```
finance-data-migration/
├── CLAUDE.md              # This file
├── migrate.py             # CLI entry point (download, validate, apply, restore)
├── config.py              # Constants, field maps, column positions, workflow states
├── sharepoint.py          # Download Excel from Yazeed's OneDrive via Graph API
├── excel_parser.py        # Parse workbook into normalized Python dicts
├── erp_client.py          # ERPNext API client (read + write)
├── validators.py          # 6 validators (one per sheet), field-by-field comparison
├── invoice_matcher.py     # Cross-match short refs (00331-2025) ↔ full (ACC-SINV-2025-00331)
├── report.py              # Terminal + CSV report generation
├── backup.py              # JSON backup before apply, restore to revert
├── backups/               # Pre-change JSON snapshots
├── reports/               # Validation CSV reports
└── workbook.xlsx          # Downloaded workbook (auto-downloaded or manual)
```

## Quotation Child Tables (Sheet 3 & 5)
The Quotation has multiple custom child tables for pricing:
- `custom_sg_resources_price` — resource-level pricing (Sheet 3 maps here)
- `custom_other` — Third Party Cost Saudi (SMEs)
- `custom_tpc__osa` — Third Party Cost Outside Saudi
- `custom_travel_cost__egypt` — Egypt travel costs
- `custom_travel_cost__local` — Local travel costs
- `taxes` — VAT/charges
- `items` — single lump-sum line item (NOT individual deliverables)

## Task Doctype (Sheet 7 Deliverables)
Deliverable milestones are stored as `Task` records linked to the Project. Key custom fields:
- `custom_deliverable_name`, `custom_deliverable_amount`, `custom_amount_`
- `custom_estimated_invoice_date`, `custom_actual_invoice_date`
- `custom_invoice_no` (Link to Sales Invoice)
- `custom_invoice_payment_status`, `custom_collection_status`
- `custom_actual_payment_date`, `custom_expect_payment_date`
- `custom_work_start_date`, `custom_actual_start_date`, `custom_actual_end_date`

## Sheet 7: Two-Level Deliverables Validation
Multiple sheet deliverables can map to one Sales Invoice. Validation works in two levels:

### Level 1: Sheet deliverables vs Sales Invoice `items` child table
- Group sheet deliverables by invoice ref
- For each invoice: fetch SI detail, compare **count** (sheet deliverables vs SI items) and **amount sum**
- Per deliverable: match by name against SI `items.description` (HTML stripped), compare individual amounts
- This replaces the old per-row grand_total comparison that produced false UPDATEs

### Level 2: Sheet deliverables vs Task doctype
- Match each deliverable by name against Task `subject` or `custom_deliverable_name`
- Compare `custom_deliverable_amount`
- If no Task found → CREATE

## Project Date Fields
The Project doctype has TWO sets of date fields:
- `expected_start_date` / `expected_end_date` — standard ERPNext fields (NOT what the sheet maps to)
- `custom_official_start_date` / `official_end_date` — custom fields for official contract dates (Sheet 1 maps HERE)
- `custom_awarded_date` — awarding date (NOT `custom_awarding_date`)

## Quotation Lookup
Quotation is found via: Project → opportunity → Quotation. If that fails (e.g., Quotation not linked to opportunity), falls back to direct lookup by sheet reference.
- If Quotation exists but isn't linked to opportunity → note `NOT_LINKED`, still proceed with validation
- If Quotation belongs to a different project → `BLOCKED`, skip all field comparisons

## Expense Sources (Sheet 6) — GL Entry
Expenses use **GL Entry** as the single source of truth (per Yazeed's feedback).
- Query: `GL Entry` where `project = PROJ-XXXX`, group by `account`
- Expense accounts start with `5` (5xxx)
- Sum debit - credit per account for net amount
- **Sheet 6 is READ-ONLY** — validation only, never UPDATE the GL
- Previous approach (PPI + source doctypes) was replaced

## Partner / Engagement Manager Fields
- **Partner**: `project_lead` (Employee ID) / `project_lead_name` (name) — NOT `custom_partner_name_value`
- **Engagement Manager**: `project_manager` (Employee ID) / `project_manager_name` (name) — NOT `custom_engagement_manager`
- When updating, should look up Employee by name and set the Employee ID link field

## Customer Name Validation
- Use Project's existing `customer` field (CRM-CUS-*) — NOT fuzzy name search
- If Project already has a customer linked, compare `customer_name` against sheet value
- If no customer linked, do a lookup — but the Project's own link is the source of truth

## Sheet 3 Rate = Quotation Discount
- Sheet "Rate" column contains the **discount percentage** as decimal (e.g., 0.36 = 36%)
- ERP stores as percentage in `custom_discount_` (e.g., 36.0)
- Validator converts ERP to decimal before comparing
- This is a **Quotation-level field**, not per-resource

## Tax Handling (Sheet 5)
- For `Sales Tax & Charges` rows: only update qty=1 in the tax child row
- The Quotation module auto-calculates the tax total from the quotation price
- Do NOT push tax totals directly

## Read-Only Modules (never UPDATE)
- **GL Entry** (Sheet 6) — source of truth, read-only
- **Sales Invoice items** (Sheet 7) — informational comparison only
- **Payment Entry** (Sheet 7) — informational

## Missing/Deleted Rows
- **PED rows in ERP but not in sheet** → assumed correct (SKIP)
- **3P Planned rows in ERP but not in sheet** → marked TO_DELETE (not actually deleted)
- **Egypt employees** identified by `SG-E-*` prefix — different cost logic (invoices for past, rate card for future)

## Invoice Format Cross-Matching
Sheet uses short format: `00331-2025` (sequence-year)
ERP uses full format: `ACC-SINV-2025-00331` (prefix-year-sequence)
The `invoice_matcher.py` normalizes both to `(year, sequence)` tuples for matching.

## Project Workflow States
Valid states: Need Assigned Resources, Resources Assigned need approval, In Progress, Client Approval Pending, Approved, Rejected, Awarded, Running, On Hold, Completed, Cancelled, Investment, Expected to be Awarded.
- "Signed" is NOT valid — flagged as BLOCKED
- Transitions are constrained (e.g., Awarded → Need Assigned Resources → Running)

## ERP Environments
- **Production**: `https://engine.strategicgears.com`
- **UAT**: `https://strategicgears-uat.frappe.cloud` (current target in config.py)
- Same API token works for both
- Switch by changing `ERP_BASE` in `config.py`

## Slack Coordination
- Migration channel: `C09D0202CNP`, Thread: `1776150865.623519`
- Original channel: `C09QNU090NS`, Thread: `1773689016.118239`
- Stakeholders: Hashim (`U08GPAMV93Q`), Yazeed (`U08T55Y9SS2`), Abduljalil (`U067U3SR9RP`)

## Yazeed Feedback History
- **v1**: 15 comments — GL source, Task CREATE bug, customer lookup, Egypt resources, partner/EM fields
- **v2**: 16 comments — wrong customer match, partner=project_lead_name, GL/SI read-only, tax qty=1, TO_DELETE order
- **v3/v4**: Approved. One sheet error (PROJ-1313 tax) fixed by Yazeed.

## Apply Implementation Status
| Module | UPDATE | CREATE | Status |
|---|---|---|---|
| Project (Sheet 1) | Fields, dates, billability | — | Working |
| PED (Sheet 4) | to_date, ratio_ | New rows | Working |
| Quotation (Sheet 3) | custom_discount_, resources | New resource rows | **Not implemented** |
| Quotation (Sheet 5) | 3P costs, tax qty=1 | New 3P rows | **Not implemented** |
| Task (Sheet 7) | — | Full Task creation | **Not implemented** |
| GL Entry (Sheet 6) | — | — | Read-only (no apply) |
| Sales Invoice (Sheet 7) | — | — | Read-only (no apply) |

## Dependencies
- `openpyxl` — Excel parsing
- `requests` — ERP API calls
- `msal` — SharePoint auth
- `python-dotenv` — .env loading
- All available in `/Users/khalifah/Projects/sg/venv/`
