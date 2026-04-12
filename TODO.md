# TODO — Finance Data Migration

## Next Session (2026-04-13)

### 1. Skip Third-Party Actual (Sheet 6) from apply
- Sheet 6 data comes from Purchase Invoice, Expense Claim, Journal Entry — these are source documents that shouldn't be modified via this tool
- Keep validation/reporting (useful for comparison) but skip from `apply --execute`
- The expenses are already in ERP from the source doctypes — the sheet is a review layer, not a correction source

### 2. Task Creation (Sheet 7 Deliverables)
When applying Sheet 7 deliverables, ensure Task creation includes ALL fields:
- `subject` — deliverable name
- `project` — link to PROJ-XXXX
- `custom_deliverable_name` — same as subject (Long Text field)
- `custom_deliverable_amount` — deliverable amount (SAR)
- `custom_amount_` — same amount (Float field, may be duplicate)
- `custom_estimated_invoice_date` — planned invoicing date
- `custom_actual_invoice_date` — actual invoicing date
- `custom_invoice_no` — Link to Sales Invoice (use matched SI name)
- `custom_invoice_payment_status` — invoicing status (Paid/Invoiced/Not Invoiced/Partially Paid)
- `custom_actual_payment_date` — payment date
- `custom_work_start_date` — if available
- `custom_client` — link to Customer
- `custom_client_name` — customer name
- `custom_project_name` — project name
- `custom_engagement_leader` — from Project
- `custom_project_manager` — from Project
- `status` — map from invoicing status (e.g., Paid → Completed, Not Invoiced → Open)

Verify:
- Check which fields are mandatory on the Task doctype before creating
- Test creation on one deliverable first (dry run → single create → verify in ERP UI)
- Ensure `custom_invoice_no` uses the full SI name (ACC-SINV-XXXX-XXXXX), not the short ref
