# TODO — Finance Data Migration

## Next: Implement Apply Logic for UAT

### 1. Task Creation (Sheet 7) — 108 items in Batch 1
Create Task records with all custom fields:
- `subject`, `project`, `custom_deliverable_name`, `custom_deliverable_amount`
- `custom_estimated_invoice_date`, `custom_actual_invoice_date`
- `custom_invoice_no` (Link to Sales Invoice — use full ACC-SINV-* name from invoice_matcher)
- `custom_invoice_payment_status`, `custom_actual_payment_date`
- `custom_client`, `custom_client_name`, `custom_project_name`
- `custom_engagement_leader`, `custom_project_manager`
- `status` — map from invoicing status
- Test on single deliverable first on UAT

### 2. Quotation Updates (Sheet 3/5)
- Update `custom_discount_` (quotation-level field) — convert decimal back to percentage
- Add new rows to `custom_sg_resources_price` child table
- Add new rows to `custom_other` / `custom_tpc__osa` / travel tables
- Tax handling: set tax row with type=VAT and qty=1, let module auto-calculate

### 3. PED New Rows (Sheet 4) — 5 items in Batch 1
- Add new rows to `distribution_detail` child table on existing PED

### 4. Skip Sheet 6 from Apply
- Sheet 6 (GL Entry) is read-only — ensure `apply` skips all Sheet 6 results
- Sheet 7 SI comparisons are read-only — ensure `apply` skips SI-targeted results

## Done (Validation)
- [x] All 7 sheet validators working with field-by-field comparison
- [x] GL Entry as expense source (Sheet 6)
- [x] Task vs SI items two-level validation (Sheet 7)
- [x] Customer validation against Project's linked Customer
- [x] Partner = project_lead_name, EM = project_manager_name
- [x] Rate = custom_discount_ (decimal comparison)
- [x] PED missing from sheet = assumed correct
- [x] 3P planned missing from sheet = TO_DELETE
- [x] SI/GL read-only (never UPDATE)
- [x] Tax noted for qty=1 approach
- [x] erp_action + erp_module + erp_target_field on every row
- [x] Yazeed v1 + v2 feedback addressed, v4 approved

## Pending (ERP Development — Not Migration Tool)
- [ ] End-Client Name field doesn't exist in ERP — needs custom field creation
- [ ] Contractuals Link field — needs custom field or file migration
- [ ] Partner/EM audit trail — child tables for change history
- [ ] Egypt resource cost logic — invoices for past, rate card for future
- [ ] Egypt quotation role mapping — "Director" doesn't exist in ERP
