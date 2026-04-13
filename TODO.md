# TODO — Finance Data Migration

## Yazeed's Comments (2026-04-13) — All Items

### DONE
- [x] Bug fix: Task MATCH on non-existent Task — show CREATE not MATCH for deliverable fields when Task doesn't exist
- [x] Sheet 6 source: Use GL Entry (project field) instead of PPI/source doctypes — group by account for breakdown
- [x] Missing in sheet for 3P planned = mark as TO_DELETE (not actually delete)
- [x] Missing in sheet for PED = assume correct (SKIP, not MISSING_IN_SHEET)
- [x] Customer name: validate against Customer doctype (CRM-CUS-*), match closest
- [x] Egypt employee identification: SG-E-* prefix for Egypt resources
- [x] Rate field: confirmed as custom_discount_ (already fixed)
- [x] Skip Sheet 6 from apply (read-only validation from GL)

### PENDING (ERP Development Required)
- [ ] End-Client Name field (`custom_end_client_name`) — doesn't exist in ERP, needs to be created
- [ ] Contractuals Link field — needs custom field creation or file migration approach
- [ ] Partner/EM audit trail — child tables for change history (ERP feature request)
- [ ] Egypt quotation role mapping — "Director" role doesn't exist, mapped as "Manager working in Saudi"
- [ ] Tax entry validation — ensure computed tax total matches quotation

### Task Creation (Sheet 7)
- [ ] Implement Task creation with all custom fields
- [ ] Test on single deliverable before batch
