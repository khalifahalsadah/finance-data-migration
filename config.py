"""Configuration constants and field mappings for finance data migration."""

ERP_BASE = 'https://engine.strategicgears.com'
ERP_TOKEN = '96abe3a71ba471f:d82538d96d9c305'

# SharePoint / OneDrive
EXCEL_FILE_ID = '8B8607E9-7CA6-4384-957F-29E09FCFBE43'
EXCEL_USER = 'yazeed.su@strategicgears.com'
ENV_PATH = '/Users/khalifah/Projects/sg/.env'

# Valid Project workflow states
VALID_WORKFLOW_STATES = {
    'Need Assigned Resources', 'Resources Assigned need approval',
    'In Progress', 'Client Approval Pending', 'Approved', 'Rejected',
    'Awarded', 'Running', 'On Hold', 'Completed', 'Cancelled',
    'Investment', 'Expected to be Awarded',
}

# Sheet 1 → Project doctype field mapping
# (erp_col, corrected_col, status_col, label, api_field)
PROJECT_FIELDS = [
    (8, 9, 10, 'Project Status', 'workflow_state'),
    (11, 12, 13, 'Billability', 'custom_billability_type'),
    (14, 15, 16, 'OD Reference', 'custom_od_reference_'),
    (17, 18, 19, 'Project Name', 'project_name'),
    (20, 21, 22, 'Client Name', 'customer_name'),
    (23, 24, 25, 'End-Client Name', 'custom_end_client_name'),
    (26, 27, 28, 'Awarding Date', 'custom_awarded_date'),
    (29, 30, 31, 'Official Start Date', 'custom_official_start_date'),
    (32, 33, 34, 'Official End Date', 'official_end_date'),
    (35, 36, 37, 'Contractuals Link', 'custom_contractuals_link'),
    (38, 39, 40, 'Partner', 'custom_partner_name_value'),
    (41, 42, 43, 'Engagement Manager', 'custom_engagement_manager'),
]

# Sheet 3: Resources Planned column positions
# (erp_col, corrected_col) pairs
RESOURCES_PLANNED_COLS = {
    'level': (6, 7),
    'ratio': (8, 9),
    'uom': (10, 11),
    'qty': (12, 13),
    'rate': (14, 15),
}
RESOURCES_PLANNED_STATUS_COL = 16
RESOURCES_PLANNED_OPP_COL = 4
RESOURCES_PLANNED_QUOT_COL = 5

# Sheet 4: Resources Actual column positions
RESOURCES_ACTUAL_COLS = {
    'name': (4, 5),
    'role': (6, 7),
    'start': (8, 9),
    'end': (10, 11),
    'ratio': (12, 13),
    'rate': (14, 15),
}
RESOURCES_ACTUAL_STATUS_COL = 16

# Sheet 5: Third-Party Planned column positions
THIRDPARTY_PLANNED_COLS = {
    'entry_type': (6, 7),
    'description': (8, 9),
    'rate': (10, 11),
    'qty': (12, 13),
    'uom': (14, 15),
    'total': (16, 17),
}
THIRDPARTY_PLANNED_STATUS_COL = 18
THIRDPARTY_PLANNED_OPP_COL = 4
THIRDPARTY_PLANNED_QUOT_COL = 5

# Sheet 6: Third-Party Actual column positions
THIRDPARTY_ACTUAL_COLS = {
    'cost_type': (4, None),  # no corrected col for cost_type
    'vendor': (5, 6),
    'ref': (7, 8),
    'amount': (9, 10),
    'date': (11, 12),
}
THIRDPARTY_ACTUAL_STATUS_COL = 13

# Sheet 7: Deliverables & Invoices column positions
DELIVERABLES_COLS = {
    'deliverable': 4,
    'amount': 5,
    'planned_date': 6,
    'invoicing_status': 7,
    'actual_date': 8,
    'invoice_num': 9,
    'invoice_date': 10,
    'invoice_amount': 11,
    'invoice_status': 12,
    'payment_date': 13,
}
DELIVERABLES_STATUS_COL = 14

# Sheet 2: Totals column positions (for reconciliation)
TOTALS_FIELDS = [
    (4, 5, 6, 'Planned Cost (till date)'),
    (7, 8, 9, 'Planned Cost (till completion)'),
    (10, 11, 12, 'Actual Cost (till date)'),
    (13, 14, 15, 'Actual Cost (till completion)'),
    (16, 17, 18, 'Planned Revenue (till date)'),
    (19, 20, 21, 'Planned Revenue (till completion)'),
    (22, 23, 24, 'Actual Revenue'),
    (25, 26, 27, 'Planned Duration (months)'),
    (28, 29, 30, 'Actual Duration (months)'),
]

# Data row start positions per sheet
SHEET_DATA_START_ROW = {
    '1. Project Info': 5,
    '2. Totals': 5,
    '3. Resources (Planned)': 4,
    '4. Resources (Actual)': 4,
    '5. Third-Party (Planned)': 4,
    '6. Third-Party (Actual)': 4,
    '7. Deliverables & Invoices': 4,
}
