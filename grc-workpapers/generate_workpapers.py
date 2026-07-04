"""
Big 4 style GRC / Oracle Risk Management Cloud (ITGC + ITAC) Audit Workpaper Package.
Generates a multi-tab Excel workbook with sample data, workings and supporting evidence
references. For educational / template use.
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ---------- Styling helpers ----------
NAVY = "1F3864"
BLUE = "2E5496"
LIGHT = "D9E1F2"
GREY = "F2F2F2"
GREEN = "C6EFCE"
RED = "FFC7CE"
AMBER = "FFEB9C"
WHITE = "FFFFFF"

thin = Side(style="thin", color="BFBFBF")
border = Border(left=thin, right=thin, top=thin, bottom=thin)

def title_cell(ws, cell, text, size=14):
    ws[cell] = text
    ws[cell].font = Font(bold=True, size=size, color=NAVY)

def header_row(ws, row, headers, start_col=1, fill=BLUE, fontcolor=WHITE):
    for i, h in enumerate(headers):
        c = ws.cell(row=row, column=start_col + i, value=h)
        c.font = Font(bold=True, color=fontcolor, size=10)
        c.fill = PatternFill("solid", fgColor=fill)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = border

def data_rows(ws, start_row, rows, start_col=1, zebra=True, colormap=None):
    """colormap: dict {col_index(0-based): {value:fill}} for conditional coloring."""
    for r, row in enumerate(rows):
        for i, val in enumerate(row):
            c = ws.cell(row=start_row + r, column=start_col + i, value=val)
            c.border = border
            c.alignment = Alignment(vertical="center", wrap_text=True)
            c.font = Font(size=10)
            if zebra and r % 2 == 1:
                c.fill = PatternFill("solid", fgColor=GREY)
            if colormap and i in colormap and str(val) in colormap[i]:
                c.fill = PatternFill("solid", fgColor=colormap[i][str(val)])
    return start_row + len(rows)

def set_widths(ws, widths, start_col=1):
    for i, w in enumerate(widths):
        ws.column_dimensions[get_column_letter(start_col + i)].width = w

def note(ws, cell, text, italic=True):
    ws[cell] = text
    ws[cell].font = Font(italic=italic, size=9, color="595959")
    ws[cell].alignment = Alignment(wrap_text=True, vertical="top")


wb = Workbook()

# =========================================================
# TAB 0: INDEX
# =========================================================
ws = wb.active
ws.title = "00_Index"
ws.sheet_properties.tabColor = NAVY
title_cell(ws, "B2", "GRC / ITGC-ITAC AUDIT WORKPAPER PACKAGE", 16)
note(ws, "B3", "Client: ABC Manufacturing Ltd.  |  Period: FY 2025-26  |  System: Oracle Fusion Cloud ERP (24C)")
note(ws, "B4", "Prepared for education/template use. Replace sample data with actual engagement data.")

header_row(ws, 6, ["WP Ref", "Section", "Deliverable / Tab", "Purpose"], start_col=2)
idx = [
    ["A-1", "Planning", "01_Scoping_Memo", "Scope, in-scope apps & processes, materiality"],
    ["A-2", "Planning", "02_IPE_Register", "Completeness & accuracy of reports used"],
    ["RCM", "Control Design", "03_RCM_P2P", "Risk Control Matrix - Procure to Pay"],
    ["C", "ITGC", "04_ITGC_Testing", "Access, Change, Operations, Backup tests"],
    ["D", "ITAC", "05_ITAC_Testing", "Automated application controls (3-way match etc.)"],
    ["E-1", "SoD", "06_SoD_Ruleset", "Segregation of Duties conflict rules"],
    ["E-2", "SoD", "07_SoD_Conflict_Analysis", "AAC scan results & disposition"],
    ["E-3", "Security", "08_Role_Design", "Custom role design worksheet"],
    ["F", "Substantive", "09_Test_of_Details", "Reperformance / recalculation"],
    ["G", "Deficiency", "10_Deficiency_Log", "Exception log & severity assessment"],
    ["H", "Remediation", "11_Remediation_Roadmap", "Management action plan & phasing"],
    ["I", "Reporting", "12_Findings_Report", "Formatted findings for audit committee"],
    ["-", "Reference", "13_Sampling_Guidance", "Sample size reference & tickmark legend"],
]
end = data_rows(ws, 7, idx, start_col=2)
set_widths(ws, [10, 16, 26, 55], start_col=2)

# =========================================================
# TAB 1: SCOPING MEMO
# =========================================================
ws = wb.create_sheet("01_Scoping_Memo")
ws.sheet_properties.tabColor = BLUE
title_cell(ws, "B2", "WP A-1: SCOPING MEMORANDUM")
meta = [
    ["Client", "ABC Manufacturing Ltd."],
    ["Period", "01-Apr-2025 to 31-Mar-2026 (FY 2025-26)"],
    ["System", "Oracle Fusion Cloud ERP (Release 24C)"],
    ["Objective", "Evaluate design & operating effectiveness of ITGC & ITAC over financially significant applications supporting ICFR / SOX 404"],
    ["Testing Approach", "Controls reliance + limited substantive"],
    ["Overall Materiality", "INR 12.5 Cr"],
    ["Performance Materiality", "INR 8.0 Cr"],
    ["Prepared By", "________________   Date: __/__/____"],
    ["Reviewed By", "________________   Date: __/__/____"],
]
r = 4
for k, v in meta:
    ws.cell(row=r, column=2, value=k).font = Font(bold=True, size=10)
    c = ws.cell(row=r, column=3, value=v)
    c.alignment = Alignment(wrap_text=True, vertical="center")
    c.font = Font(size=10)
    r += 1
set_widths(ws, [4, 24, 70], start_col=1)

ws.cell(row=r+1, column=2, value="IN-SCOPE APPLICATIONS").font = Font(bold=True, color=NAVY, size=11)
header_row(ws, r+2, ["Application / Module", "Significant?", "Basis"], start_col=2)
apps = [
    ["Oracle General Ledger", "Yes", "Financial statement close"],
    ["Oracle Payables", "Yes", "Expenses / Accounts Payable"],
    ["Oracle Receivables", "Yes", "Revenue / Accounts Receivable"],
    ["Oracle Fixed Assets", "Yes", "Depreciation / PP&E"],
    ["Oracle Inventory", "Yes", "Cost of Goods Sold"],
    ["Oracle HCM Payroll", "Yes", "Payroll expense"],
]
cmap = {1: {"Yes": GREEN}}
data_rows(ws, r+3, apps, start_col=2, colormap=cmap)
set_widths(ws, [4, 24, 14, 40], start_col=1)

# =========================================================
# TAB 2: IPE REGISTER
# =========================================================
ws = wb.create_sheet("02_IPE_Register")
ws.sheet_properties.tabColor = BLUE
title_cell(ws, "B2", "WP A-2: IPE (Information Produced by Entity) REGISTER")
note(ws, "B3", "Every report used in testing must have Completeness & Accuracy (C&A) validated - most common Big 4 review comment.")
header_row(ws, 5, ["Report Name", "Source System", "Used For", "Row Count", "C&A Validated?", "Validation Method", "Evidence Ref"], start_col=2)
ipe = [
    ["AP Aging Report", "Oracle Payables", "Sample selection - AP", "1,842", "Yes", "Reconciled to GL; parameters screenshot", "A-2.1"],
    ["Change Log Extract", "Oracle FSM", "ITGC change testing", "18", "Yes", "Full population; sequence completeness check", "A-2.2"],
    ["User Access Listing", "Security Console", "SoD / access testing", "342", "Yes", "Reconciled to HR active employee list", "A-2.3"],
    ["Fixed Asset Register", "Oracle Fixed Assets", "Depreciation reperformance", "1,240", "Yes", "Reconciled to GL net book value", "A-2.4"],
    ["Invoice Register", "Oracle Payables", "3-way match testing", "9,655", "Yes", "Reconciled to AP subledger control a/c", "A-2.5"],
]
cmap = {4: {"Yes": GREEN}}
data_rows(ws, 6, ipe, start_col=2, colormap=cmap)
set_widths(ws, [4, 22, 18, 22, 12, 15, 34, 12], start_col=1)

# =========================================================
# TAB 3: RCM - P2P
# =========================================================
ws = wb.create_sheet("03_RCM_P2P")
ws.sheet_properties.tabColor = "375623"
title_cell(ws, "B2", "RISK CONTROL MATRIX - PROCURE TO PAY (P2P)")
headers = ["Risk ID", "Sub-process", "Risk Description", "Rating", "Control ID",
           "Control Description", "Type", "Nature", "Frequency", "Owner",
           "SoD Impact", "Assertion", "Test Procedure"]
header_row(ws, 4, headers, start_col=2)
rcm = [
    ["P2P-R-01", "Requisitioning", "Purchases made without budget/authorization", "High", "P2P-C-01",
     "System enforces budget check & approval hierarchy on requisition", "Preventive", "Automated", "Per txn", "Procurement Mgr", "No", "Occurrence", "Config review + 1 positive/negative test"],
    ["P2P-R-02", "Purchase Order", "Unauthorized buyer creates PO", "High", "P2P-C-02",
     "Only authorized buyers can create PO; approval limits enforced", "Preventive", "Automated", "Per txn", "Procurement Mgr", "Yes", "Occurrence", "Config + access review"],
    ["P2P-R-03", "Goods Receipt", "Goods recorded not actually received", "Med", "P2P-C-03",
     "Receiver segregated from buyer; GRN matched to PO", "Preventive", "IT-Dependent", "Per txn", "Warehouse Head", "Yes", "Existence", "Walkthrough + sample 25"],
    ["P2P-R-04", "Invoicing", "Duplicate invoice / payment", "High", "P2P-C-04",
     "System blocks duplicate invoice number per supplier", "Preventive", "Automated", "Per txn", "AP Manager", "No", "Accuracy", "Config + negative test"],
    ["P2P-R-05", "Matching", "Invoice paid despite PO/GRN mismatch", "High", "P2P-C-05",
     "Automated 3-way match; hold applied if variance > tolerance", "Preventive", "Automated", "Per txn", "AP Manager", "No", "Accuracy/Valuation", "Config + positive/negative test"],
    ["P2P-R-06", "Payment", "Unauthorized / fraudulent payment", "High", "P2P-C-06",
     "Payment approval segregated from invoice entry; dual control on bank", "Preventive", "Manual", "Per txn", "Treasury Head", "Yes", "Occurrence", "Sample 25 + SoD scan"],
    ["P2P-R-07", "Master Data", "Fake supplier created for fraud", "High", "P2P-C-07",
     "Supplier creation segregated from payment; new supplier approval", "Preventive", "IT-Dependent", "Per event", "Master Data Team", "Yes", "Occurrence", "Sample + SoD scan"],
    ["P2P-R-08", "Reconciliation", "Errors in AP not detected timely", "Med", "P2P-C-08",
     "Monthly AP subledger to GL reconciliation by independent person", "Detective", "Manual", "Monthly", "AP Manager", "Yes", "Completeness", "Sample 2 months"],
]
cmap = {3: {"High": RED, "Med": AMBER}, 8: {"Automated": GREEN}, 10: {"Yes": AMBER}}
data_rows(ws, 5, rcm, start_col=2, colormap=cmap)
set_widths(ws, [4, 10, 12, 30, 8, 12, 34, 12, 12, 10, 14, 9, 15, 30], start_col=1)

# =========================================================
# TAB 4: ITGC TESTING
# =========================================================
ws = wb.create_sheet("04_ITGC_Testing")
ws.sheet_properties.tabColor = "7030A0"
title_cell(ws, "B2", "WP C: ITGC TESTING - Access to Programs & Data (Sample)")
note(ws, "B3", "Control ITGC-01: New user access authorized before provisioning. Population=47, Sample=25 (random). Legend: Y=verified, N=exception.")
header_row(ws, 5, ["Sr", "User ID", "Access Request Form?", "Approved By", "Approval before Grant?", "Result", "Evidence Ref"], start_col=2)
itgc = [
    [1, "jsmith", "Y", "Reporting Manager", "Y", "Pass", "C-1.1"],
    [2, "rpatel", "Y", "Dept Head", "Y", "Pass", "C-1.2"],
    [3, "akhan", "N", "N/A", "N", "FAIL", "C-1.3"],
    [4, "slodha", "Y", "Reporting Manager", "Y", "Pass", "C-1.4"],
    [5, "dmehta", "Y", "Dept Head", "Y", "Pass", "C-1.5"],
    [6, "vnair", "Y", "Reporting Manager", "Y", "Pass", "C-1.6"],
    ["...", "...", "...", "...", "...", "...", "..."],
    [25, "mgupta", "Y", "Reporting Manager", "Y", "Pass", "C-1.25"],
]
cmap = {4: {"N": RED, "Y": GREEN}, 5: {"FAIL": RED, "Pass": GREEN}}
end = data_rows(ws, 6, itgc, start_col=2, colormap=cmap)
ws.cell(row=end+1, column=2, value="EXCEPTIONS: 1 of 25 (akhan - no approval evidence). Exception rate 4%.").font = Font(bold=True, color=RED, size=10)
ws.cell(row=end+2, column=2, value="CONCLUSION: Control NOT operating effectively. Refer Deficiency Log item D-03.").font = Font(bold=True, size=10)
set_widths(ws, [4, 6, 12, 20, 18, 22, 12, 12], start_col=1)

# Change mgmt block
sr = end + 5
ws.cell(row=sr, column=2, value="WP C-2: CHANGE MANAGEMENT (Sample)").font = Font(bold=True, color=NAVY, size=12)
note(ws, f"B{sr+1}", "ITGC-05: Changes tested & approved before prod. Population=18, Sample=8. A1=CR documented, A2=UAT sign-off, A3=Approved before migration, A4=Dev != Migrator.")
header_row(ws, sr+2, ["Sr", "CR#", "A1", "A2", "A3", "A4", "Result"], start_col=2)
chg = [
    [1, "CR-1021", "Y", "Y", "Y", "Y", "Pass"],
    [2, "CR-1044", "Y", "Y", "Y", "N", "FAIL"],
    [3, "CR-1052", "Y", "Y", "Y", "Y", "Pass"],
    [4, "CR-1061", "Y", "Y", "Y", "Y", "Pass"],
    ["...", "...", "...", "...", "...", "...", "..."],
    [8, "CR-1099", "Y", "Y", "Y", "Y", "Pass"],
]
cmap2 = {5: {"N": RED, "Y": GREEN}, 6: {"FAIL": RED, "Pass": GREEN}}
end2 = data_rows(ws, sr+3, chg, start_col=2, colormap=cmap2)
ws.cell(row=end2+1, column=2, value="EXCEPTION: CR-1044 - Developer 'dev01' also had migration access (SoD - links to E-2 / D-05).").font = Font(bold=True, color=RED, size=10)

# =========================================================
# TAB 5: ITAC TESTING
# =========================================================
ws = wb.create_sheet("05_ITAC_Testing")
ws.sheet_properties.tabColor = "C55A11"
title_cell(ws, "B2", "WP D: ITAC TESTING - Automated Application Controls")
ws.cell(row=4, column=2, value="D-1: 3-Way Match Control (ITAC-P2P-05)").font = Font(bold=True, color=NAVY, size=12)

ws.cell(row=6, column=2, value="TEST OF DESIGN - Config verified (Payables Options > Matching)").font = Font(bold=True, size=10)
header_row(ws, 7, ["Config Parameter", "Expected", "Actual", "Result", "Evidence"], start_col=2)
tod = [
    ["Invoice Match Option", "Purchase Order", "Purchase Order", "Pass", "D-1.1"],
    ["Quantity Tolerance", "0%", "0%", "Pass", "D-1.1"],
    ["Price Tolerance", "2% (per policy)", "2%", "Pass", "D-1.1"],
    ["Match Approval Level", "3-Way", "3-Way", "Pass", "D-1.1"],
]
cmap = {3: {"Pass": GREEN}}
end = data_rows(ws, 8, tod, start_col=2, colormap=cmap)

sr = end + 2
ws.cell(row=sr, column=2, value="TEST OF ONE - Positive & Negative testing").font = Font(bold=True, size=10)
header_row(ws, sr+1, ["Test Type", "Invoice #", "Scenario", "Expected Behaviour", "Actual", "Result", "Evidence"], start_col=2)
toe = [
    ["Positive", "INV-4521", "Qty & price within tolerance", "No hold - allow payment", "No hold applied", "Pass", "D-1.2"],
    ["Negative", "INV-6698", "Price 8% > PO (exceeds 2% tol)", "Price hold applied; payment blocked", "Price hold auto-applied; blocked", "Pass", "D-1.3"],
]
end2 = data_rows(ws, sr+2, toe, start_col=2, colormap={5: {"Pass": GREEN}})
ws.cell(row=end2+2, column=2, value="ITGC DEPENDENCY: Config unchanged during period (per change log WP C-2). Test of one valid for full period.").font = Font(italic=True, size=9, color="595959")
ws.cell(row=end2+3, column=2, value="CONCLUSION: Control operating effectively. No exceptions.").font = Font(bold=True, color="375623", size=10)
set_widths(ws, [4, 16, 18, 30, 26, 12, 12], start_col=1)

# =========================================================
# TAB 6: SoD RULESET
# =========================================================
ws = wb.create_sheet("06_SoD_Ruleset")
ws.sheet_properties.tabColor = "C00000"
title_cell(ws, "B2", "WP E-1: SEGREGATION OF DUTIES (SoD) RULESET")
note(ws, "B3", "Core principle: no single person controls a complete transaction cycle. Ruleset used to configure Oracle AAC Access Models.")
header_row(ws, 5, ["Rule ID", "Process", "Function A", "Function B (Conflicts)", "Risk", "Business Impact"], start_col=2)
sod = [
    ["SOD-01", "P2P", "Maintain Supplier", "Process AP Payment", "High", "Fake supplier created and self-paid"],
    ["SOD-02", "R2R", "Create Journal", "Approve Journal", "High", "Fraudulent entry self-approved"],
    ["SOD-03", "Treasury", "Maintain Bank Account", "Make Payment", "High", "Payment diverted to own account"],
    ["SOD-04", "P2P", "Enter Purchase Order", "Approve Purchase Order", "High", "Unauthorized purchase self-approved"],
    ["SOD-05", "O2C", "Maintain Customer", "Post Cash Receipt", "Med", "Misapplied receipts / lapping"],
    ["SOD-06", "Security", "User Administration", "Any Financial Transaction", "Critical", "Self-grant access & post entries"],
    ["SOD-07", "Payroll", "Maintain Payroll Data", "Approve Payroll Run", "High", "Ghost employee / inflated pay"],
    ["SOD-08", "P2P", "Enter Invoice", "Approve Payment", "High", "Fictitious invoice self-paid"],
    ["SOD-09", "Inventory", "Adjust Inventory", "Physical Custody of Stock", "Med", "Theft masked by adjustment"],
    ["SOD-10", "R2R", "Post Journal", "Perform Reconciliation", "Med", "Errors/fraud concealed in recon"],
    ["SOD-11", "FA", "Create/Retire Asset", "Approve Asset Disposal", "Med", "Unauthorized asset disposal"],
    ["SOD-12", "O2C", "Approve Credit Limit", "Create Sales Order", "Med", "Sales beyond approved credit"],
]
cmap = {4: {"Critical": "FF0000", "High": RED, "Med": AMBER}}
data_rows(ws, 6, sod, start_col=2, colormap=cmap)
note(ws, "B20", "Note: Full production ruleset typically contains 60-100+ rules. Above is a representative extract.")
set_widths(ws, [4, 10, 12, 24, 26, 12, 40], start_col=1)

# =========================================================
# TAB 7: SoD CONFLICT ANALYSIS
# =========================================================
ws = wb.create_sheet("07_SoD_Conflict_Analysis")
ws.sheet_properties.tabColor = "C00000"
title_cell(ws, "B2", "WP E-2: SoD CONFLICT ANALYSIS (Oracle AAC output)")
note(ws, "B3", "Population: 342 active users scanned via Oracle AAC (path-based). 64 users with 1+ SoD violation.")
header_row(ws, 5, ["Rule", "# Users", "Severity", "Disposition", "Comment"], start_col=2)
summ = [
    ["SOD-01", 12, "High", "Remediate", "Remove supplier maintenance access"],
    ["SOD-02", 8, "High", "Mitigate", "Independent monthly journal review"],
    ["SOD-04", 23, "High", "Remediate", "Split PO entry vs approval roles"],
    ["SOD-06", 3, "Critical", "Remediate", "Immediate - remove User Admin from finance"],
    ["SOD-08", 11, "High", "Remediate", "Custom role redesign"],
    ["SOD-others", 7, "Med", "Accept/Mitigate", "Compensating controls documented"],
]
cmap = {3: {"Critical": "FF0000", "High": RED, "Med": AMBER}, 4: {"Remediate": LIGHT}}
end = data_rows(ws, 6, summ, start_col=2, colormap=cmap)
set_widths(ws, [4, 12, 10, 12, 16, 45], start_col=1)

sr = end + 2
ws.cell(row=sr, column=2, value="DETAILED CONFLICT SAMPLE (SOD-01) - path-based analysis").font = Font(bold=True, color=NAVY, size=11)
detail = [
    ["User", "bshah"],
    ["Conflicting Roles", "Custom_AP_Clerk + Supplier_Maintenance"],
    ["Path A", "bshah -> AP_Clerk_Job -> Payables_Payment_Duty -> 'Create Payment' privilege"],
    ["Path B", "bshah -> Supplier_Job -> Supplier_Mgmt_Duty -> 'Create Supplier' privilege"],
    ["Risk", "Can create fake supplier and self-pay"],
    ["Interim Mitigation", "AP Manager reviews new-supplier + payment report weekly (CC-01)"],
    ["Remediation", "Remove Supplier_Maintenance role by 31-Aug-2026"],
]
rr = sr + 1
for k, v in detail:
    ws.cell(row=rr, column=2, value=k).font = Font(bold=True, size=10)
    ws.cell(row=rr, column=3, value=v).font = Font(size=10)
    ws.cell(row=rr, column=3).alignment = Alignment(wrap_text=True)
    rr += 1

# =========================================================
# TAB 8: ROLE DESIGN
# =========================================================
ws = wb.create_sheet("08_Role_Design")
ws.sheet_properties.tabColor = "203864"
title_cell(ws, "B2", "WP E-3: CUSTOM ROLE DESIGN WORKSHEET")
note(ws, "B3", "Custom role: XX_AP_MANAGER_CUSTOM (copied from seeded 'Accounts Payable Manager'). Principle: least privilege + SoD clean.")
header_row(ws, 5, ["Duty Role", "Action", "Reason", "SoD Rule Addressed"], start_col=2)
rd = [
    ["Supplier Profile Management Duty", "REMOVE", "Conflicts with payment (SOD-01)", "SOD-01"],
    ["Bank Account Setup Duty", "REMOVE", "Conflicts with payment (SOD-03)", "SOD-03"],
    ["Payables Invoice Approval Duty", "RETAIN", "Core manager function", "-"],
    ["Payables Payment Approval Duty", "RETAIN", "Core manager function", "-"],
    ["AP Inquiry & Reporting Duty", "RETAIN", "Read-only reporting", "-"],
    ["User Account Management Duty", "REMOVE", "Conflicts with any txn (SOD-06)", "SOD-06"],
]
cmap = {1: {"REMOVE": RED, "RETAIN": GREEN}}
end = data_rows(ws, 6, rd, start_col=2, colormap=cmap)
ws.cell(row=end+1, column=2, value="DATA SECURITY: Business Unit = 'India Operations' only.").font = Font(bold=True, size=10)
ws.cell(row=end+2, column=2, value="SoD RE-SCAN AFTER DESIGN (simulated in AAC before deploy): 0 conflicts - PASS.").font = Font(bold=True, color="375623", size=10)
set_widths(ws, [4, 34, 12, 34, 18], start_col=1)

# =========================================================
# TAB 9: TEST OF DETAILS
# =========================================================
ws = wb.create_sheet("09_Test_of_Details")
ws.sheet_properties.tabColor = "548235"
title_cell(ws, "B2", "WP F-1: TEST OF DETAILS - Depreciation Reperformance")
note(ws, "B3", "Objective: verify accuracy of FA depreciation (Accuracy/Valuation). Population=1,240; Sample=30 (monetary + random). Formula auto-computes.")
header_row(ws, 5, ["Sr", "Asset ID", "Cost (INR)", "Life (yrs)", "Method", "System Dep", "Recalc Dep", "Diff", "Result"], start_col=2)
# We'll write with formulas for recalc (SLM = cost/life)
assets = [
    [1, "FA-1001", 500000, 5, "SLM", 100000],
    [2, "FA-1044", 240000, 4, "SLM", 60000],
    [3, "FA-1120", 750000, 5, "SLM", 150000],
    [4, "FA-1355", 360000, 3, "SLM", 120000],
    [5, "FA-1502", 900000, 10, "SLM", 90000],
    [6, "FA-1780", 480000, 4, "SLM", 120000],
    [7, "FA-2210", 900000, 3, "SLM", 300300],  # exception
]
start = 6
for r, a in enumerate(assets):
    row = start + r
    for i, val in enumerate(a):
        c = ws.cell(row=row, column=2 + i, value=val)
        c.border = border; c.font = Font(size=10)
        c.alignment = Alignment(vertical="center")
    # recalc dep col H (col 8 => letter), cost col D(4->'D'? actual col index: B=2 Sr, C=3 AssetID, D=4 Cost, E=5 Life, F=6 Method, G=7 SysDep, H=8 Recalc, I=9 Diff, J=10 Result
    ws.cell(row=row, column=8).value = f"=ROUND(D{row}/E{row},0)"   # recalc
    ws.cell(row=row, column=9).value = f"=G{row}-H{row}"            # diff
    ws.cell(row=row, column=10).value = f'=IF(ABS(I{row})<1,"Pass","EXCEPTION")'
    for cc in range(8, 11):
        ws.cell(row=row, column=cc).border = border
        ws.cell(row=row, column=cc).font = Font(size=10)
endrow = start + len(assets)
ws.cell(row=endrow+1, column=2, value="EXCEPTION: FA-2210 - rounding config diff INR 300 (immaterial).").font = Font(bold=True, color="C55A11", size=10)
ws.cell(row=endrow+2, column=2, value="CONCLUSION: Depreciation calculation accurate. Immaterial difference noted, no impact.").font = Font(bold=True, size=10)
set_widths(ws, [4, 6, 12, 12, 10, 10, 12, 12, 10, 12], start_col=1)

# =========================================================
# TAB 10: DEFICIENCY LOG
# =========================================================
ws = wb.create_sheet("10_Deficiency_Log")
ws.sheet_properties.tabColor = "BF8F00"
title_cell(ws, "B2", "WP G: DEFICIENCY / EXCEPTION LOG & SEVERITY")
header_row(ws, 4, ["Def #", "Deficiency Description", "Ref WP", "Could cause misstatement?", "Magnitude", "Likelihood", "Compensating Control?", "Severity", "Status"], start_col=2)
defs = [
    ["D-03", "1 user provisioned without approval evidence", "C-1", "Yes (limited)", "Low", "Remote", "Periodic access review", "Control Deficiency", "Open"],
    ["D-05", "Developer had prod migration access (SoD)", "C-2/E-2", "Yes", "Moderate", "Reasonably possible", "None effective", "Significant Deficiency", "Open"],
    ["D-08", "3 users: User Admin + financial txn access", "E-2", "Yes", "Material", "Reasonably possible", "None", "MATERIAL WEAKNESS", "Open"],
    ["D-11", "Rounding diff in depreciation config", "F-1", "Yes (immaterial)", "Trivial", "Remote", "GL recon", "Control Deficiency", "Open"],
]
cmap = {7: {"MATERIAL WEAKNESS": "FF0000", "Significant Deficiency": RED, "Control Deficiency": AMBER}}
data_rows(ws, 5, defs, start_col=2, colormap=cmap)
note(ws, "B11", "Severity logic: (1) Could it cause misstatement? (2) Magnitude? (3) Likelihood? (4) Compensating controls? -> classify Control Deficiency / Significant Deficiency / Material Weakness.")
set_widths(ws, [4, 8, 32, 10, 18, 12, 16, 18, 20, 10], start_col=1)

# =========================================================
# TAB 11: REMEDIATION ROADMAP
# =========================================================
ws = wb.create_sheet("11_Remediation_Roadmap")
ws.sheet_properties.tabColor = "2E75B6"
title_cell(ws, "B2", "WP H: REMEDIATION ROADMAP / MANAGEMENT ACTION PLAN")
header_row(ws, 4, ["Def #", "Remediation Action", "Priority", "Phase", "Owner", "Target Date", "Validation Method", "Status"], start_col=2)
rem = [
    ["D-08", "Remove User Admin duty from 3 finance users", "P1-Critical", "Phase 1 (0-30d)", "IT Security Head", "15-Jul-2026", "AAC re-scan = 0 conflicts", "In Progress"],
    ["D-05", "Segregate developer & migration access", "P2-High", "Phase 2 (1-3m)", "IT Manager", "31-Aug-2026", "Change log + SoD re-test", "Planned"],
    ["SOD-04", "Redesign PO custom roles (entry vs approval)", "P2-High", "Phase 2 (1-3m)", "GRC Lead", "31-Aug-2026", "Role re-scan + UAT", "Planned"],
    ["D-03", "Enforce access request workflow before grant", "P3-Med", "Phase 3 (3-6m)", "HR + IT", "30-Sep-2026", "Re-test sample of new users", "Planned"],
    ["Monitoring", "Deploy AAC/ATC continuous monitoring", "P3-Med", "Phase 3 (3-6m)", "GRC Lead", "31-Oct-2026", "Dashboard live + incident SLA", "Planned"],
    ["Cert", "Implement quarterly access certification", "P3-Med", "Phase 3 (3-6m)", "GRC Lead", "31-Oct-2026", "First cycle completed", "Planned"],
]
cmap = {3: {"P1-Critical": RED, "P2-High": AMBER}, 7: {}}
data_rows(ws, 5, rem, start_col=2, colormap=cmap)
sr = 13
ws.cell(row=sr, column=2, value="PHASED APPROACH").font = Font(bold=True, color=NAVY, size=11)
ph = [
    ["Phase 1 (0-30 days)", "Critical access removal, disable super-users, obvious SoD fixes - quick wins"],
    ["Phase 2 (1-3 months)", "Custom role redesign, AAC model deployment, approval workflow fixes"],
    ["Phase 3 (3-6 months)", "Continuous monitoring, periodic access certification, training"],
]
rr = sr + 1
for k, v in ph:
    ws.cell(row=rr, column=2, value=k).font = Font(bold=True, size=10)
    ws.cell(row=rr, column=4, value=v).font = Font(size=10)
    ws.cell(row=rr, column=4).alignment = Alignment(wrap_text=True)
    rr += 1
set_widths(ws, [4, 10, 34, 12, 16, 16, 14, 26, 12], start_col=1)

# =========================================================
# TAB 12: FINDINGS REPORT
# =========================================================
ws = wb.create_sheet("12_Findings_Report")
ws.sheet_properties.tabColor = "808080"
title_cell(ws, "B2", "WP I: FINDINGS REPORT (Audit Committee Format)")
header_row(ws, 4, ["Finding #", "Title", "Rating", "Condition", "Criteria", "Cause", "Effect", "Recommendation", "Mgmt Response"], start_col=2)
find = [
    ["D-08", "Excessive privileged access - SoD violation", "Material Weakness",
     "3 finance users hold User Administration + transaction posting in Oracle Fusion",
     "SoD principle - admin must be segregated from transaction processing (Policy SEC-04 / SOX)",
     "Roles granted during migration; never reviewed",
     "Risk of unauthorized access grant & fraudulent entries; potential material misstatement",
     "Remove User Admin duty; implement quarterly access certification",
     "Agreed. Owner: IT Security Head. Target 15-Jul-2026"],
    ["D-05", "Inadequate change segregation", "Significant Deficiency",
     "Developer had production migration access",
     "Dev must be segregated from prod migration (change mgmt policy)",
     "No enforced segregation in deployment process",
     "Unauthorized/untested changes could reach production",
     "Segregate roles; enforce migration approval gate",
     "Agreed. Owner: IT Manager. Target 31-Aug-2026"],
]
cmap = {2: {"Material Weakness": "FF0000", "Significant Deficiency": RED}}
data_rows(ws, 5, find, start_col=2, colormap=cmap)
set_widths(ws, [4, 8, 22, 16, 26, 26, 22, 28, 26, 24], start_col=1)

# =========================================================
# TAB 13: SAMPLING GUIDANCE
# =========================================================
ws = wb.create_sheet("13_Sampling_Guidance")
ws.sheet_properties.tabColor = "A6A6A6"
title_cell(ws, "B2", "REFERENCE: SAMPLING GUIDANCE & TICKMARK LEGEND")
header_row(ws, 4, ["Control Frequency", "Approx Population", "Suggested Sample Size"], start_col=2)
samp = [
    ["Annual", "1", "1"],
    ["Quarterly", "4", "2"],
    ["Monthly", "12", "2 - 5"],
    ["Weekly", "52", "5 - 15"],
    ["Daily", "250+", "15 - 25"],
    ["Multiple times per day", "Many", "25 - 40"],
    ["Automated (config)", "N/A", "1 (test of one) + ITGC reliance"],
]
data_rows(ws, 5, samp, start_col=2)
sr = 14
ws.cell(row=sr, column=2, value="TICKMARK LEGEND").font = Font(bold=True, color=NAVY, size=11)
legend = [
    ["Y / ✓", "Verified to supporting evidence"],
    ["N / ✗", "Exception noted"],
    ["Pass", "Attribute met, no exception"],
    ["FAIL", "Attribute not met - deficiency"],
    ["E", "Evidence attached (referenced)"],
    ["N/A", "Not applicable"],
]
rr = sr + 1
for k, v in legend:
    ws.cell(row=rr, column=2, value=k).font = Font(bold=True, size=10)
    ws.cell(row=rr, column=3, value=v).font = Font(size=10)
    rr += 1
set_widths(ws, [4, 22, 30, 20], start_col=1)

# Freeze header panes on data-heavy tabs
for name, cell in [("03_RCM_P2P", "A5"), ("06_SoD_Ruleset", "A6"), ("07_SoD_Conflict_Analysis", "A6"),
                   ("10_Deficiency_Log", "A5"), ("11_Remediation_Roadmap", "A5")]:
    wb[name].freeze_panes = cell

out = "/projects/sandbox/GRC_ITAC_Audit_Workpaper_Package.xlsx"
wb.save(out)
print("Saved:", out)
print("Tabs:", wb.sheetnames)
