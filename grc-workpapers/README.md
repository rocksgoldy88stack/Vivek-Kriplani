# GRC / Oracle ITGC-ITAC Audit Workpaper Package

Big 4 style audit workpaper package for **Oracle Risk Management Cloud (RMC) / GRC** engagements,
covering ITGC + ITAC testing, Risk Control Matrices across 5 cycles, an 87-rule SoD ruleset,
ATC transaction monitoring models, role design, test of details, deficiency evaluation and a
remediation roadmap - plus formatted Word deliverables.

> For education / template use. Replace all sample data with actual engagement data.

## Files

| File | Description |
|------|-------------|
| `GRC_ITAC_Audit_Workpaper_Package.xlsx` | Multi-tab Excel workbook (20 tabs) with sample data, workings and supporting references |
| `Scoping_Memorandum.docx` | Formatted Word scoping memo deliverable |
| `Findings_Report.docx` | Formatted Word findings report (audit committee format) |
| `generate_workpapers.py` | Python (openpyxl) script that generates the workbook |
| `generate_word_docs.py` | Python (python-docx) script that generates the Word docs |

## Workbook Tabs (20)

| Tab | Content |
|-----|---------|
| `00_Index` | Master working paper index |
| `01_Scoping_Memo` | Scope, in-scope apps & processes, materiality |
| `02_IPE_Register` | Completeness & accuracy of reports used (IPE) |
| `03_RCM_P2P` | Risk Control Matrix - Procure to Pay |
| `04_RCM_O2C` | Risk Control Matrix - Order to Cash |
| `05_RCM_R2R` | Risk Control Matrix - Record to Report |
| `06_RCM_Inventory` | Risk Control Matrix - Inventory |
| `07_RCM_Payroll` | Risk Control Matrix - Payroll (Hire to Retire) |
| `08_ITGC_Testing` | Worked examples - Access & Change management testing |
| `09_ITGC_Templates` | Reusable test templates across 4 ITGC domains |
| `10_ITAC_Testing` | Automated application controls (3-way match, positive/negative testing) |
| `11_ATC_Transaction_Models` | 15 transaction monitoring models (duplicate payment, ghost vendor, etc.) |
| `12_SoD_Ruleset` | 87-rule production Segregation of Duties matrix |
| `13_SoD_Conflict_Analysis` | Oracle AAC scan results & disposition (path-based) |
| `14_Role_Design` | Custom role design worksheet (least privilege) |
| `15_Test_of_Details` | Depreciation reperformance (auto-calculating formulas) |
| `16_Deficiency_Log` | Exception log & severity assessment |
| `17_Remediation_Roadmap` | Phased management action plan |
| `18_Findings_Report` | Audit committee format findings |
| `19_Sampling_Guidance` | Sample size reference & tickmark legend |

## SoD Ruleset Coverage (87 rules)

P2P (15), O2C (12), R2R (10), Payroll/H2R (12), Inventory (8), Fixed Assets (7),
Treasury/Cash (8), Security/System Admin (8), Cross-process/Master Data (7).

## ATC Transaction Models (15)

Duplicate payments, ghost vendor (bank match), vendor=employee address, duplicate invoice,
split transactions, threshold gaming, round-dollar amounts, weekend postings, large manual
journals, dormant vendor reactivation, duplicate employees, post-termination payments,
new-vendor rushed payment, credit memo spike, invoice sequence anomaly.

## Regenerate

```bash
pip install openpyxl python-docx
python generate_workpapers.py   # builds the Excel workbook
python generate_word_docs.py    # builds the two Word deliverables
```
