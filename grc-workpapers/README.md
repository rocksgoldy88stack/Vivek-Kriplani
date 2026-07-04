# GRC / Oracle ITGC-ITAC Audit Workpaper Package

Big 4 style audit workpaper package for **Oracle Risk Management Cloud (RMC) / GRC** engagements,
covering ITGC + ITAC testing, Risk Control Matrix, SoD analysis, role design, test of details,
deficiency evaluation and remediation roadmap.

> For education / template use. Replace all sample data with actual engagement data.

## Files

| File | Description |
|------|-------------|
| `GRC_ITAC_Audit_Workpaper_Package.xlsx` | Multi-tab Excel workbook (14 tabs) with sample data, workings and supporting references |
| `generate_workpapers.py` | Python (openpyxl) script that generates the workbook |

## Workbook Tabs

| Tab | Content |
|-----|---------|
| `00_Index` | Master working paper index |
| `01_Scoping_Memo` | Scope, in-scope apps & processes, materiality |
| `02_IPE_Register` | Completeness & accuracy of reports used (IPE) |
| `03_RCM_P2P` | Risk Control Matrix - Procure to Pay |
| `04_ITGC_Testing` | Access & Change management attribute testing |
| `05_ITAC_Testing` | Automated application controls (3-way match, positive/negative testing) |
| `06_SoD_Ruleset` | Segregation of Duties conflict rules |
| `07_SoD_Conflict_Analysis` | Oracle AAC scan results & disposition (path-based) |
| `08_Role_Design` | Custom role design worksheet (least privilege) |
| `09_Test_of_Details` | Depreciation reperformance (auto-calculating formulas) |
| `10_Deficiency_Log` | Exception log & severity assessment |
| `11_Remediation_Roadmap` | Phased management action plan |
| `12_Findings_Report` | Audit committee format findings |
| `13_Sampling_Guidance` | Sample size reference & tickmark legend |

## Regenerate

```bash
pip install openpyxl
python generate_workpapers.py
```
