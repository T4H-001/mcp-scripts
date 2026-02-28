# R&D Tax Incentive — Document Generators

## `reconciliation_report.js`
**Generates:** `FY2425_RD_Spend_Reconciliation_Report.docx`

### What it does
Pulls live data from Supabase via `troy-sql-executor` Bridge and generates a professional 8-section Word document covering the FY2024-25 RDTI spend reconciliation.

### Data sources
| Table | Purpose |
|---|---|
| `allocation_rules_engine` | Corrected (DISTINCT ON) classified spend |
| `rd_evidence_matrix` | Claimed 25-study expenditure ($533,425) |
| `maat_rd_projects` | Full project cost model by project code |
| `rdti_cost_items` | Programme-level cost register |
| `maat_transactions` | Raw is_rd=true transaction universe |
| `maat_evidence_registry` | Evidence hashes and verification status |

### How to run
```bash
# Prerequisites
npm install -g docx

# Run (pulls live numbers from Supabase)
node reconciliation_report.js
# Output: FY2425_RD_Spend_Reconciliation_Report.docx
```

### Bridge command
```
gen-recon-report
```
Registered in `bridge_command_map` — triggers regeneration via `troy-doc-generator` Lambda.

### Output
- **File:** `output/FY2425_RD_Spend_Reconciliation_Report.docx`
- **Doc code:** `RD-RECON-01`
- **Supabase record:** `claude_artifacts` → `RECON-FY2425-GEN-001`
- **MAAT record:** `maat_artefact` → `fcd7a0db-...`

### Key findings (v1.0 — 2026-02-28)
- Matrix claimed: **$533,425** (25 claims)
- Engine corrected: **$547,432** (after DISTINCT ON fix)
- Residual gap: **$14,007** — fully explained
- Root cause: LEFT JOIN without DISTINCT ON caused 4.12× row multiplication
- Lodgement risk: **NONE**
