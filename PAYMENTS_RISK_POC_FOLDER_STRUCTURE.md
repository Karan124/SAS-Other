# Payments Risk PoC - Complete Folder Structure

## Root Directory: /payments_risk_poc/

```
/payments_risk_poc/
│
├── 00_documentation/
│   ├── juno_profile_report.html                      # Sprint 1: JUNO profiling results
│   ├── holocentric_profile_report.html               # Sprint 1: Holocentric profiling results
│   ├── linkage_profile_report.html                   # Sprint 1: Linkage tables profiling
│   ├── data_quality_findings.md                      # Sprint 1: Data quality issues summary
│   ├── reusability_handover.docx                     # Sprint 9: Reusability documentation
│   └── technical_runbook_v1_5.docx                   # Original technical runbook
│
├── 01_prompts/
│   ├── step02_process_inclusion_pp6.txt              # Sprint 3: Process embedding classification prompt
│   ├── step03_payment_nature_processes.txt           # Sprint 3: Payment Nature on processes prompt
│   ├── step04_control_inclusion_p6.txt               # Sprint 3: Control embedding classification prompt
│   ├── step05a_holocentric_bespoke.txt               # Sprint 4: Bespoke discovery Holocentric Tasks prompt
│   ├── step05c_incident_bespoke.txt                  # Sprint 4: Bespoke discovery incident remediation prompt
│   ├── step07b_payment_nature_controls_fallback.txt  # Sprint 5: Payment Nature on controls (unlinked) prompt
│   ├── step08_ct_mapping.txt                         # Sprint 6: Master Controls CT1-CT28 mapping prompt
│   ├── step09_lifecycle_placement.txt                # Sprint 6: Lifecycle stages A-G placement prompt
│   └── prompt_versions.json                          # Prompt version tracking with SHA-256 hashes
│
├── 02_code/
│   │
│   ├── sas/
│   │   ├── step01_profile_and_load.sas               # Sprint 1: Data profiling and loading
│   │   ├── step02_process_inclusion.sas              # Sprint 2: Find payment processes (deterministic PP1-PP5)
│   │   ├── step04_control_inclusion.sas              # Sprint 2: Find payment controls (deterministic P1-P5)
│   │   ├── step06_convergence_loop.sas               # Sprint 3: Convergence loop PP4/P2 iteration
│   │   ├── step07a_payment_nature_inheritance.sas    # Sprint 5: Payment Nature inheritance from processes
│   │   ├── step07c_sanity_check.sas                  # Sprint 5: Sanity check inherited vs LLM
│   │   ├── step10_bridge_table.sas                   # Sprint 6: Build bridge table
│   │   ├── step11a_coverage_heatmaps.sas             # Sprint 7: Coverage heatmaps (Nature/VS/Process by CT)
│   │   ├── step11b_strengthen_view.sas               # Sprint 7: Strengthen view (automated, preventative, exemplars)
│   │   ├── step11c_duplicates_best_of_breed.sas      # Sprint 8: Duplicate clusters and best-of-breed scoring
│   │   ├── step11d_governance_misclassified.sas      # Sprint 7: Governance controls wrongly tagged 7.4.x
│   │   ├── step11e_gap_register.sas                  # Sprint 8: Gap register (six gap types)
│   │   ├── step11f_juno_registration_backlog.sas     # Sprint 8: JUNO registration backlog (bespoke controls)
│   │   └── step11h_monitoring_baseline.sas           # Sprint 8: Monitoring baseline per CT
│   │
│   ├── python/
│   │   └── llm_orchestration/
│   │       ├── config.py                             # Sprint 1: LLM configuration (models, rate limits)
│   │       ├── test_llm_access.py                    # Sprint 1: Test Claude and Gemini API access
│   │       ├── compute_all_embeddings.py             # Sprint 2: Compute embeddings for full universes
│   │       ├── fuzzy_matcher.py                      # Sprint 2: Fuzzy matching for bespoke discovery
│   │       ├── audit_logger.py                       # Sprint 2: Audit trail logging functions
│   │       ├── prompt_executor.py                    # Sprint 2: LLM prompt execution with retry logic
│   │       ├── mechanism_similarity.py               # Sprint 7: Mechanism similarity calculator
│   │       ├── queue_resolver.py                     # Sprint 6: SME review queue resolution workflow
│   │       └── llm_caller.py                         # Sprint 1-9: Generic LLM API wrapper
│   │
│   └── sql/
│       └── teradata_passthrough/
│           ├── step01_juno_profile.sql               # Sprint 1: JUNO profiling query (BIP1ViewA.CONTROLS)
│           ├── step01_linkage_profile.sql            # Sprint 1: Linkage tables profiling queries
│           └── step04_control_inclusion.sql          # Sprint 2: Control inclusion queries (P1, P3, P4)
│
├── 03_data/
│   │
│   ├── raw/                                          # Source data extracts (CSV from upstream systems)
│   │   ├── holocentric_processes_YYYYMMDD.csv        # Holocentric process extract (L2/L3/Tasks)
│   │   └── .gitkeep                                  # Other raw files loaded directly from Teradata
│   │
│   ├── staging/                                      # SAS staging datasets (first load from raw/Teradata)
│   │   ├── juno_staging.sas7bdat                     # Sprint 1: JUNO controls loaded
│   │   ├── holocentric_staging.sas7bdat              # Sprint 1: Holocentric processes loaded
│   │   ├── linkage_ctrl_proc.sas7bdat                # Sprint 1: Control-Process linkage
│   │   ├── linkage_ctrl_oblig.sas7bdat               # Sprint 1: Control-Obligation linkage
│   │   ├── linkage_ctrl_incident.sas7bdat            # Sprint 1: Control-Incident linkage
│   │   ├── linkage_ctrl_issue.sas7bdat               # Sprint 1: Control-Issue linkage
│   │   ├── obligations_staging.sas7bdat              # Sprint 1: Obligations master
│   │   ├── incidents_staging.sas7bdat                # Sprint 1: Incidents master
│   │   └── issues_staging.sas7bdat                   # Sprint 1: Issues master
│   │
│   ├── intermediate/                                 # Work-in-progress outputs from each step
│   │   ├── payment_relevant_apps.sas7bdat            # Sprint 1: Payment-relevant applications list
│   │   ├── step02_payment_process_register.sas7bdat  # Sprint 2-3: Payment processes with PP1-PP6 paths fired
│   │   ├── step04_payment_control_register.sas7bdat  # Sprint 2-3: Payment controls with P1-P6 paths fired
│   │   ├── step05_bespoke_register.sas7bdat          # Sprint 4: Bespoke controls from Holocentric/incidents
│   │   ├── step07_control_to_nature_inherited.sas7bdat  # Sprint 5: Payment Nature inherited from processes
│   │   ├── step07_control_to_nature_all.sas7bdat     # Sprint 5: Payment Nature all sources (inherited + LLM)
│   │   ├── step07_sanity_check_results.csv           # Sprint 5: Sanity check disagreement analysis
│   │   ├── step08_ct_mapping.sas7bdat                # Sprint 6: Controls to CT1-CT28 mapping
│   │   ├── step09_lifecycle_placement.sas7bdat       # Sprint 6: Lifecycle stages A-G placement
│   │   ├── step10_bridge_table.sas7bdat              # Sprint 6: Master bridge table
│   │   ├── convergence_log.csv                       # Sprint 3: Pass-by-pass convergence tracking
│   │   └── incidents_summary.sas7bdat                # Sprint 6: Incident counts aggregated per control
│   │
│   ├── outputs/                                      # Final deliverables (Step 11 analytical outputs)
│   │   ├── 11a_coverage_heatmaps.xlsx                # Sprint 7: Three heatmaps (Nature/VS/Process by CT)
│   │   ├── 11b_strengthen_view.xlsx                  # Sprint 7: Strong controls (automated, preventative, exemplars)
│   │   ├── 11c_duplicates_best_of_breed.xlsx         # Sprint 8: Duplicate clusters with best-of-breed ranking
│   │   ├── 11d_governance_misclassified.xlsx         # Sprint 7: Governance controls wrongly tagged 7.4.x
│   │   ├── 11e_gap_register_consolidated.xlsx        # Sprint 8: Six gap types consolidated
│   │   ├── 11f_juno_registration_backlog.xlsx        # Sprint 8: Bespoke controls for JUNO registration
│   │   └── 11h_monitoring_baseline.xlsx              # Sprint 8: Monitoring baseline per CT with key controls
│   │
│   └── validation/                                   # Validation datasets (not in original spec but recommended)
│       ├── step02_validation_results.csv             # Sprint 3: Step 2 validation metrics
│       ├── step03_validation_results.csv             # Sprint 5: Step 3 validation metrics
│       ├── step05_validation_results.csv             # Sprint 4: Step 5 validation metrics
│       ├── step08_validation_results.csv             # Sprint 6: Step 8 validation metrics
│       └── step09_validation_results.csv             # Sprint 6: Step 9 validation metrics
│
├── 04_embeddings/
│   ├── holocentric_process_embeddings.npy            # Sprint 2: Process embeddings (3072-dim, ~1200 items)
│   ├── embedding_index_processes.csv                # Sprint 2: Index mapping row to l3_activity_id
│   ├── control_embeddings.npy                        # Sprint 2: Control embeddings (3072-dim, ~4500 items)
│   └── embedding_index_controls.csv                 # Sprint 2: Index mapping row to ctrl_id
│
├── 05_audit_trail/                                   # Full audit trail per LLM step
│   ├── audit_trail_step02_deterministic.csv         # Sprint 2: Step 2 deterministic paths audit
│   ├── audit_trail_step02_pp6.csv                   # Sprint 3: Step 2 PP6 embedding path audit
│   ├── audit_trail_step03.csv                       # Sprint 3: Step 3 Payment Nature on processes audit
│   ├── audit_trail_step04_deterministic.csv         # Sprint 2: Step 4 deterministic paths audit
│   ├── audit_trail_step04_p6.csv                    # Sprint 3: Step 4 P6 embedding path audit
│   ├── audit_trail_step05a.csv                      # Sprint 4: Step 5a Holocentric bespoke audit
│   ├── audit_trail_step05c.csv                      # Sprint 4: Step 5c incident remediation bespoke audit
│   ├── audit_trail_step07b.csv                      # Sprint 5: Step 7b Payment Nature controls fallback audit
│   ├── audit_trail_step08.csv                       # Sprint 6: Step 8 CT mapping audit
│   ├── audit_trail_step09.csv                       # Sprint 6: Step 9 lifecycle placement audit
│   └── validation_alerts.log                        # Sprint 7: Validation alerts (if agreement <70%)
│
├── 06_sme_review/
│   ├── step02_golden_set.xlsx                       # Sprint 3: Golden set for process inclusion (30 items)
│   ├── step02_pp6_queue.xlsx                        # Sprint 3: PP6 embedding candidates for SME review
│   ├── step03_golden_set.xlsx                       # Sprint 5: Golden set for Payment Nature on processes (30 items)
│   ├── step04_p6_queue.xlsx                         # Sprint 3: P6 embedding candidates for SME review
│   ├── step05_golden_set.xlsx                       # Sprint 4: Golden set for bespoke discovery (40 items)
│   ├── step07_unlinked_queue.xlsx                   # Sprint 5: Unlinked controls Payment Nature for SME review
│   ├── step08_golden_set.xlsx                       # Sprint 6: Golden set for CT mapping (50 items)
│   ├── step09_golden_set.xlsx                       # Sprint 6: Golden set for lifecycle placement (50 items, reuses step08)
│   ├── sme_queue_combined.xlsx                      # Sprint 6: Combined review queue (Low-conf + disagreements)
│   └── sme_decisions_consolidated.csv               # Sprint 6-8: SME decisions tracked (Approve/Revise/Reject)
│
└── README.md                                         # Project overview and folder structure guide
```

---

## Detailed Folder Descriptions

### 00_documentation/
**Purpose:** Profiling reports, data quality findings, and project documentation.

**Key Files:**
- **juno_profile_report.html** - Sprint 1 output showing total controls, Payments-tagged (7.4.x), effectiveness ratings, process links, orphan keys
- **holocentric_profile_report.html** - Sprint 1 output showing L2/L3/Task counts, description coverage, linkage to JUNO
- **linkage_profile_report.html** - Sprint 1 output showing linkage densities, incident severities, obligation sources
- **data_quality_findings.md** - Sprint 1 consolidated findings (high null rates, orphan keys, missing sources)
- **reusability_handover.docx** - Sprint 9 output documenting generic backbone vs Payments-specific overlay

---

### 01_prompts/
**Purpose:** LLM prompt templates with version control.

**Naming Convention:** `step0X_[description].txt`

**Key Files:**
- Each prompt file contains:
  - System Context (taxonomy definitions, business rules)
  - Task definition
  - Input format
  - Output JSON schema
  - Confidence rubric (High/Medium/Low criteria)
  - Examples (2-3 worked examples)

- **prompt_versions.json** - Tracks SHA-256 hash per prompt file, version number, last_updated timestamp, validation_status (Pending/Pass/Fail)

**Versioning Rule:** Any change to a prompt file triggers:
1. Recompute SHA-256 hash
2. Update prompt_versions.json
3. Set validation_status='Pending'
4. Block production runs until re-validated

---

### 02_code/sas/
**Purpose:** SAS scripts for data processing, profiling, joining, and bridge table assembly.

**Naming Convention:** `step0X_[description].sas`

**Key Files:**
- **step01_profile_and_load.sas** - Loads all source data via Teradata passthrough, generates profiling reports
- **step02_process_inclusion.sas** - Deterministic paths PP1-PP5, confidence rollup, payment process register
- **step04_control_inclusion.sas** - Deterministic paths P1-P5, taxonomy hygiene finding, payment control register
- **step06_convergence_loop.sas** - Iterates Steps 2 and 4 with PP4/P2 paths until stable, convergence log
- **step07a_payment_nature_inheritance.sas** - Joins controls to processes, inherits Payment Natures (UNION)
- **step10_bridge_table.sas** - Master analytical table join with indexes on 5 dimensions
- **step11X_*.sas** - Analytical output queries (coverage heatmaps, gaps, duplicates, strengthen, governance, monitoring)

---

### 02_code/python/llm_orchestration/
**Purpose:** Python scripts for LLM API calls, embeddings, fuzzy matching, audit logging.

**Key Files:**
- **config.py** - Configuration constants:
  ```python
  PRIMARY_MODEL = "claude-sonnet-4-20250514"
  INDEPENDENT_MODEL = "gemini-1.5-pro"
  MAX_TOKENS = 4000
  RATE_LIMIT_CALLS_PER_SEC = 10
  EMBEDDING_MODEL = "text-embedding-3-large"
  EMBEDDING_DIMENSIONS = 3072
  ```

- **test_llm_access.py** - Sprint 1 validation script testing Claude and Gemini API connectivity, JSON response parsing

- **compute_all_embeddings.py** - Sprint 2 batch processor:
  - Computes embeddings for FULL Holocentric process universe (~1,200 items)
  - Computes embeddings for FULL JUNO control universe (~4,500 items)
  - Saves to numpy .npy files with CSV indexes
  - Handles rate limiting, retry logic
  - Cost: ~$0.12 total for 5,700 items

- **fuzzy_matcher.py** - Sprint 2 fuzzy matching tool:
  - Function: `fuzzy_match(text1, text2, threshold=0.80)` returns (is_match, similarity)
  - Function: `find_best_match(candidate_text, juno_df, threshold=0.80)` returns best ctrl_id or None
  - Used in Step 5a (Holocentric bespoke) and 5c (incident remediation)

- **mechanism_similarity.py** - Sprint 7 mechanism similarity calculator:
  - Function: `compute_mechanism_similarity(ctrl_A, ctrl_B, embeddings_dict)` returns 0-1 score
  - Concatenates: description embedding (3072-dim) + automation (3-dim) + control_type (3-dim) = 3078-dim
  - Computes cosine similarity
  - CRITICAL RULE: if automation differs, NOT duplicate (defense-in-depth)

- **audit_logger.py** - Sprint 2 audit trail functions:
  - Function: `log_llm_call(step_id, prompt_hash, input_data, raw_response, parsed_result)`
  - Logs: prompt hash (SHA-256), model ID, timestamp, input hash, raw response, parsed result, confidence
  - Enables reproducibility: given audit record, reconstruct exact API call

- **prompt_executor.py** - Generic LLM caller with retry logic, rate limiting, audit logging

- **queue_resolver.py** - Sprint 6 SME review queue resolution:
  - Reads SME decisions from sme_queue_combined.xlsx
  - Propagates decisions back to source datasets (step02, step03, step05, step07, step08, step09)
  - Updates confidence levels and flags based on SME adjudication

---

### 02_code/sql/teradata_passthrough/
**Purpose:** SQL queries for Teradata passthrough in SAS.

**Schema:** All queries use `BIP1ViewA` schema.

**Key Files:**
- **step01_juno_profile.sql** - SELECT from BIP1ViewA.CONTROLS WHERE status IN ('Active','Under Review')
- **step01_linkage_profile.sql** - SELECT from BIP1ViewA.CONTROL_PROCESS_LINK, CONTROL_OBLIGATION_LINK, CONTROL_INCIDENT_LINK, CONTROL_ISSUE_LINK, OBLIGATIONS, INCIDENTS, ISSUES
- **step04_control_inclusion.sql** - Queries for P1 (taxonomy), P3 (obligation), P4 (incident/issue) paths

---

### 03_data/raw/
**Purpose:** Source data extracts in CSV format.

**Files:**
- **holocentric_processes_YYYYMMDD.csv** - Holocentric extract with L2/L3/Task hierarchy, value streams, applications
- Other sources (JUNO, linkages, obligations, incidents, issues) loaded directly from Teradata via passthrough (no CSV files)

---

### 03_data/staging/
**Purpose:** First-load SAS datasets from raw/Teradata.

**Naming Convention:** `[source]_staging.sas7bdat`

**Files:** All loaded in Sprint 1 step01_profile_and_load.sas

---

### 03_data/intermediate/
**Purpose:** Work-in-progress outputs from each step.

**Naming Convention:** `step0X_[description].sas7bdat` or `.csv`

**Key Files:**
- **step02_payment_process_register.sas7bdat** - Schema: l3_activity_id, l3_activity_name, l2_process_id, l2_process_name, value_stream, applications, pp1_fired, pp2_fired, pp3_fired, pp4_fired (Sprint 3), pp5_fired, pp6_fired (Sprint 3), primary_path, confidence_rollup, included
  
- **step04_payment_control_register.sas7bdat** - Schema: ctrl_id, ctrl_name, ctrl_description, risk_library_id, control_type, automation_level, owner_business_unit, overall_effectiveness, p1_fired, p2_fired (Sprint 3), p3_fired, p4_fired, p5_fired, p6_fired (Sprint 3), primary_path, confidence_rollup, included

- **step05_bespoke_register.sas7bdat** - Schema: bespoke_id, source_type (Holocentric-Task, remediation), source_process_id, source_task_id, control_description, confidence

- **step10_bridge_table.sas7bdat** - Schema: ctrl_id, ctrl_name, control_category, control_nature, automation_level, overall_effectiveness, ct_outcome, payment_nature, lifecycle_stage, process_id, process_name, value_stream, linkage_type, incident_count, incident_severity_max, framework_alignment

- **convergence_log.csv** - Schema: pass_number, new_processes_added, new_controls_added, total_processes_high_medium, total_controls_high_medium, converged (true/false)

---

### 03_data/outputs/
**Purpose:** Final deliverables (Step 11 analytical outputs).

**Naming Convention:** `11X_[description].xlsx`

**Key Files:**

**11a_coverage_heatmaps.xlsx** - Three tabs:
- Tab 1: Payment Nature by CT (5 rows x 28 columns)
- Tab 2: Value Stream by CT (8-12 rows x 28 columns)
- Tab 3: Process by CT (200-500 rows x 28 columns)
- Each cell: Structural coverage, Effective coverage, Failing coverage
- Conditional formatting: GREEN if Effective>3 AND Failing=0, AMBER if Effective 1-2 OR Failing>0, RED if Structural=0

**11b_strengthen_view.xlsx** - Three tabs:
- Tab 1: Automated controls in High-risk Payment Natures (C to C, I to C)
- Tab 2: Preventative controls at critical lifecycle stages (A, B, C)
- Tab 3: Exemplar candidates (Effective + no incidents 24 months)

**11c_duplicates_best_of_breed.xlsx** - Schema per cluster:
- cluster_id, ctrl_id, ctrl_name, ct_outcome, payment_nature, lifecycle_stage, effectiveness_score, automation_score, incident_free_score, design_clarity_score, coverage_breadth_score, total_score, best_of_breed_flag (Y/N), recommendation (Retire/Merge into best-of-breed)

**11d_governance_misclassified.xlsx** - Schema:
- ctrl_id, ctrl_name, ctrl_description, owner_business_unit, current_tag (7.4.x), framework_alignment (Governance), framework_rationale, recommendation (Remove from 7.4.x, re-home under Governance Risk)

**11e_gap_register_consolidated.xlsx** - Six tabs (one per gap type):
- Type A: No coverage (Payment Nature x CT combinations with zero controls)
- Type B: Insufficient rating (all controls Requires Improvement or Ineffective)
- Type C: Unverified (overall_effectiveness IS NULL)
- Type D: Failing (Effective but has incidents in 12 months)
- Type E: Misplacement (structural_stage != llm_stage)
- Type F: Process gap (processes with no linked controls)
- Schema: gap_type, payment_nature, ct_outcome, process_id, nature_risk, ct_criticality, overall_priority (P0/P1/P2), recommended_action

**11f_juno_registration_backlog.xlsx** - Schema:
- bespoke_id, source_type, control_description, source_process_id, ct_outcome, payment_nature, lifecycle_stage, applications, recommended_owner, priority (P0/P1/P2)

**11h_monitoring_baseline.xlsx** - One tab per CT (28 tabs):
- Schema per CT: key_control_id, key_control_name, monitoring_intensity (Critical/High/Medium/Low), existing_evidence (latest_assessment_date, KRI_exists, automated_monitoring), gap_flag (Y/N), recommendation

---

### 04_embeddings/
**Purpose:** Precomputed embeddings for reuse across steps.

**Files:**
- **holocentric_process_embeddings.npy** - Numpy array shape (1200, 3072) containing embeddings for all Holocentric processes (L3 description + concatenated Task descriptions)
- **embedding_index_processes.csv** - Schema: row_index, l3_activity_id (maps row number to process ID)
- **control_embeddings.npy** - Numpy array shape (4500, 3072) containing embeddings for all JUNO controls (ctrl_description)
- **embedding_index_controls.csv** - Schema: row_index, ctrl_id (maps row number to control ID)

**Computed in:** Sprint 2 (compute_all_embeddings.py)
**Reused in:** Sprint 3 PP6, Sprint 3 P6, Sprint 7 mechanism similarity

**Loading Example:**
```python
import numpy as np
import pandas as pd

embeddings = np.load('04_embeddings/holocentric_process_embeddings.npy')
index = pd.read_csv('04_embeddings/embedding_index_processes.csv')
embeddings_dict = {row['l3_activity_id']: embeddings[row['row_index']] for _, row in index.iterrows()}
```

---

### 05_audit_trail/
**Purpose:** Complete audit trail for every LLM decision.

**Naming Convention:** `audit_trail_step0X[_method].csv`

**Schema (all audit files):**
- step_id (e.g., 'step03')
- item_id (process_id or ctrl_id)
- prompt_hash (SHA-256 of prompt file content)
- model_id (e.g., 'claude-sonnet-4-20250514')
- run_timestamp (ISO format)
- input_hash (SHA-256 of concatenated input fields)
- raw_response (full JSON from LLM)
- parsed_result (extracted fields)
- confidence (High/Medium/Low)
- pass_number (for convergence loop steps)

**Purpose:** Enables reproducibility - given audit record, can reconstruct exact API call and verify result.

---

### 06_sme_review/
**Purpose:** Golden sets for validation and review queues for Low-confidence items.

**Golden Set Files:**

**step02_golden_set.xlsx** - Sprint 3, 30 processes:
- Schema: l3_activity_id, process_name, process_description, sme_label (is_payment_process Y/N), sme_rationale, independent_llm_label, independent_rationale, primary_llm_label, primary_confidence, agreement (SME vs Independent), validation_result (Pass/Fail)
- Stratification: 10 clear payment, 10 clear non-payment, 10 edge cases

**step03_golden_set.xlsx** - Sprint 5, 30 processes:
- Schema: l3_activity_id, process_name, sme_applicable_natures (comma-separated), sme_rationale, independent_applicable_natures, primary_applicable_natures, primary_confidence, precision_per_nature, recall_per_nature, validation_result
- Stratification: 10 single-nature, 10 multi-nature, 10 ambiguous

**step05_golden_set.xlsx** - Sprint 4, 40 items:
- Schema: item_id, source_type (Holocentric-Task / incident-remediation), item_text, sme_label (is_control Y/N), sme_rationale, independent_label, primary_label, primary_confidence, precision_per_stream, validation_result
- Stratification: 20 Holocentric Tasks (10 control, 10 not), 20 remediation texts (10 control, 10 not)

**step08_golden_set.xlsx** - Sprint 6, 50 controls:
- Schema: ctrl_id, ctrl_description, sme_ct_outcomes (comma-separated), sme_rationale, independent_ct_outcomes, primary_ct_outcomes, primary_confidence, precision_per_ct, recall_per_ct, validation_result
- Stratification: All 28 CTs covered, mix of automated/manual, JUNO/bespoke

**step09_golden_set.xlsx** - Sprint 6, 50 controls (reuses step08):
- Schema: ctrl_id, ctrl_description, sme_lifecycle_stage (A-G), sme_rationale, independent_stage, primary_stage, primary_confidence, exact_match (Y/N), validation_result
- Same 50 controls from step08

**Review Queue Files:**

**sme_queue_combined.xlsx** - Sprint 6-8 consolidated queue:
- Schema: item_id, source_step, item_description, primary_result, independent_result, disagreement_type (Low-confidence / Disagreement), nature_risk (High/Med/Low), ct_criticality (P0/P1/P2), priority_score, assigned_sme, status (Pending/Reviewed), sme_decision (Approve Primary / Approve Independent / Revise Custom / Reject), sme_rationale, resolution_date
- Populated by: Any LLM step where confidence=Low OR primary != independent
- Priority: (nature_risk * 3) + (ct_criticality * 2)

**sme_decisions_consolidated.csv** - Tracks all SME decisions:
- Schema: decision_id, item_id, source_step, sme_decision, sme_rationale, resolution_date, resolved_by_sme

---

### README.md
**Purpose:** Project overview, folder structure guide, quick start instructions.

**Contents:**
- Project overview and objectives
- Folder structure hierarchy
- File naming conventions
- Quick start guide (how to run each sprint)
- Dependencies and prerequisites
- Contact information

---

## File Naming Conventions Summary

| Type | Convention | Example |
|------|-----------|---------|
| Profiling reports | `[source]_profile_report.html` | juno_profile_report.html |
| Prompts | `step0X_[description].txt` | step03_payment_nature_processes.txt |
| SAS scripts | `step0X_[description].sas` | step02_process_inclusion.sas |
| Python scripts | `[function]_[description].py` | compute_all_embeddings.py |
| SQL queries | `step0X_[description].sql` | step01_juno_profile.sql |
| Staging data | `[source]_staging.sas7bdat` | juno_staging.sas7bdat |
| Intermediate data | `step0X_[description].sas7bdat` | step02_payment_process_register.sas7bdat |
| Output deliverables | `11X_[description].xlsx` | 11a_coverage_heatmaps.xlsx |
| Embeddings | `[source]_embeddings.npy` | holocentric_process_embeddings.npy |
| Audit trail | `audit_trail_step0X[_method].csv` | audit_trail_step03.csv |
| Golden sets | `step0X_golden_set.xlsx` | step08_golden_set.xlsx |
| Review queues | `step0X_[type]_queue.xlsx` | step02_pp6_queue.xlsx |

---

## Sprint-to-Folder Mapping

| Sprint | Folders Used | Key Files Created |
|--------|-------------|-------------------|
| Sprint 1 | 00_documentation/, 03_data/raw/, 03_data/staging/, 02_code/sql/, 02_code/sas/, 02_code/python/ | juno_profile_report.html, holocentric_profile_report.html, linkage_profile_report.html, all staging datasets, test_llm_access.py |
| Sprint 2 | 01_prompts/, 02_code/python/, 02_code/sas/, 03_data/intermediate/, 04_embeddings/, 05_audit_trail/ | All prompt files, compute_all_embeddings.py, fuzzy_matcher.py, step02/04 registers, process/control embeddings, audit trail files |
| Sprint 3 | 02_code/sas/, 03_data/intermediate/, 05_audit_trail/, 06_sme_review/ | convergence_log.csv, step02/04 registers updated with PP4/P2, step02/03 golden sets, PP6/P6 queues |
| Sprint 4 | 02_code/python/, 03_data/intermediate/, 05_audit_trail/, 06_sme_review/ | step05_bespoke_register, step05 golden set, audit trail step05a/c |
| Sprint 5 | 02_code/sas/, 03_data/intermediate/, 05_audit_trail/, 06_sme_review/ | step07 control-to-nature files, sanity check results, step03 golden set, step07 unlinked queue |
| Sprint 6 | 02_code/sas/, 03_data/intermediate/, 05_audit_trail/, 06_sme_review/ | step08 CT mapping, step09 lifecycle placement, step10 bridge table, sme_queue_combined, step08/09 golden sets |
| Sprint 7 | 02_code/sas/, 02_code/python/, 03_data/outputs/ | 11a coverage heatmaps, 11b strengthen view, 11d governance misclassified, mechanism_similarity.py |
| Sprint 8 | 02_code/sas/, 03_data/outputs/ | 11c duplicates, 11e gap register, 11f JUNO backlog, 11h monitoring baseline |
| Sprint 9 | 00_documentation/ | reusability_handover.docx |

---

## Data Flow Summary

```
Sprint 1: BIP1ViewA (Teradata) → 03_data/staging/ → 00_documentation/profiling reports

Sprint 2: 03_data/staging/ → 02_code/sas/python → 03_data/intermediate/step02,04 registers
                                                 → 04_embeddings/ (ALL computed here)
                                                 → 05_audit_trail/

Sprint 3: 03_data/intermediate/step02,04 → 02_code/sas (convergence) → step02,04 updated
                                                                      → 06_sme_review/golden sets

Sprint 4: 03_data/staging/issues,incidents → 02_code/python/LLM → 03_data/intermediate/step05 bespoke
                                                                 → 06_sme_review/golden set

Sprint 5: 03_data/intermediate/step02,04,07 → 02_code/sas → step07 Payment Nature all sources
                                                          → 06_sme_review/golden set

Sprint 6: 03_data/intermediate/step02-09 → 02_code/sas → step10 bridge table
                                                       → 06_sme_review/sme_queue_combined

Sprint 7-8: 03_data/intermediate/step10 bridge → 02_code/sas queries → 03_data/outputs/11X analytical outputs

Sprint 9: All outputs → 00_documentation/reusability_handover.docx
```

---

## Total File Count Estimate

- **00_documentation:** 6 files
- **01_prompts:** 9 files
- **02_code/sas:** 13 files
- **02_code/python:** 9 files
- **02_code/sql:** 3 files
- **03_data/raw:** 1 file (+ Teradata direct loads)
- **03_data/staging:** 9 files
- **03_data/intermediate:** 12 files
- **03_data/outputs:** 7 files
- **03_data/validation:** 5 files
- **04_embeddings:** 4 files
- **05_audit_trail:** 11 files
- **06_sme_review:** 11 files

**Total: ~100 files across 18 folders**

---

**END OF FOLDER STRUCTURE DOCUMENTATION**
