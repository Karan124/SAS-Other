# Payment Control Linkage Methodology
## Payments Controls PoC — Control Coverage Analysis

---

## 1. Objective

Determine which JUNO payment controls are linked to classified Holocentric payment
processes, and assess payment control coverage across:

- Payment Category (5 valid values)
- Payment Lifecycle Stage (6 valid stages)
- Gold Control Outcome (CT1–CT28)

---

## 2. Input Data Model

### File 1 — JUNO Payment Controls (`juno_payment_controls_gold.xlsx`)
- **Grain:** One row per JUNO control
- **Population:** 708 pre-defined payment controls
- **Key fields:** `Control_ID`, `CTRL_NAME`, `CTRL_DESC`, `gold_control` (CT1–CT28)
- **Assumption:** All 708 controls are in scope. CTRL_STUS is not used as a filter
  at this stage; the payment control population is pre-defined.

### File 2 — JUNO–Holocentric Linkage Inventory (`juno_holo_deterministic_linkage.xlsx`)
- **Grain:** One row per JUNO control × Holocentric activity/process link
- **Key fields:** `CTRL_ID`, `l3_activity_uuid`, `l2_process_uuid`
- **Nature:** Many-to-many. One control can link to multiple Holocentric activities.
- **Row types:**
  - **L3 rows:** `l3_activity_uuid` is populated (primary linkage type)
  - **L2-only rows:** `l3_activity_uuid` is null but `l2_process_uuid` is populated
- **Join to File 1:** `CTRL_ID` = `Control_ID`

### File 3 — Holocentric Payment Process Inventory (`holocentric_payment_processes.xlsx`)
- **Grain:** One row per Holocentric L3 activity
- **Population:** Payment-related processes only (output of category + lifecycle runs)
- **Key fields:** `l3_process_UUID`, `l2_process_UUID`, `process_category`,
  `process_lifecycle_stage`
- **Join to File 2 (L3):** `l3_process_UUID` = `l3_activity_uuid`
- **Join to File 2 (L2 fallback):** `l2_process_UUID` = `l2_process_uuid`

---

## 3. Data Normalisation

Before any joins are performed, the following normalisation is applied:

### 3.1 Lifecycle Stage Variants
The following observed variant strings are mapped to the canonical form:

| Variant (observed)              | Canonical form                     |
|--------------------------------|------------------------------------|
| Posting, Accounting & Detection | Posting & Accounting, Detection    |
| Posting & Accounting & Detection| Posting & Accounting, Detection    |

### 3.2 Payment Category Normalisation
Values of `0`, empty strings, null, or unrecognised strings in `process_category`
are mapped to `Unclassified / Missing`.

### 3.3 UUID Normalisation
All UUID fields are stripped of whitespace and lowercased before joining to avoid
case-sensitivity mismatches between systems.

---

## 4. Linkage Methodology

### 4.1 Overview

The linkage uses a **three-table join** with an **L3-first, L2-fallback hierarchy**:

```
File 1 (Controls)
    ↓  join on Control_ID = CTRL_ID
File 2 (Linkage) — L3 rows only
    ↓  inner join on l3_activity_uuid = l3_process_UUID
File 3 (Payment Processes)
    → Payment linkage detail (one row per control-process pair)
```

### 4.2 L3 Primary Linkage

**Join chain:**
```
Controls.Control_ID
  = Linkage.CTRL_ID (L3 rows: l3_activity_uuid is not null)
  INNER JOIN Processes.l3_process_UUID on Linkage.l3_activity_uuid
```

**Result:** One row per (Control_ID, l3_process_UUID) pair.
Deduplicated — if a control links to the same L3 activity via multiple linkage
rows, only one pair is retained.

**This is the authoritative linkage layer.** It directly connects a JUNO control
to a classified payment process at the most granular level available, and inherits
`process_category` and `process_lifecycle_stage` directly.

### 4.3 L2 Fallback Linkage

**Applied only to controls with zero L3 payment links.**

Some rows in the linkage file contain `l2_process_uuid` but no `l3_activity_uuid`.
For controls that are not matched via L3, these rows are used as a fallback:

```
Linkage.l2_process_uuid (L2-only rows)
  INNER JOIN Processes.l2_process_UUID
```

**Critical design decision:** The L2 fallback does NOT expand to individual L3
rows. One L2 process can have many L3 children, and expanding would cause
misleading over-attribution of categories and lifecycle stages. Instead, the L2
fallback aggregates all distinct categories and lifecycle stages from the L3
children under the matched L2 process and assigns them at control level.

**L2 fallback output:**
- `link_type = L2_fallback`
- `confidence = Low`
- Pipe-delimited lists of categories and stages (no primary assigned at row level)
- Not used for individual control-to-process attribution

### 4.4 Why L2 Expansion Was Rejected

Joining directly on `l2_process_uuid = l2_process_UUID` without filtering
first inflates rows to 36,000+ because each L2 process has multiple L3 children,
and each child activity creates a separate row for each linked control. This causes
every control linked at L2 to appear linked to every child activity under that L2
process, regardless of whether those specific activities are relevant to that
control.

---

## 5. Population Definitions

After the three-table join, all 708 controls are classified into four mutually
exclusive populations. The sum of all populations must equal 708.

### Population A1 — Payment-linked, pure
**Criteria:** At least one L3 payment link exists AND the total number of L3
Holocentric links equals the number of L3 payment links (all Holo links are
to payment processes).

### Population A2 — Payment-linked, mixed
**Criteria:** At least one L3 payment link exists BUT some L3 Holocentric links
do not match any payment process (the control also links to non-payment Holocentric
activities).

### Population B — Holocentric-linked, no payment match
**Criteria:** The control has rows in the linkage file (L3 or L2) but none of
those rows join to a payment process in File 3.
This may indicate the control governs non-payment Holocentric processes or the
linkage inventory requires review.

### Population C — No Holocentric linkage
**Criteria:** The control has no rows in the linkage file at all.
These controls have no operational linkage to any Holocentric process.
These represent genuine coverage gaps from a linkage perspective.

---

## 6. Multi-Process Classification Logic

### 6.1 Retain All Linkages
The detail table retains every control-process pair. A control linked to six
payment processes has six rows. This is the analytical foundation and must not
be collapsed before analysis.

### 6.2 Control Profile Aggregation
For each control in Population A, the following are computed from the detail table:

| Field                   | Derivation                                              |
|-------------------------|---------------------------------------------------------|
| `payment_categories`    | Pipe-delimited sorted list of unique categories         |
| `lifecycle_stages`      | Pipe-delimited sorted list of unique lifecycle stages   |
| `primary_category`      | Most frequent category (ties broken alphabetically)     |
| `primary_lifecycle_stage` | Most frequent stage (ties broken alphabetically)      |
| `payment_process_count` | Count of distinct l3_process_UUIDs linked               |
| `category_count`        | Count of distinct categories                            |
| `stage_count`           | Count of distinct lifecycle stages                      |
| `is_multi_category`     | True if category_count > 1                              |
| `is_multi_stage`        | True if stage_count > 1                                 |

**A control that covers multiple categories is correctly recorded as multi-category.
This is not an error — it reflects the genuine breadth of that control.**

### 6.3 Primary Category/Stage Selection Logic
Where multiple values appear with equal frequency, ties are broken alphabetically
to ensure deterministic, reproducible results. The primary values are provided
for convenience but should not be used as the sole basis for coverage analysis —
the full sets should always be considered.

---

## 7. Coverage Analysis

Three coverage matrices are produced, each showing unique control counts per cell.

### Matrix 1 — Payment Category × Lifecycle Stage
Rows: 5 valid categories (+ Unclassified / Missing)
Columns: 6 valid lifecycle stages
Cells: count of unique controls covering that category-stage combination

### Matrix 2 — Gold Control (CT1–CT28) × Payment Category
Rows: CT1–CT28
Columns: 5 valid categories
Cells: count of unique controls for that CT covering that category

### Matrix 3 — Gold Control (CT1–CT28) × Lifecycle Stage
Rows: CT1–CT28
Columns: 6 valid lifecycle stages
Cells: count of unique controls for that CT covering that stage

### 7.1 Gap Identification
A gap is defined as any valid combination with a cell count of zero:
- **Category × Stage gap:** No controls cover that category-stage combination
- **CT × Category gap:** A gold control has no controls covering a category
- **CT × Stage gap:** A gold control has no controls covering a lifecycle stage

Total possible combinations:
- Category × Stage: 5 × 6 = 30
- CT × Category: 28 × 5 = 140
- CT × Stage: 28 × 6 = 168

---

## 8. Output Datasets

| Sheet                    | Grain                          | Purpose                              |
|--------------------------|--------------------------------|--------------------------------------|
| `summary`                | Metrics                        | Key counts and coverage at a glance  |
| `ctrl_population`        | One row per control (all 708)  | Population classification A1/A2/B/C  |
| `ctrl_payment_profile`   | One row per control (A + L2FB) | Aggregated categories, stages, counts|
| `ctrl_payment_detail`    | One row per control-process pair | Full L3 payment linkage detail     |
| `ctrl_l2_fallback`       | One row per control (L2 only)  | L2 fallback summary (Low confidence) |
| `coverage_cat_x_stage`   | Category × Stage matrix        | Control counts per cell              |
| `coverage_ct_x_category` | CT × Category matrix           | Control counts per cell              |
| `coverage_ct_x_stage`    | CT × Stage matrix              | Control counts per cell              |
| `gap_analysis`           | One row per gap identified     | Zero-coverage combinations           |
| `validation_checks`      | One row per check              | Automated integrity validation       |

---

## 9. Validation Framework

The following checks are run automatically and results written to the
`validation_checks` sheet.

| Check | Expected | Failure indicates |
|-------|----------|-------------------|
| Total controls = 708 | 708 | File 1 has incorrect population |
| Population A1+A2+B+C = 708 | 708 | Classification logic error |
| No duplicate control-process pairs | 0 | Deduplication failed |
| All categories are valid | None invalid | Normalisation incomplete |
| All lifecycle stages are valid | None invalid | Normalisation incomplete |
| Stage normalisation complete | 0 variants remaining | Variant mapping incomplete |
| All gold controls are CT1-CT28 | None invalid | Data quality issue in File 1 |
| L2 fallback controls have zero L3 payment links | 0 overlap | Logic error in fallback filter |

---

## 10. Risks and Assumptions

### Assumptions

| # | Assumption |
|---|-----------|
| 1 | All 708 controls in File 1 are in scope (CTRL_STUS not used as filter) |
| 2 | `l3_activity_uuid` in File 2 and `l3_process_UUID` in File 3 represent the same entity despite different naming conventions |
| 3 | `process_category` and `process_lifecycle_stage` in File 3 reflect the latest LLM classification run output |
| 4 | A control can legitimately govern multiple payment categories and lifecycle stages |
| 5 | The `gold_control` field in File 1 contains exactly one CT value per control |

### Risks

| # | Risk | Mitigation |
|---|------|-----------|
| 1 | CTRL_STUS may include inactive controls that should be excluded | Confirm active status taxonomy; re-run with filter if required |
| 2 | L2-only linkage rows may overstate coverage due to L2 expansion | L2 fallback is aggregated and flagged Low confidence; not used for attribution |
| 3 | Some controls link to many processes, inflating average linkage counts | Multi-process profile is preserved in full; primary values are tie-broken deterministically |
| 4 | Unclassified categories in File 3 reduce coverage accuracy | Unclassified is tracked separately and visible in matrices |
| 5 | UUID case/whitespace inconsistencies between files may cause missed joins | All UUIDs are lowercased and stripped before joining |

---

## 11. Key Design Decisions

**Why retain all linkages rather than selecting a primary process?**
The end objective is coverage analysis — identifying gaps. Discarding linkages
would artificially reduce coverage counts and potentially hide genuine gaps. The
control profile layer provides primary values for convenience, but the full
multi-process linkage is preserved in the detail table.

**Why is L2 fallback aggregated rather than expanded?**
Expanding L2 links to individual L3 children causes every control linked at L2
to appear linked to every child activity under that L2 process. This inflates
coverage counts and misrepresents individual control-to-process relationships.
Aggregation at control level with Low confidence is the defensible alternative.

**Why is primary category/stage determined by frequency (mode)?**
Frequency-based selection reflects the dominant payment context for that control
across its linked processes. It is deterministic, reproducible, and does not
require SME intervention for every control. Where ties occur, alphabetical
ordering ensures consistency.

---

*End of methodology document.*
