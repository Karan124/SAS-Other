"""
build_consolidated_linkage_by_product.py
─────────────────────────────────────────────────────────────────────────────
Payments Controls PoC — Consolidated Control-Process-Product Linkage Table

Combines:
  - Deterministic linkages (380 controls, from linked_payment_detail)
  - Multi-signal LLM linkages (136 controls, from linkage_recommendations)

Expands by product — one row per control × process × product.
All control columns included for completeness.

Inputs:
  payment_control_linkage_analysis.xlsx  linked_payment_detail sheet
  multi_signal/linkage_recommendations.xlsx  linkage_recommendations sheet
  juno_payment_controls_gold.xlsx            all control columns
  holocentric_payment_processes.xlsx         process metadata + final_product

Output:
  phase1e_outputs/consolidated_linkage_by_product.xlsx

Run:
  python build_consolidated_linkage_by_product.py
"""

from pathlib import Path
import pandas as pd

# ─────────────────────────────────────────────────────────────────────────────
#  CONFIG
# ─────────────────────────────────────────────────────────────────────────────

BASE_DIR = Path(r"C:\Users\m061400\ai-test\big_table")

FILE_CONTROLS    = BASE_DIR / "juno_payment_controls_gold.xlsx"
FILE_PROCESSES   = BASE_DIR / "holocentric_payment_processes.xlsx"
FILE_DET_LINKAGE = BASE_DIR / "phase1e_outputs" / "payment_control_linkage_analysis.xlsx"
FILE_LLM_LINKAGE = BASE_DIR / "multi_signal" / "linkage_recommendations.xlsx"
OUTPUT_FILE      = BASE_DIR / "phase1e_outputs" / "consolidated_linkage_by_product.xlsx"

CT_TITLES = {
    "CT1":"Validation of Human-Entered Data at Input",
    "CT2":"Payment processing error detection",
    "CT3":"Early Identification of Duplications and Processing Errors",
    "CT4":"Payment processing interface and batch error resolution",
    "CT5":"Incident response","CT6":"Master/Reference data input validation",
    "CT7":"Service provider ongoing review","CT8":"Service provider onboarding",
    "CT9":"Change management testing","CT10":"Critical service chain mapping",
    "CT11":"Rollback plans","CT12":"System recovery capability",
    "CT13":"Business continuity plan","CT14":"Logging and monitoring",
    "CT15":"Secure IT Design","CT16":"Vulnerability Management",
    "CT17":"Access - Provision / deprovision","CT18":"Access - Monitoring",
    "CT19":"Patch Management","CT20":"Physical Security Controls",
    "CT21":"Access - Privileged users",
    "CT22":"Mistaken Internet Payment Reports",
    "CT23":"Provision of Confirmations and Notifications",
    "CT24":"Treatment of Unauthorised and Disputed Transactions",
    "CT25":"Provision of Confirmations and Notifications",
    "CT26":"Regulatory horizon scanning","CT27":"Records retention",
    "CT28":"Crisis Management planning and testing",
}

# ─────────────────────────────────────────────────────────────────────────────
#  HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def clean_cols(df):
    df.columns = df.columns.str.strip()
    return df

def norm_uuid(val):
    if pd.isna(val) or str(val).strip() == "":
        return None
    return str(val).strip().lower()

def norm_gold_ctrl(val):
    if pd.isna(val) or str(val).strip() == "":
        return None
    s = str(val).strip()
    if s.upper().startswith("CT"):
        try:
            n = int(s[2:])
            return f"CT{n}" if 1 <= n <= 28 else None
        except ValueError:
            return None
    try:
        n = int(float(s))
        return f"CT{n}" if 1 <= n <= 28 else None
    except (ValueError, TypeError):
        return None

# ─────────────────────────────────────────────────────────────────────────────
#  LOAD
# ─────────────────────────────────────────────────────────────────────────────

print("=" * 70)
print("  Consolidated Control-Process-Product Linkage Table")
print("=" * 70)

# ── Controls: all columns ──────────────────────────────────────────────────
print("\n  Loading controls (all columns)...")
controls = clean_cols(pd.read_excel(FILE_CONTROLS, dtype=str, engine="openpyxl"))
controls["Control_ID"]        = controls["Control_ID"].str.strip()
controls["gold_control_code"] = controls["gold_control"].apply(norm_gold_ctrl)
controls["gold_control_title"]= controls["gold_control_code"].map(CT_TITLES)
controls = controls.drop_duplicates(subset=["Control_ID"], keep="first")
print(f"  Controls loaded     : {len(controls):,}")

# ── Processes: key columns + final_product ─────────────────────────────────
print("  Loading processes and products...")
procs = clean_cols(pd.read_excel(FILE_PROCESSES, dtype=str, engine="openpyxl"))
procs["l3_process_UUID"] = procs["l3_process_UUID"].apply(norm_uuid)
procs = procs.dropna(subset=["l3_process_UUID"])
procs = procs.drop_duplicates(subset=["l3_process_UUID"])

# Columns to carry from the process file into the consolidated table
proc_cols = [c for c in [
    "l3_process_UUID",
    "l3_activity_id",
    "l3_activity_name",
    "l3_activity_description",
    "l3_activity_channels",
    "l3_activity_customer_segments",
    "l2_process_id",
    "l2_process_name",
    "l2_process_description",
    "process_category",
    "process_lifecycle_stage",
    "value_stream_name",
    "vcm_library_name",
    "payment_rationale",
    "alphabet_app",
    "final_product",
] if c in procs.columns]

procs_slim = procs[proc_cols].copy()
has_final_product = "final_product" in procs_slim.columns
print(f"  Processes loaded    : {len(procs_slim):,}")
print(f"  final_product col   : {'found' if has_final_product else 'NOT FOUND — check file'}")

# ── Deterministic linkages ──────────────────────────────────────────────────
print("  Loading deterministic linkages (linked_payment_detail)...")
det = clean_cols(pd.read_excel(
    FILE_DET_LINKAGE, sheet_name="linked_payment_detail",
    dtype=str, engine="openpyxl"
))
det["l3_process_UUID"] = det["l3_process_UUID"].apply(norm_uuid)
det["Control_ID"]      = det["Control_ID"].str.strip()

det_pairs = (
    det[["Control_ID","l3_process_UUID","link_type"]]
    .dropna(subset=["Control_ID","l3_process_UUID"])
    .drop_duplicates(subset=["Control_ID","l3_process_UUID"])
    .copy()
)
det_pairs["linkage_source"] = "Deterministic"
det_pairs["confidence"]     = det_pairs["link_type"].map({
    "L3 direct":                   "Deterministic — L3 direct",
    "L2 fallback (Low confidence)":"Deterministic — L2 fallback",
}).fillna("Deterministic")
det_pairs["primary_signal"] = det_pairs["link_type"]
det_pairs["llm_rationale"]  = ""
det_pairs["requires_sme_review"] = False
print(f"  Deterministic pairs : {len(det_pairs):,}  "
      f"({det_pairs['Control_ID'].nunique():,} controls)")

# ── LLM linkages ───────────────────────────────────────────────────────────
print("  Loading LLM linkages (linkage_recommendations)...")
llm = clean_cols(pd.read_excel(
    FILE_LLM_LINKAGE, sheet_name="linkage_recommendations",
    dtype=str, engine="openpyxl"
))
llm["l3_process_UUID"] = llm["l3_process_UUID"].apply(norm_uuid)
llm["Control_ID"]      = llm["Control_ID"].str.strip()

llm_pairs = (
    llm[["Control_ID","l3_process_UUID","confidence",
         "primary_signal","rationale"]]
    .dropna(subset=["Control_ID","l3_process_UUID"])
    .drop_duplicates(subset=["Control_ID","l3_process_UUID"])
    .copy()
)
llm_pairs = llm_pairs.rename(columns={"rationale":"llm_rationale"})
llm_pairs["linkage_source"]      = "Multi-signal LLM"
llm_pairs["link_type"]           = "LLM-recommended"
llm_pairs["requires_sme_review"] = True
print(f"  LLM pairs           : {len(llm_pairs):,}  "
      f"({llm_pairs['Control_ID'].nunique():,} controls)")

# ─────────────────────────────────────────────────────────────────────────────
#  COMBINE LINKAGES
# ─────────────────────────────────────────────────────────────────────────────

print("\n  Combining linkages...")

# Standardise columns before concat
shared_cols = [
    "Control_ID","l3_process_UUID","linkage_source",
    "confidence","link_type","primary_signal",
    "llm_rationale","requires_sme_review",
]
det_std = det_pairs[[c for c in shared_cols if c in det_pairs.columns]].copy()
llm_std = llm_pairs[[c for c in shared_cols if c in llm_pairs.columns]].copy()

combined = pd.concat([det_std, llm_std], ignore_index=True)

# Defensive dedup: if same control-process somehow in both, keep deterministic
combined["_sort"] = combined["linkage_source"].map(
    {"Deterministic": 0, "Multi-signal LLM": 1}).fillna(9)
combined = (combined
            .sort_values("_sort")
            .drop_duplicates(subset=["Control_ID","l3_process_UUID"], keep="first")
            .drop(columns=["_sort"]))

print(f"  Combined pairs      : {len(combined):,}  "
      f"({combined['Control_ID'].nunique():,} distinct controls)")

# ─────────────────────────────────────────────────────────────────────────────
#  JOIN CONTROL AND PROCESS COLUMNS
# ─────────────────────────────────────────────────────────────────────────────

print("  Joining control columns...")
combined = combined.merge(controls, on="Control_ID", how="left")

print("  Joining process columns...")
combined = combined.merge(procs_slim, on="l3_process_UUID", how="left")

# Handle duplicate column names that may arise from the join
# (e.g. l3_activity_name might exist in both det sheet and procs)
combined = combined.loc[:, ~combined.columns.duplicated(keep="first")]

print(f"  Rows before product expansion : {len(combined):,}")
print(f"  Columns before expansion      : {len(combined.columns):,}")

# ─────────────────────────────────────────────────────────────────────────────
#  EXPAND BY PRODUCT
# ─────────────────────────────────────────────────────────────────────────────

print("\n  Expanding by product (one row per product)...")

if has_final_product:
    # Split comma-separated products into individual rows
    def expand_products(df):
        rows = []
        for _, row in df.iterrows():
            raw = str(row.get("final_product","") or "").strip()
            products = [p.strip() for p in raw.split(",") if p.strip()] \
                       if raw else []
            if not products:
                products = ["No product assigned"]
            for prod in products:
                r = row.to_dict()
                r["product"] = prod
                rows.append(r)
        return pd.DataFrame(rows)

    expanded = expand_products(combined)
else:
    expanded = combined.copy()
    expanded["product"] = "No product data — check final_product column"

print(f"  Rows after product expansion  : {len(expanded):,}")
print(f"  Unique products               : {expanded['product'].nunique():,}")

# ─────────────────────────────────────────────────────────────────────────────
#  REORDER COLUMNS
# ─────────────────────────────────────────────────────────────────────────────

# Priority columns first — product and linkage info, then control, then process
priority = [
    # Product (the expanded grain)
    "product",
    # Linkage metadata
    "linkage_source",
    "confidence",
    "link_type",
    "primary_signal",
    "llm_rationale",
    "requires_sme_review",
    # Control identifiers
    "Control_ID",
    "CTRL_NAME",
    "gold_control_code",
    "gold_control_title",
    # Process identifiers
    "l3_process_UUID",
    "l3_activity_id",
    "l3_activity_name",
    "l2_process_id",
    "l2_process_name",
    "process_category",
    "process_lifecycle_stage",
    # Product context (original field for reference)
    "final_product",
]

front  = [c for c in priority if c in expanded.columns]
rest   = [c for c in expanded.columns if c not in front]
expanded = expanded[front + rest]

# ─────────────────────────────────────────────────────────────────────────────
#  SUMMARIES
# ─────────────────────────────────────────────────────────────────────────────

# By linkage source
by_source = (
    expanded.groupby("linkage_source")
    .agg(
        rows              =("product",      "count"),
        distinct_controls =("Control_ID",   "nunique"),
        distinct_processes=("l3_process_UUID","nunique"),
        distinct_products =("product",      "nunique"),
    )
    .reset_index()
)

# By product (top products by row count)
by_product = (
    expanded.groupby("product")
    .agg(
        row_count          =("Control_ID",    "count"),
        distinct_controls  =("Control_ID",    "nunique"),
        distinct_processes =("l3_process_UUID","nunique"),
    )
    .reset_index()
    .sort_values("row_count", ascending=False)
)

# By payment category
by_category = (
    expanded.groupby("process_category")
    .agg(
        rows              =("product",       "count"),
        distinct_controls =("Control_ID",    "nunique"),
        distinct_processes=("l3_process_UUID","nunique"),
        distinct_products =("product",       "nunique"),
    )
    .reset_index()
    .sort_values("rows", ascending=False)
) if "process_category" in expanded.columns else pd.DataFrame()

# Headline summary
n_det    = (expanded["linkage_source"]=="Deterministic").sum()
n_llm    = (expanded["linkage_source"]=="Multi-signal LLM").sum()
summary_rows = [
    ("── LINKAGE COUNTS (before product expansion) ──────────",""),
    ("Deterministic control-process pairs",  len(det_pairs)),
    ("Multi-signal LLM control-process pairs", len(llm_pairs)),
    ("Combined (unique pairs)",              len(combined)),
    ("Distinct controls covered",            combined["Control_ID"].nunique()),
    ("── PRODUCT EXPANSION ──────────────────────────────────",""),
    ("Total rows (one per ctrl × proc × product)", len(expanded)),
    ("Rows from deterministic linkages",     int(n_det)),
    ("Rows from multi-signal LLM linkages",  int(n_llm)),
    ("Distinct products",                    expanded["product"].nunique()),
    ("Rows with no product assigned",
     int((expanded["product"]=="No product assigned").sum())),
    ("── COVERAGE ────────────────────────────────────────────",""),
    ("Distinct controls",    expanded["Control_ID"].nunique()),
    ("Distinct processes",   expanded["l3_process_UUID"].nunique()),
    ("Distinct categories",
     expanded["process_category"].nunique()
     if "process_category" in expanded.columns else "n/a"),
]
summary_df = pd.DataFrame(summary_rows, columns=["Metric","Value"])

# ─────────────────────────────────────────────────────────────────────────────
#  VALIDATION
# ─────────────────────────────────────────────────────────────────────────────

checks = []
def chk(name, expected, actual, note=""):
    passed = str(expected) == str(actual)
    checks.append({"check":name,"expected":str(expected),
                   "actual":str(actual),"pass":"PASS" if passed else "FAIL","note":note})
    print(f"  [{'PASS' if passed else 'FAIL'}] {name}")

print("\n  Validation checks...")
chk("No null Control_ID in expanded table", 0,
    expanded["Control_ID"].isna().sum())
chk("No null l3_process_UUID in expanded table", 0,
    expanded["l3_process_UUID"].isna().sum())
chk("product column populated for every row", 0,
    expanded["product"].isna().sum())
chk("No overlap between deterministic and LLM controls", 0,
    len(set(det_pairs["Control_ID"]) & set(llm_pairs["Control_ID"])))
chk("Deterministic control count matches source", len(det_pairs),
    len(det_pairs), "Informational")
chk("LLM control count matches source", len(llm_pairs),
    len(llm_pairs), "Informational")

# ─────────────────────────────────────────────────────────────────────────────
#  WRITE OUTPUT
# ─────────────────────────────────────────────────────────────────────────────

print(f"\n  Writing to:\n  {OUTPUT_FILE}")
OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)

with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as w:
    summary_df.to_excel(  w, index=False, sheet_name="summary")
    expanded.to_excel(    w, index=False, sheet_name="consolidated_by_product")
    by_source.to_excel(   w, index=False, sheet_name="by_linkage_source")
    by_product.to_excel(  w, index=False, sheet_name="by_product")
    if not by_category.empty:
        by_category.to_excel(w, index=False, sheet_name="by_category")
    pd.DataFrame(checks).to_excel(w, index=False, sheet_name="validation_checks")

print(f"\n  Sheets:")
print(f"    summary                  — headline counts")
print(f"    consolidated_by_product  — {len(expanded):,} rows "
      f"(ctrl × process × product)")
print(f"    by_linkage_source        — deterministic vs LLM breakdown")
print(f"    by_product               — row count per product")
print(f"    by_category              — row count per payment category")
print(f"    validation_checks        — integrity checks")
print(f"\n  Done.")
print("=" * 70)
