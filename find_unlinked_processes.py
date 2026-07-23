"""
find_unlinked_processes.py
─────────────────────────────────────────────────────────────────────────────
Identifies Holocentric payment processes that have no linked JUNO controls.

Reads:
  holocentric_payment_processes.xlsx     — full payment process inventory
  payment_control_linkage_analysis.xlsx  — existing linkage output
    └─ linked_payment_detail             — confirmed control-process pairs

Writes:
  unlinked_payment_processes.xlsx with sheets:
    summary                 — headline counts and coverage rate
    unlinked_processes      — full list of processes with no linked controls
                              + expected but missing CT types per process
    by_category             — breakdown by payment category
    by_lifecycle_stage      — breakdown by lifecycle stage
    expected_ct_gap_summary — which CT outcomes have the most uncovered processes
"""

from pathlib import Path
import pandas as pd

# ─────────────────────────────────────────────────────────────────────────────
#  CONFIG
# ─────────────────────────────────────────────────────────────────────────────

FILE_PROCESSES  = Path(r"C:\Users\m061400\ai-test\big_table\holocentric_payment_processes.xlsx")
FILE_LINKAGE_WB = Path(r"C:\Users\m061400\ai-test\big_table\phase1e_outputs\payment_control_linkage_analysis.xlsx")
OUTPUT_FILE     = Path(r"C:\Users\m061400\ai-test\big_table\phase1e_outputs\unlinked_payment_processes.xlsx")

# ─────────────────────────────────────────────────────────────────────────────
#  CT-TO-LIFECYCLE EXPECTED MAP  (copied from payment_control_linkage_analysis.py)
# ─────────────────────────────────────────────────────────────────────────────

_S1 = ["Initiation & Validation & Authorisation"]
_S2 = ["Execution & Early Processing Assurance"]
_S3 = ["Clearing / Settlement"]
_S4 = ["Posting & Accounting, Detection"]
_S5 = ["Notification & Reporting"]
_S6 = ["Incident response, disputes, recovery followups"]

_BASE = {
    "CT1": _S1, "CT6": _S1,
    "CT3": _S2, "CT4": _S2,
    "CT14": _S3,
    "CT2": _S4, "CT27": _S4,
    "CT5": _S6,
}

CT_CATEGORY_STAGE_EXPECTED = {
    "Customer to Customer":   {**_BASE, "CT22": _S5, "CT23": _S5, "CT24": _S6},
    "Customer to Institution":{**_BASE, "CT22": _S5, "CT23": _S5, "CT24": _S6},
    "Institution to Customer":{**_BASE, "CT23": _S5},
    "Institution to Institution": {**_BASE},
    "Supplier / Contractor / Employee Payments": {**_BASE},
}

CT_TITLES = {
    "CT1":"Validation of Human-Entered Data at Input",
    "CT2":"Payment processing error detection",
    "CT3":"Early Identification of Duplications and Processing Errors",
    "CT4":"Payment processing interface and batch error resolution",
    "CT5":"Incident response",
    "CT6":"Master/Reference data input validation",
    "CT14":"Logging and monitoring",
    "CT22":"Mistaken Internet Payment Reports",
    "CT23":"Provision of Confirmations and Notifications",
    "CT24":"Treatment of Unauthorised and Disputed Transactions",
    "CT27":"Records retention",
}

LIFECYCLE_NORMALISATION = {
    "Posting, Accounting & Detection":    "Posting & Accounting, Detection",
    "Posting & Accounting & Detection":   "Posting & Accounting, Detection",
    "Posting & Accounting and Detection": "Posting & Accounting, Detection",
}

# ─────────────────────────────────────────────────────────────────────────────
#  HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def norm_stage(val):
    if pd.isna(val) or str(val).strip() == "":
        return None
    s = str(val).strip()
    return LIFECYCLE_NORMALISATION.get(s, s)


def expected_cts_for_process(category, stage):
    """
    Return the list of CT codes expected for a process given its
    payment category and lifecycle stage.
    Returns empty list if category or stage not in the mapping.
    """
    cat_map = CT_CATEGORY_STAGE_EXPECTED.get(str(category).strip(), {})
    if not cat_map:
        return []
    return [ct for ct, stages in cat_map.items() if stage in stages]


# ─────────────────────────────────────────────────────────────────────────────
#  LOAD
# ─────────────────────────────────────────────────────────────────────────────

print("=" * 65)
print("  Payment Process Control Coverage Gap Analysis")
print("=" * 65)

print("\n  Loading payment processes...")
procs = pd.read_excel(FILE_PROCESSES, dtype=str, engine="openpyxl")
procs.columns = procs.columns.str.strip()
procs["l3_process_UUID"] = procs["l3_process_UUID"].str.strip().str.lower()
procs = procs.dropna(subset=["l3_process_UUID"])
procs = procs[procs["l3_process_UUID"] != ""]
procs["process_lifecycle_stage"] = procs["process_lifecycle_stage"].apply(norm_stage)
procs = procs.drop_duplicates(subset=["l3_process_UUID"])
print(f"  Unique payment processes : {len(procs):,}")

print("  Loading linked_payment_detail from linkage workbook...")
detail = pd.read_excel(
    FILE_LINKAGE_WB,
    sheet_name="linked_payment_detail",
    dtype=str,
    engine="openpyxl"
)
detail.columns = detail.columns.str.strip()
detail["l3_process_UUID"] = detail["l3_process_UUID"].str.strip().str.lower()
linked_uuids = set(detail["l3_process_UUID"].dropna().unique())
print(f"  Processes with linked controls : {len(linked_uuids):,}")

# ─────────────────────────────────────────────────────────────────────────────
#  IDENTIFY UNLINKED PROCESSES
# ─────────────────────────────────────────────────────────────────────────────

print("\n  Identifying unlinked processes...")
all_uuids    = set(procs["l3_process_UUID"].unique())
unlinked_ids = all_uuids - linked_uuids

unlinked = procs[procs["l3_process_UUID"].isin(unlinked_ids)].copy()
unlinked = unlinked.sort_values(
    ["process_category", "process_lifecycle_stage", "l3_process_UUID"]
)

print(f"  Processes WITH linked controls    : {len(linked_uuids):,}")
print(f"  Processes WITHOUT linked controls : {len(unlinked):,}")
print(f"  Total payment processes           : {len(all_uuids):,}")
print(f"  Control coverage rate             : {len(linked_uuids)/len(all_uuids)*100:.1f}%")
print(f"  Coverage gap rate                 : {len(unlinked)/len(all_uuids)*100:.1f}%")

# ─────────────────────────────────────────────────────────────────────────────
#  ENRICH WITH EXPECTED BUT MISSING CT TYPES
# ─────────────────────────────────────────────────────────────────────────────

print("\n  Computing expected but missing CT types per process...")

def get_expected_cts(row):
    cts = expected_cts_for_process(
        row.get("process_category", ""),
        row.get("process_lifecycle_stage", "")
    )
    return cts


unlinked["expected_ct_codes"] = unlinked.apply(
    lambda r: "; ".join(get_expected_cts(r)), axis=1
)

unlinked["expected_ct_titles"] = unlinked.apply(
    lambda r: "; ".join(
        f"{ct} — {CT_TITLES.get(ct, '')}"
        for ct in get_expected_cts(r)
    ),
    axis=1
)

unlinked["expected_ct_count"] = unlinked.apply(
    lambda r: len(get_expected_cts(r)), axis=1
)

# ─────────────────────────────────────────────────────────────────────────────
#  SUMMARY BREAKDOWNS
# ─────────────────────────────────────────────────────────────────────────────

def count_table(col):
    """Compare total vs unlinked counts per value of col."""
    total_counts   = procs[col].value_counts().rename("total_processes")
    unlinked_counts= unlinked[col].value_counts().rename("unlinked_processes")
    merged = pd.concat([total_counts, unlinked_counts], axis=1).fillna(0)
    merged["unlinked_processes"] = merged["unlinked_processes"].astype(int)
    merged["linked_processes"]   = (
        merged["total_processes"].astype(int) - merged["unlinked_processes"]
    )
    merged["pct_unlinked"] = (
        merged["unlinked_processes"] / merged["total_processes"] * 100
    ).round(1)
    return merged.reset_index().rename(columns={"index": col}).sort_values(
        "unlinked_processes", ascending=False
    )

by_category = (count_table("process_category")
               if "process_category" in unlinked.columns else pd.DataFrame())

by_stage = (count_table("process_lifecycle_stage")
            if "process_lifecycle_stage" in unlinked.columns else pd.DataFrame())

# CT gap summary: which CT outcomes have the most uncovered processes
ct_gap_rows = []
for ct, title in CT_TITLES.items():
    n_processes_needing_ct = 0
    for _, row in unlinked.iterrows():
        if ct in get_expected_cts(row):
            n_processes_needing_ct += 1
    if n_processes_needing_ct > 0:
        ct_gap_rows.append({
            "gold_control_code": ct,
            "gold_control_title": title,
            "unlinked_processes_needing_this_ct": n_processes_needing_ct,
        })

ct_gap_summary = (pd.DataFrame(ct_gap_rows)
                  .sort_values("unlinked_processes_needing_this_ct", ascending=False)
                  if ct_gap_rows else pd.DataFrame())

# ─────────────────────────────────────────────────────────────────────────────
#  OVERALL SUMMARY
# ─────────────────────────────────────────────────────────────────────────────

summary_rows = [
    ("Total payment processes (unique UUIDs)",   len(all_uuids)),
    ("Processes with ≥1 linked control",         len(linked_uuids)),
    ("Processes with NO linked controls",        len(unlinked)),
    ("Control coverage rate",                   f"{len(linked_uuids)/len(all_uuids)*100:.1f}%"),
    ("Coverage gap rate",                       f"{len(unlinked)/len(all_uuids)*100:.1f}%"),
    ("Unlinked processes with known expected CTs",
     int((unlinked["expected_ct_count"] > 0).sum())),
    ("Unlinked processes with no expected CTs (category/stage unknown)",
     int((unlinked["expected_ct_count"] == 0).sum())),
]
summary_df = pd.DataFrame(summary_rows, columns=["Metric", "Value"])

# ─────────────────────────────────────────────────────────────────────────────
#  WRITE OUTPUT
# ─────────────────────────────────────────────────────────────────────────────

print(f"\n  Writing output to:\n  {OUTPUT_FILE}")
OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)

# Select and order columns for the unlinked sheet
unlinked_cols = [
    "l3_process_UUID",
    "l3_activity_id",
    "l3_activity_name",
    "l3_activity_description",
    "l2_process_id",
    "l2_process_name",
    "process_category",
    "process_lifecycle_stage",
    "expected_ct_codes",
    "expected_ct_titles",
    "expected_ct_count",
    "value_stream_name",
    "vcm_library_name",
    "payment_rationale",
]
unlinked_out = unlinked[[c for c in unlinked_cols if c in unlinked.columns]]

with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as w:
    summary_df.to_excel(   w, index=False, sheet_name="summary")
    unlinked_out.to_excel( w, index=False, sheet_name="unlinked_processes")
    if not by_category.empty:
        by_category.to_excel(w, index=False, sheet_name="by_category")
    if not by_stage.empty:
        by_stage.to_excel(  w, index=False, sheet_name="by_lifecycle_stage")
    if not ct_gap_summary.empty:
        ct_gap_summary.to_excel(w, index=False, sheet_name="expected_ct_gap_summary")

print(f"\n  Sheets:")
print(f"    summary                — headline counts and coverage rate")
print(f"    unlinked_processes     — {len(unlinked_out):,} processes with no linked controls")
print(f"                             + expected_ct_codes and expected_ct_titles columns")
print(f"    by_category            — breakdown by payment category")
print(f"    by_lifecycle_stage     — breakdown by lifecycle stage")
print(f"    expected_ct_gap_summary— which CT outcomes most need coverage")

print("\n" + "=" * 65)
print("  Done.")
print("=" * 65)
