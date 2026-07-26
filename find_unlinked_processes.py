"""
find_unlinked_processes.py
─────────────────────────────────────────────────────────────────────────────
Payments Controls PoC — Payment Process Coverage Gap Analysis

Two operations:

  Operation 1 — Payment processes not linked to any PAYMENT controls
    Population : ~780 core payment processes (holocentric_payment_processes.xlsx)
    Controls   : 708 payment controls (juno_payment_controls_gold.xlsx)
    Linkage    : juno_holo_deterministic_linkage.xlsx

  Operation 2 — Payment processes not linked to ANY controls at all
    Population : same ~780 payment processes
    Controls   : all controls (big_table/control_text_fields.csv, CTRL_ID column)
    Linkage    : juno_holo_deterministic_linkage.xlsx

UUID normalisation: all UUIDs stripped of whitespace and uppercased before
any join operation. Format preserved: {XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}

A process is considered LINKED if its l3_process_UUID matches l3_activity_uuid
in the linkage file (L3 match), OR its l2_process_UUID matches l2_process_uuid
in the linkage file (L2 match), where the linked control is in the target
control population.

Run:
  python find_unlinked_processes.py
"""

from pathlib import Path
import pandas as pd

# ─────────────────────────────────────────────────────────────────────────────
#  CONFIG
# ─────────────────────────────────────────────────────────────────────────────

BASE_DIR = Path(r"C:\Users\m061400\ai-test\big_table")

FILE_PROCESSES       = BASE_DIR / "holocentric_payment_processes.xlsx"
FILE_LINKAGE         = BASE_DIR / "juno_holo_deterministic_linkage.xlsx"
FILE_PAYMENT_CTRLS   = BASE_DIR / "juno_payment_controls_gold.xlsx"
FILE_ALL_CTRLS       = BASE_DIR / "control_text_fields.csv"
OUTPUT_FILE          = BASE_DIR / "phase1e_outputs" / "unlinked_payment_processes.xlsx"

LIFECYCLE_NORMALISATION = {
    "Posting, Accounting & Detection":    "Posting & Accounting, Detection",
    "Posting & Accounting & Detection":   "Posting & Accounting, Detection",
    "Posting & Accounting and Detection": "Posting & Accounting, Detection",
}

CT_TITLES = {
    "CT1":"Validation of Human-Entered Data at Input",
    "CT2":"Payment processing error detection",
    "CT3":"Early Identification of Duplications and Processing Errors",
    "CT4":"Payment processing interface and batch error resolution",
    "CT5":"Incident response","CT6":"Master/Reference data input validation",
    "CT14":"Logging and monitoring","CT22":"Mistaken Internet Payment Reports",
    "CT23":"Provision of Confirmations and Notifications",
    "CT24":"Treatment of Unauthorised and Disputed Transactions",
    "CT27":"Records retention",
}

_S1 = ["Initiation & Validation & Authorisation"]
_S2 = ["Execution & Early Processing Assurance"]
_S3 = ["Clearing / Settlement"]
_S4 = ["Posting & Accounting, Detection"]
_S5 = ["Notification & Reporting"]
_S6 = ["Incident response, disputes, recovery followups"]
_BASE = {"CT1":_S1,"CT6":_S1,"CT3":_S2,"CT4":_S2,"CT14":_S3,"CT2":_S4,
         "CT27":_S4,"CT5":_S6}
CT_CATEGORY_STAGE_EXPECTED = {
    "Customer to Customer":   {**_BASE,"CT22":_S5,"CT23":_S5,"CT24":_S6},
    "Customer to Institution":{**_BASE,"CT22":_S5,"CT23":_S5,"CT24":_S6},
    "Institution to Customer":{**_BASE,"CT23":_S5},
    "Institution to Institution":{**_BASE},
    "Supplier / Contractor / Employee Payments":{**_BASE},
}

# ─────────────────────────────────────────────────────────────────────────────
#  HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def clean_cols(df):
    df.columns = df.columns.str.strip()
    return df

def norm_uuid(val):
    """
    Normalise a UUID for joining.
    - Strips whitespace
    - Converts to UPPERCASE (preserves curly-brace format)
    - Returns None for blank / null / nan values
    """
    if pd.isna(val):
        return None
    s = str(val).strip()
    if s.lower() in ("", "nan", "none", "null"):
        return None
    return s.upper()

def norm_stage(val):
    if pd.isna(val) or str(val).strip() == "":
        return None
    s = str(val).strip()
    return LIFECYCLE_NORMALISATION.get(s, s)

def expected_cts(category, stage):
    cat_map = CT_CATEGORY_STAGE_EXPECTED.get(str(category).strip(), {})
    return [ct for ct, stages in cat_map.items()
            if stage in stages] if cat_map else []

# ─────────────────────────────────────────────────────────────────────────────
#  LOAD FILES
# ─────────────────────────────────────────────────────────────────────────────

print("=" * 70)
print("  Payment Process Coverage Gap Analysis")
print("=" * 70)

print("\n  Loading files...")

# Payment processes (~780 curated)
procs = clean_cols(pd.read_excel(FILE_PROCESSES, dtype=str, engine="openpyxl"))
procs["l3_process_UUID"] = procs["l3_process_UUID"].apply(norm_uuid)
procs["l2_process_UUID"] = procs["l2_process_UUID"].apply(norm_uuid)
procs["process_lifecycle_stage"] = procs["process_lifecycle_stage"].apply(norm_stage)

# Rename product/service col if it has a slash
prod_col = next((c for c in procs.columns
                 if "product" in c.lower() and "service" in c.lower()), None)
if prod_col and prod_col != "l3_activity_product_service":
    procs = procs.rename(columns={prod_col: "l3_activity_product_service"})

procs = procs.dropna(subset=["l3_process_UUID"])
procs = procs.drop_duplicates(subset=["l3_process_UUID"])
print(f"  Payment processes     : {len(procs):,}")

# Linkage file
linkage = clean_cols(pd.read_excel(FILE_LINKAGE, dtype=str, engine="openpyxl"))
linkage["CTRL_ID"]          = linkage["CTRL_ID"].str.strip().str.upper()
linkage["l3_activity_uuid"] = linkage["l3_activity_uuid"].apply(norm_uuid)
linkage["l2_process_uuid"]  = linkage["l2_process_uuid"].apply(norm_uuid)
print(f"  Linkage rows          : {len(linkage):,}")

# 708 payment controls
pay_ctrls = clean_cols(pd.read_excel(FILE_PAYMENT_CTRLS, dtype=str, engine="openpyxl"))
pay_ctrls["Control_ID"] = pay_ctrls["Control_ID"].str.strip().str.upper()
pay_ctrls = pay_ctrls.drop_duplicates(subset=["Control_ID"])
payment_ctrl_ids = set(pay_ctrls["Control_ID"].dropna())
print(f"  Payment controls      : {len(payment_ctrl_ids):,}")

# All controls
all_ctrls = clean_cols(pd.read_csv(FILE_ALL_CTRLS, dtype=str))
all_ctrls["CTRL_ID"] = all_ctrls["CTRL_ID"].str.strip().str.upper()
all_ctrl_ids = set(all_ctrls["CTRL_ID"].dropna())
print(f"  All controls (CSV)    : {len(all_ctrl_ids):,}")

# ─────────────────────────────────────────────────────────────────────────────
#  CORE LINKAGE LOGIC
# ─────────────────────────────────────────────────────────────────────────────

def find_linked_process_uuids(linkage_df, ctrl_id_set):
    """
    Given a linkage dataframe and a set of control IDs, return:
      - linked_by_l3 : set of l3_process_UUIDs linked via L3 match
      - linked_by_l2 : set of l2_process_UUIDs linked via L2 match
    A process UUID is "linked" if ANY row in the filtered linkage file
    points to it from within the target control population.
    """
    # Filter linkage to only the target control population
    filtered = linkage_df[linkage_df["CTRL_ID"].isin(ctrl_id_set)].copy()

    # L3: l3_activity_uuid in the filtered linkage → these are linked process UUIDs
    linked_by_l3 = set(filtered["l3_activity_uuid"].dropna().unique())

    # L2: l2_process_uuid in the filtered linkage → linked via L2
    linked_by_l2 = set(filtered["l2_process_uuid"].dropna().unique())

    return linked_by_l3, linked_by_l2


def classify_processes(procs_df, linked_l3, linked_l2):
    """
    For each process in procs_df, determine if it is linked:
      - via L3: l3_process_UUID in linked_l3
      - via L2: l2_process_UUID in linked_l2
      - not linked: neither
    Returns the dataframe with two new columns: is_linked, link_match_type
    """
    df = procs_df.copy()
    df["matched_l3"] = df["l3_process_UUID"].isin(linked_l3)
    df["matched_l2"] = df["l2_process_UUID"].isin(linked_l2)

    def match_type(row):
        if row["matched_l3"] and row["matched_l2"]:
            return "L3 and L2"
        if row["matched_l3"]:
            return "L3"
        if row["matched_l2"]:
            return "L2 only"
        return "Not linked"

    df["link_match_type"] = df.apply(match_type, axis=1)
    df["is_linked"]       = df["matched_l3"] | df["matched_l2"]
    return df


# ─────────────────────────────────────────────────────────────────────────────
#  OPERATION 1 — processes not linked to PAYMENT controls
# ─────────────────────────────────────────────────────────────────────────────

print("\n  Operation 1: processes not linked to payment controls...")
l3_pay, l2_pay = find_linked_process_uuids(linkage, payment_ctrl_ids)
procs_op1 = classify_processes(procs, l3_pay, l2_pay)

op1_linked   = procs_op1[procs_op1["is_linked"]].copy()
op1_unlinked = procs_op1[~procs_op1["is_linked"]].copy()

print(f"  Linked to payment controls    : {len(op1_linked):,}")
print(f"  NOT linked to payment controls: {len(op1_unlinked):,}")

# ─────────────────────────────────────────────────────────────────────────────
#  OPERATION 2 — processes not linked to ANY controls
# ─────────────────────────────────────────────────────────────────────────────

print("\n  Operation 2: processes not linked to any controls...")
l3_all, l2_all = find_linked_process_uuids(linkage, all_ctrl_ids)
procs_op2 = classify_processes(procs, l3_all, l2_all)

op2_linked   = procs_op2[procs_op2["is_linked"]].copy()
op2_unlinked = procs_op2[~procs_op2["is_linked"]].copy()

print(f"  Linked to any control         : {len(op2_linked):,}")
print(f"  NOT linked to any control     : {len(op2_unlinked):,}")

# ─────────────────────────────────────────────────────────────────────────────
#  CROSS-COMPARISON — categorise every process
# ─────────────────────────────────────────────────────────────────────────────

def categorise(row):
    """
    Four mutually exclusive categories based on the two operations.
    """
    linked_pay = row["l3_process_UUID"] in l3_pay or row["l2_process_UUID"] in l2_pay
    linked_any = row["l3_process_UUID"] in l3_all or row["l2_process_UUID"] in l2_all

    if linked_pay:
        return "Linked to payment controls"
    if linked_any and not linked_pay:
        return "Linked to non-payment controls only"
    return "Not linked to any controls"

procs_cross = procs.copy()
procs_cross["coverage_category"] = procs_cross.apply(categorise, axis=1)
procs_cross = procs_cross.sort_values(
    ["coverage_category","process_category","process_lifecycle_stage"])

cross_counts = procs_cross["coverage_category"].value_counts()
for cat, n in cross_counts.items():
    print(f"    {cat}: {n:,}")

# ─────────────────────────────────────────────────────────────────────────────
#  ENRICH UNLINKED WITH EXPECTED CT TYPES
# ─────────────────────────────────────────────────────────────────────────────

def add_expected_cts(df):
    df = df.copy()
    df["expected_ct_codes"] = df.apply(
        lambda r: "; ".join(expected_cts(
            r.get("process_category",""), r.get("process_lifecycle_stage",""))),
        axis=1
    )
    df["expected_ct_titles"] = df.apply(
        lambda r: "; ".join(
            f"{ct} — {CT_TITLES.get(ct,'')}"
            for ct in expected_cts(
                r.get("process_category",""), r.get("process_lifecycle_stage",""))),
        axis=1
    )
    df["expected_ct_count"] = df.apply(
        lambda r: len(expected_cts(
            r.get("process_category",""), r.get("process_lifecycle_stage",""))),
        axis=1
    )
    return df

op1_unlinked = add_expected_cts(op1_unlinked)
op2_unlinked = add_expected_cts(op2_unlinked)
procs_cross  = add_expected_cts(procs_cross)

# ─────────────────────────────────────────────────────────────────────────────
#  BREAKDOWN TABLES
# ─────────────────────────────────────────────────────────────────────────────

def breakdown(unlinked_df, all_df, col):
    if col not in unlinked_df.columns:
        return pd.DataFrame()
    total   = all_df[col].value_counts().rename("total")
    unlinked= unlinked_df[col].value_counts().rename("unlinked")
    t = pd.concat([total, unlinked], axis=1).fillna(0)
    t["unlinked"] = t["unlinked"].astype(int)
    t["linked"]   = t["total"].astype(int) - t["unlinked"]
    t["pct_unlinked"] = (t["unlinked"] / t["total"] * 100).round(1)
    return t.reset_index().rename(columns={"index": col}).sort_values(
        "unlinked", ascending=False)

op1_by_cat   = breakdown(op1_unlinked, procs, "process_category")
op1_by_stage = breakdown(op1_unlinked, procs, "process_lifecycle_stage")
op2_by_cat   = breakdown(op2_unlinked, procs, "process_category")
op2_by_stage = breakdown(op2_unlinked, procs, "process_lifecycle_stage")

# Expected CT gap summary for Op1 (which CTs are missing most)
ct_gap_rows = []
for ct, title in CT_TITLES.items():
    n = sum(1 for _, r in op1_unlinked.iterrows()
            if ct in expected_cts(r.get("process_category",""),
                                  r.get("process_lifecycle_stage","")))
    if n > 0:
        ct_gap_rows.append({"gold_control_code":ct,
                            "gold_control_title":title,
                            "op1_unlinked_processes_needing_ct":n})
ct_gap_df = (pd.DataFrame(ct_gap_rows)
             .sort_values("op1_unlinked_processes_needing_ct", ascending=False)
             if ct_gap_rows else pd.DataFrame())

# ─────────────────────────────────────────────────────────────────────────────
#  OUTPUT COLUMN ORDER
# ─────────────────────────────────────────────────────────────────────────────

KEY_COLS = [
    "l3_process_UUID","l3_activity_id","l3_activity_name",
    "l3_activity_description","l2_process_id","l2_process_name",
    "process_category","process_lifecycle_stage",
    "expected_ct_codes","expected_ct_titles","expected_ct_count",
    "value_stream_name","vcm_library_name","payment_rationale",
]

def order_cols(df, extra_front=None):
    front = (extra_front or []) + [c for c in KEY_COLS if c in df.columns]
    rest  = [c for c in df.columns if c not in front]
    return df[[c for c in front if c in df.columns] + rest]

op1_out   = order_cols(op1_unlinked)
op2_out   = order_cols(op2_unlinked)
cross_out = order_cols(procs_cross, extra_front=["coverage_category"])

# ─────────────────────────────────────────────────────────────────────────────
#  SUMMARY
# ─────────────────────────────────────────────────────────────────────────────

summary_rows = [
    ("── INPUT POPULATIONS ───────────────────────────────────────",""),
    ("Total payment processes",                   len(procs)),
    ("Payment controls (708 population)",         len(payment_ctrl_ids)),
    ("All controls (control_text_fields.csv)",    len(all_ctrl_ids)),
    ("Total linkage rows",                        len(linkage)),
    ("── OPERATION 1: vs PAYMENT CONTROLS ───────────────────────",""),
    ("Linked to ≥1 payment control",             len(op1_linked)),
    ("NOT linked to any payment control",        len(op1_unlinked)),
    ("Coverage rate (payment controls)",
     f"{len(op1_linked)/len(procs)*100:.1f}%"),
    ("Gap rate (payment controls)",
     f"{len(op1_unlinked)/len(procs)*100:.1f}%"),
    ("── OPERATION 2: vs ALL CONTROLS ───────────────────────────",""),
    ("Linked to ≥1 control (any)",               len(op2_linked)),
    ("NOT linked to any control",                len(op2_unlinked)),
    ("Coverage rate (all controls)",
     f"{len(op2_linked)/len(procs)*100:.1f}%"),
    ("Gap rate (all controls)",                  f"{len(op2_unlinked)/len(procs)*100:.1f}%"),
    ("── CROSS-COMPARISON ───────────────────────────────────────",""),
    ("Linked to payment controls",
     int(cross_counts.get("Linked to payment controls", 0))),
    ("Linked to non-payment controls only",
     int(cross_counts.get("Linked to non-payment controls only", 0))),
    ("Not linked to any controls",
     int(cross_counts.get("Not linked to any controls", 0))),
]
summary_df = pd.DataFrame(summary_rows, columns=["Metric","Value"])

# ─────────────────────────────────────────────────────────────────────────────
#  VALIDATION
# ─────────────────────────────────────────────────────────────────────────────

checks = []
def chk(name, expected, actual, note=""):
    passed = str(expected) == str(actual)
    checks.append({"check":name,"expected":str(expected),
                   "actual":str(actual),
                   "pass":"PASS" if passed else "FAIL","note":note})
    print(f"  [{'PASS' if passed else 'FAIL'}] {name}")

print("\n  Validation checks...")
chk("Op1 linked + unlinked = total processes",
    len(procs), len(op1_linked) + len(op1_unlinked))
chk("Op2 linked + unlinked = total processes",
    len(procs), len(op2_linked) + len(op2_unlinked))
chk("Cross-comparison sums to total processes",
    len(procs), int(cross_counts.sum()))
chk("Op2 unlinked ≤ Op1 unlinked (all controls ≥ payment controls)",
    True, len(op2_unlinked) <= len(op1_unlinked),
    "If false: more processes are unlinked from all controls than from payment controls only — investigate.")
chk("No null l3_process_UUID in process population", 0,
    procs["l3_process_UUID"].isna().sum())
chk("Payment controls found in linkage file", True,
    bool(linkage["CTRL_ID"].isin(payment_ctrl_ids).any()),
    "If false: CTRL_ID format mismatch between controls and linkage files.")
chk("All controls (CSV) found in linkage file", True,
    bool(linkage["CTRL_ID"].isin(all_ctrl_ids).any()),
    "If false: CTRL_ID format mismatch between CSV and linkage files.")

# ─────────────────────────────────────────────────────────────────────────────
#  WRITE OUTPUT
# ─────────────────────────────────────────────────────────────────────────────

print(f"\n  Writing to:\n  {OUTPUT_FILE}")
OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)

with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as w:
    summary_df.to_excel(w, index=False, sheet_name="summary")
    cross_out.to_excel( w, index=False, sheet_name="all_processes_categorised")
    op1_out.to_excel(   w, index=False, sheet_name="op1_unlinked_from_pay_ctrls")
    op2_out.to_excel(   w, index=False, sheet_name="op2_unlinked_from_all_ctrls")
    op1_by_cat.to_excel(  w, index=False, sheet_name="op1_by_category")
    op1_by_stage.to_excel(w, index=False, sheet_name="op1_by_lifecycle_stage")
    op2_by_cat.to_excel(  w, index=False, sheet_name="op2_by_category")
    op2_by_stage.to_excel(w, index=False, sheet_name="op2_by_lifecycle_stage")
    if not ct_gap_df.empty:
        ct_gap_df.to_excel(w, index=False, sheet_name="op1_ct_gap_summary")
    pd.DataFrame(checks).to_excel(w, index=False, sheet_name="validation_checks")

print(f"\n  Sheets (10):")
print(f"    summary                        — headline counts for both operations")
print(f"    all_processes_categorised      — {len(procs):,} processes, "
      f"each assigned to one of three coverage categories")
print(f"    op1_unlinked_from_pay_ctrls    — {len(op1_out):,} processes "
      f"not linked to any of the 708 payment controls")
print(f"    op2_unlinked_from_all_ctrls    — {len(op2_out):,} processes "
      f"not linked to any control at all")
print(f"    op1_by_category / op1_by_lifecycle_stage  — Op1 breakdowns")
print(f"    op2_by_category / op2_by_lifecycle_stage  — Op2 breakdowns")
print(f"    op1_ct_gap_summary             — which CT types are missing most")
print(f"    validation_checks              — 7 integrity checks")
print(f"\n  Done.")
print("=" * 70)
