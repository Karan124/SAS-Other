"""
payment_control_linkage_analysis.py
─────────────────────────────────────────────────────────────────────────────
Payments Controls PoC — Control-to-Process Linkage Analysis

Produces exactly three outputs:
  1. linked_payment      — controls linked to at least one Holocentric payment process
  2. linked_non_payment  — controls linked to Holocentric processes, but none are payment
  3. not_linked          — controls with no Holocentric linkage at all

Input files (update CONFIG paths before running):
  FILE_CONTROLS  — juno_payment_controls_gold.xlsx
  FILE_LINKAGE   — juno_holo_deterministic_linkage.xlsx
  FILE_PROCESSES — holocentric_payment_processes.xlsx

Run:
  python payment_control_linkage_analysis.py
"""

from pathlib import Path
from collections import Counter

import pandas as pd

# ─────────────────────────────────────────────────────────────────────────────
#  CONFIG
# ─────────────────────────────────────────────────────────────────────────────

FILE_CONTROLS  = r"Z:\path\to\juno_payment_controls_gold.xlsx"
FILE_LINKAGE   = r"Z:\path\to\juno_holo_deterministic_linkage.xlsx"
FILE_PROCESSES = r"Z:\path\to\holocentric_payment_processes.xlsx"
OUTPUT_FILE    = r"Z:\path\to\payment_control_linkage_analysis.xlsx"

# Lifecycle stage variant normalisation
LIFECYCLE_NORMALISATION = {
    "Posting, Accounting & Detection":    "Posting & Accounting, Detection",
    "Posting & Accounting & Detection":   "Posting & Accounting, Detection",
    "Posting & Accounting and Detection": "Posting & Accounting, Detection",
}

UNCLASSIFIED_SIGNALS = {"0", "", "nan", "none", "null", "unclassified"}

VALID_CATEGORIES = [
    "Customer to Customer",
    "Customer to Institution",
    "Institution to Customer",
    "Institution to Institution",
    "Supplier / Contractor / Employee Payments",
]

VALID_LIFECYCLE_STAGES = [
    "Initiation & Validation & Authorisation",
    "Execution & Early Processing Assurance",
    "Clearing / Settlement",
    "Posting & Accounting, Detection",
    "Notification & Reporting",
    "Incident response, disputes, recovery followups",
]

# ─────────────────────────────────────────────────────────────────────────────
#  HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def clean_cols(df):
    df.columns = df.columns.str.strip()
    return df

def is_empty(val):
    if pd.isna(val):
        return True
    return str(val).strip().lower() in ("", "nan", "none", "null")

def norm_uuid(val):
    return None if is_empty(val) else str(val).strip().lower()

def norm_category(val):
    if is_empty(val):
        return "Unclassified / Missing"
    s = str(val).strip()
    return s if s in VALID_CATEGORIES else "Unclassified / Missing"

def norm_stage(val):
    if is_empty(val):
        return None
    s = str(val).strip()
    return LIFECYCLE_NORMALISATION.get(s, s)

def safe_mode(series):
    s = series.dropna()
    if s.empty:
        return None
    counts = Counter(s)
    max_count = max(counts.values())
    return sorted(k for k, v in counts.items() if v == max_count)[0]

def pipe_join(values):
    vals = sorted(str(v) for v in values
                  if v is not None and str(v).strip() not in ("", "nan"))
    return " | ".join(vals) if vals else ""

# ─────────────────────────────────────────────────────────────────────────────
#  LOAD
# ─────────────────────────────────────────────────────────────────────────────

def load_data():
    print("\n  Loading data...")
    controls = clean_cols(pd.read_excel(FILE_CONTROLS,  dtype=str))
    linkage  = clean_cols(pd.read_excel(FILE_LINKAGE,   dtype=str))
    procs    = clean_cols(pd.read_excel(FILE_PROCESSES,  dtype=str))
    print(f"  Controls  : {len(controls):,} rows")
    print(f"  Linkage   : {len(linkage):,} rows")
    print(f"  Processes : {len(procs):,} rows")
    return controls, linkage, procs

# ─────────────────────────────────────────────────────────────────────────────
#  NORMALISE
# ─────────────────────────────────────────────────────────────────────────────

def normalise(controls, linkage, procs):
    print("\n  Normalising...")
    controls["Control_ID"]   = controls["Control_ID"].str.strip()
    controls["gold_control"] = controls["gold_control"].str.strip()
    linkage["CTRL_ID"]           = linkage["CTRL_ID"].str.strip()
    linkage["l3_activity_uuid"]  = linkage["l3_activity_uuid"].apply(norm_uuid)
    linkage["l2_process_uuid"]   = linkage["l2_process_uuid"].apply(norm_uuid)
    procs["l3_process_UUID"]         = procs["l3_process_UUID"].apply(norm_uuid)
    procs["l2_process_UUID"]         = procs["l2_process_UUID"].apply(norm_uuid)
    procs["process_category"]        = procs["process_category"].apply(norm_category)
    procs["process_lifecycle_stage"] = procs["process_lifecycle_stage"].apply(norm_stage)
    return controls, linkage, procs

# ─────────────────────────────────────────────────────────────────────────────
#  BUILD THREE POPULATIONS
# ─────────────────────────────────────────────────────────────────────────────

def build_populations(controls, linkage, procs):
    """
    Classify all controls into exactly three groups:

    Group 1 — linked_payment
      Control has at least one Holocentric link (L3 or L2) that resolves
      to a payment process in the Holocentric payment processes file.

    Group 2 — linked_non_payment
      Control has at least one Holocentric link but none of those links
      resolve to any process in the payment processes file.

    Group 3 — not_linked
      Control has no rows in the Holocentric linkage file at all.
    """
    print("\n  Building populations...")

    # ── Split linkage into L3 rows and L2-only rows ───────────────────────────
    lk_l3 = linkage[linkage["l3_activity_uuid"].notna()].copy()
    lk_l2 = linkage[
        linkage["l3_activity_uuid"].isna() &
        linkage["l2_process_uuid"].notna()
    ].copy()
    print(f"  Linkage rows — L3: {len(lk_l3):,}  L2-only: {len(lk_l2):,}")

    # ── L3 primary join: linkage → payment processes ──────────────────────────
    l3_payment = (
        lk_l3
        .merge(
            procs[[
                "l3_process_UUID", "l2_process_UUID", "l2_process_name",
                "l3_activity_name", "process_category", "process_lifecycle_stage",
            ]],
            left_on="l3_activity_uuid",
            right_on="l3_process_UUID",
            how="inner"
        )
        .drop_duplicates(subset=["CTRL_ID", "l3_process_UUID"])
    )
    controls_with_l3_payment = set(l3_payment["CTRL_ID"])
    print(f"  Controls with L3 payment match : {len(controls_with_l3_payment):,}")

    # ── L2 fallback: only for controls with zero L3 payment matches ───────────
    needs_fallback = set(controls["Control_ID"]) - controls_with_l3_payment
    l2_eligible    = lk_l2[lk_l2["CTRL_ID"].isin(needs_fallback)]

    l2_payment = pd.DataFrame()
    if not l2_eligible.empty:
        l2_payment = (
            l2_eligible
            .merge(
                procs[[
                    "l2_process_UUID", "l3_process_UUID", "l2_process_name",
                    "l3_activity_name", "process_category", "process_lifecycle_stage",
                ]],
                left_on="l2_process_uuid",
                right_on="l2_process_UUID",
                how="inner"
            )
        )
    controls_with_l2_payment = set(l2_payment["CTRL_ID"]) if not l2_payment.empty else set()
    print(f"  Controls with L2 fallback match: {len(controls_with_l2_payment):,}")

    # All controls with ANY Holocentric linkage (L3 or L2)
    controls_any_holo = set(linkage["CTRL_ID"].dropna().str.strip().unique())

    # ── Group 1: linked_payment ───────────────────────────────────────────────
    # Aggregate L3 matches per control
    grp1_ids = controls_with_l3_payment | controls_with_l2_payment

    if not l3_payment.empty:
        l3_agg = (
            l3_payment.groupby("CTRL_ID")
            .agg(
                payment_categories      =("process_category",
                                          lambda x: pipe_join(x.dropna().unique())),
                lifecycle_stages        =("process_lifecycle_stage",
                                          lambda x: pipe_join(x.dropna().unique())),
                primary_category        =("process_category",        safe_mode),
                primary_lifecycle_stage =("process_lifecycle_stage", safe_mode),
                payment_process_count   =("l3_process_UUID",         "nunique"),
                sample_processes        =("l3_activity_name",
                                          lambda x: pipe_join(list(x.dropna().unique())[:5])),
            )
            .reset_index()
            .rename(columns={"CTRL_ID": "Control_ID"})
        )
        l3_agg["link_type"] = "L3 direct"
    else:
        l3_agg = pd.DataFrame(columns=["Control_ID"])

    if not l2_payment.empty:
        l2_agg = (
            l2_payment.groupby("CTRL_ID")
            .agg(
                payment_categories      =("process_category",
                                          lambda x: pipe_join(x.dropna().unique())),
                lifecycle_stages        =("process_lifecycle_stage",
                                          lambda x: pipe_join(x.dropna().unique())),
                primary_category        =("process_category",        safe_mode),
                primary_lifecycle_stage =("process_lifecycle_stage", safe_mode),
                payment_process_count   =("l3_process_UUID",         "nunique"),
                sample_processes        =("l3_activity_name",
                                          lambda x: pipe_join(list(x.dropna().unique())[:5])),
            )
            .reset_index()
            .rename(columns={"CTRL_ID": "Control_ID"})
        )
        l2_agg["link_type"] = "L2 fallback (Low confidence)"
    else:
        l2_agg = pd.DataFrame(columns=["Control_ID"])

    payment_agg = pd.concat([l3_agg, l2_agg], ignore_index=True)
    grp1 = (
        controls
        .merge(payment_agg, on="Control_ID", how="inner")
    )

    # ── Group 2: linked_non_payment ───────────────────────────────────────────
    # Controls with Holo linkage but not in any payment group
    grp2_ids = controls_any_holo - grp1_ids
    grp2 = controls[controls["Control_ID"].isin(grp2_ids)].copy()
    grp2["note"] = (
        "Control is linked to Holocentric processes but none of those "
        "processes appear in the payment process inventory."
    )

    # ── Group 3: not_linked ───────────────────────────────────────────────────
    grp3_ids = set(controls["Control_ID"]) - controls_any_holo
    grp3 = controls[controls["Control_ID"].isin(grp3_ids)].copy()
    grp3["note"] = (
        "Control has no rows in the Holocentric linkage file. "
        "No Holocentric process linkage exists for this control."
    )

    total = len(grp1) + len(grp2) + len(grp3)
    print(f"\n  Group 1 — linked to payment processes  : {len(grp1):,}")
    print(f"  Group 2 — linked to non-payment only   : {len(grp2):,}")
    print(f"  Group 3 — not linked to any process    : {len(grp3):,}")
    print(f"  Total                                  : {total:,}")

    return grp1, grp2, grp3

# ─────────────────────────────────────────────────────────────────────────────
#  VALIDATION
# ─────────────────────────────────────────────────────────────────────────────

def run_validation(controls, grp1, grp2, grp3, procs):
    print("\n  Running validation checks...")
    checks = []

    def chk(name, expected, actual, note=""):
        passed = str(expected) == str(actual)
        checks.append({
            "check":    name,
            "expected": str(expected),
            "actual":   str(actual),
            "pass":     "PASS" if passed else "FAIL",
            "note":     note,
        })
        print(f"  [{'PASS' if passed else 'FAIL'}] {name}")

    total = len(grp1) + len(grp2) + len(grp3)

    # 1. No control appears in more than one group
    ids1, ids2, ids3 = set(grp1["Control_ID"]), set(grp2["Control_ID"]), set(grp3["Control_ID"])
    overlap_12 = len(ids1 & ids2)
    overlap_13 = len(ids1 & ids3)
    overlap_23 = len(ids2 & ids3)
    chk("No overlap between groups",
        0, overlap_12 + overlap_13 + overlap_23,
        "If >0, a control appears in more than one group.")

    # 2. All controls accounted for
    chk("All controls assigned to exactly one group",
        len(controls), total)

    # 3. No invalid categories in Group 1
    if "payment_categories" in grp1.columns:
        raw_cats = set()
        for val in grp1["payment_categories"].dropna():
            for c in str(val).split(" | "):
                raw_cats.add(c.strip())
        bad = raw_cats - set(VALID_CATEGORIES) - {"Unclassified / Missing", ""}
        chk("All payment categories in Group 1 are valid", "None",
            str(bad) if bad else "None")

    # 4. No invalid lifecycle stages in Group 1
    if "lifecycle_stages" in grp1.columns:
        raw_stages = set()
        for val in grp1["lifecycle_stages"].dropna():
            for s in str(val).split(" | "):
                raw_stages.add(s.strip())
        bad_s = raw_stages - set(VALID_LIFECYCLE_STAGES) - {""}
        chk("All lifecycle stages in Group 1 are valid", "None",
            str(bad_s) if bad_s else "None")

    # 5. Stage normalisation complete
    variant_left = (procs["process_lifecycle_stage"] == "Posting, Accounting & Detection").sum()
    chk("Stage normalisation complete (no variant strings remain)", 0, variant_left)

    # 6. Group 2 has no payment matches (sanity)
    chk("Group 2 controls have no payment process links",
        0, len(set(grp2["Control_ID"]) & set(grp1["Control_ID"])))

    # 7. Group 3 has no Holo linkage
    chk("Group 3 controls have no Holo linkage",
        0, len(set(grp3["Control_ID"]) & set(grp1["Control_ID"]) |
               set(grp3["Control_ID"]) & set(grp2["Control_ID"])))

    return pd.DataFrame(checks)

# ─────────────────────────────────────────────────────────────────────────────
#  SUMMARY
# ─────────────────────────────────────────────────────────────────────────────

def build_summary(controls, grp1, grp2, grp3):
    total = len(grp1) + len(grp2) + len(grp3)
    rows = [
        ("Total payment controls in scope",                        len(controls)),
        ("───────────────────────────────────────────────────────","──────"),
        ("Group 1 — Linked to payment processes",                  len(grp1)),
        ("  of which: L3 direct",
          (grp1["link_type"] == "L3 direct").sum() if "link_type" in grp1.columns else ""),
        ("  of which: L2 fallback (Low confidence)",
          (grp1["link_type"] == "L2 fallback (Low confidence)").sum()
          if "link_type" in grp1.columns else ""),
        ("Group 2 — Linked to Holo processes (non-payment only)", len(grp2)),
        ("Group 3 — Not linked to any Holo process",              len(grp3)),
        ("───────────────────────────────────────────────────────","──────"),
        ("Total (must equal scope count)",                         total),
    ]
    if "payment_process_count" in grp1.columns:
        avg = grp1["payment_process_count"].astype(float).mean()
        rows.append(("Avg payment processes per Group 1 control", f"{avg:.1f}"))
    if "payment_categories" in grp1.columns:
        multi_cat = grp1["payment_categories"].str.contains(r" \| ").sum()
        rows.append(("Group 1 controls with multiple categories", int(multi_cat)))
    if "lifecycle_stages" in grp1.columns:
        multi_stg = grp1["lifecycle_stages"].str.contains(r" \| ").sum()
        rows.append(("Group 1 controls with multiple lifecycle stages", int(multi_stg)))
    return pd.DataFrame(rows, columns=["Metric", "Value"])

# ─────────────────────────────────────────────────────────────────────────────
#  WRITE OUTPUT
# ─────────────────────────────────────────────────────────────────────────────

def write_outputs(grp1, grp2, grp3, summary, validation):
    print(f"\n  Writing to:\n  {OUTPUT_FILE}")
    Path(OUTPUT_FILE).parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as w:
        summary.to_excel(    w, index=False, sheet_name="summary")
        grp1.to_excel(       w, index=False, sheet_name="linked_payment")
        grp2.to_excel(       w, index=False, sheet_name="linked_non_payment")
        grp3.to_excel(       w, index=False, sheet_name="not_linked")
        validation.to_excel( w, index=False, sheet_name="validation_checks")

    print(f"\n  Sheets created:")
    print(f"    summary             — key counts")
    print(f"    linked_payment      — {len(grp1):,} controls linked to payment processes")
    print(f"    linked_non_payment  — {len(grp2):,} controls linked to non-payment processes")
    print(f"    not_linked          — {len(grp3):,} controls with no Holo linkage")
    print(f"    validation_checks   — automated integrity checks")

# ─────────────────────────────────────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main():
    print("=" * 70)
    print("  Payment Control Linkage Analysis")
    print("=" * 70)

    controls, linkage, procs = load_data()
    controls, linkage, procs = normalise(controls, linkage, procs)
    grp1, grp2, grp3         = build_populations(controls, linkage, procs)
    validation               = run_validation(controls, grp1, grp2, grp3, procs)
    summary                  = build_summary(controls, grp1, grp2, grp3)

    write_outputs(grp1, grp2, grp3, summary, validation)

    failures = validation[validation["pass"] == "FAIL"]
    if not failures.empty:
        print(f"\n  WARNING: {len(failures)} validation check(s) FAILED:")
        for _, r in failures.iterrows():
            print(f"    - {r['check']}: expected {r['expected']}, got {r['actual']}")
    else:
        print("\n  All validation checks passed.")

    print("\n" + "=" * 70)
    print("  Done.")
    print("=" * 70)

if __name__ == "__main__":
    main()
