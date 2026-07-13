"""
payment_control_linkage_analysis.py
─────────────────────────────────────────────────────────────────────────────
Payments Controls PoC — Control-to-Process Linkage Analysis

Outputs (8 sheets in one Excel workbook):
  summary                    - key counts and metrics
  linked_payment             - one row per control with at least one payment link
  linked_payment_detail      - one row per control-process pair
  linked_non_payment         - one row per control linked to non-payment Holo only
  linked_non_payment_detail  - one row per linkage row for non-payment controls
  not_linked                 - controls with no Holo linkage at all
  one_big_table              - denormalised analytical table (Population A)
                               with derived columns for CT-lifecycle alignment,
                               duplicate detection, sole control identification
  validation_checks          - automated integrity checks

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

LIFECYCLE_NORMALISATION = {
    "Posting, Accounting & Detection":    "Posting & Accounting, Detection",
    "Posting & Accounting & Detection":   "Posting & Accounting, Detection",
    "Posting & Accounting and Detection": "Posting & Accounting, Detection",
}

UNCLASSIFIED_SIGNALS = {"0", "", "nan", "none", "null", "unclassified"}

# CT-to-lifecycle expected mapping (specific CTs only)
# Broad CTs (7-21, 25, 26) apply across all stages - alignment not meaningful
CT_LIFECYCLE_EXPECTED = {
    "CT1":  ["Initiation & Validation & Authorisation"],
    "CT2":  ["Execution & Early Processing Assurance"],
    "CT3":  ["Initiation & Validation & Authorisation",
              "Execution & Early Processing Assurance"],
    "CT4":  ["Execution & Early Processing Assurance",
              "Incident response, disputes, recovery followups"],
    "CT5":  ["Incident response, disputes, recovery followups"],
    "CT6":  ["Initiation & Validation & Authorisation"],
    "CT22": ["Incident response, disputes, recovery followups"],
    "CT23": ["Notification & Reporting"],
    "CT24": ["Incident response, disputes, recovery followups"],
    "CT27": ["Posting & Accounting, Detection", "Notification & Reporting"],
    "CT28": ["Incident response, disputes, recovery followups"],
}

CT_BROAD = {f"CT{i}" for i in [7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,25,26]}

# JUNO columns to include in OBT (analytically valuable; admin fields excluded)
JUNO_OBT_COLS = [
    "Control_ID", "CTRL_NAME", "gold_control",
    "CTRL_NATRE",           # Nature: preventative / detective
    "CTRL_STUS",            # Status
    "CTRL_KEY_CONTRL",      # Key control flag
    "CTRL_FREQ",            # Frequency
    "CTRL_TYP",             # Control type
    "CTRL_ASSESS_RTNG",     # Effectiveness assessment rating
    "CTRL_OE_RTNG",         # Operating effectiveness rating
    "CTRL_DE_RTNG",         # Design effectiveness rating
    "CTRL_CTGRY_1",         # JUNO category 1
    "CTRL_CTGRY_2",         # JUNO category 2
    "CTRL_CATEGORY",        # JUNO category
    "CTRL_DESC",            # Control description
    "CTRL_EVDNCD",          # How evidenced
    "CTRL_MNTRD",           # How monitored
    "CTRL_FLDR",            # Folder (location analysis)
    "CTRL_FLDR_LVL_2",      # Folder level 2
    "COMMON_CTRL_REFERENCE",# Reference to gold/common control
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
    """Three-step normalisation: exact variant, case-insensitive variant,
    case-insensitive match against valid stages."""
    if is_empty(val):
        return None
    s = str(val).strip()
    if s in LIFECYCLE_NORMALISATION:
        return LIFECYCLE_NORMALISATION[s]
    s_lower = s.lower()
    for key, canonical in LIFECYCLE_NORMALISATION.items():
        if key.lower() == s_lower:
            return canonical
    for valid in VALID_LIFECYCLE_STAGES:
        if valid.lower() == s_lower:
            return valid
    return s

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
    procs["l3_process_UUID"]          = procs["l3_process_UUID"].apply(norm_uuid)
    procs["l2_process_UUID"]          = procs["l2_process_UUID"].apply(norm_uuid)
    procs["process_category"]         = procs["process_category"].apply(norm_category)
    procs["process_lifecycle_stage"]  = procs["process_lifecycle_stage"].apply(norm_stage)
    # Rename product/service column (contains a slash)
    prod_col = next(
        (c for c in procs.columns
         if "product" in c.lower() and "service" in c.lower()), None
    )
    if prod_col and prod_col != "l3_activity_product_service":
        procs = procs.rename(columns={prod_col: "l3_activity_product_service"})
        print(f"  Renamed '{prod_col}' to 'l3_activity_product_service'")
    return controls, linkage, procs

# ─────────────────────────────────────────────────────────────────────────────
#  BUILD POPULATIONS
# ─────────────────────────────────────────────────────────────────────────────

def build_populations(controls, linkage, procs):
    print("\n  Building populations...")

    lk_l3 = linkage[linkage["l3_activity_uuid"].notna()].copy()
    lk_l2 = linkage[
        linkage["l3_activity_uuid"].isna() &
        linkage["l2_process_uuid"].notna()
    ].copy()
    print(f"  Linkage rows  L3: {len(lk_l3):,}  L2-only: {len(lk_l2):,}")

    ctrl_key_cols = [c for c in JUNO_OBT_COLS if c in controls.columns]

    holo_detail_cols = [c for c in [
        "l3_process_UUID", "l2_process_UUID", "l2_process_name",
        "l3_activity_name", "l3_activity_description",
        "process_category", "process_lifecycle_stage",
        "l3_activity_product_service",
    ] if c in procs.columns]

    # L3 primary join
    l3_payment = (
        lk_l3
        .merge(procs[holo_detail_cols],
               left_on="l3_activity_uuid", right_on="l3_process_UUID", how="inner")
        .drop_duplicates(subset=["CTRL_ID", "l3_process_UUID"])
    )
    ctrl_l3_pay = set(l3_payment["CTRL_ID"])
    print(f"  Controls with L3 payment match : {len(ctrl_l3_pay):,}")

    # L2 fallback (only for controls with zero L3 payment matches)
    l2_eligible = lk_l2[lk_l2["CTRL_ID"].isin(set(controls["Control_ID"]) - ctrl_l3_pay)]
    l2_payment  = pd.DataFrame()
    if not l2_eligible.empty:
        l2_payment = (
            l2_eligible
            .merge(procs[holo_detail_cols],
                   left_on="l2_process_uuid", right_on="l2_process_UUID", how="inner")
        )
    ctrl_l2_pay = set(l2_payment["CTRL_ID"]) if not l2_payment.empty else set()
    print(f"  Controls with L2 fallback match: {len(ctrl_l2_pay):,}")

    controls_any_holo = set(linkage["CTRL_ID"].dropna().str.strip().unique())
    grp1_ids = ctrl_l3_pay | ctrl_l2_pay

    # Aggregate helper
    def agg_to_one_row(df, label):
        if df.empty:
            return pd.DataFrame()
        a = (
            df.groupby("CTRL_ID")
            .agg(
                payment_categories      =("process_category",
                                          lambda x: pipe_join(x.dropna().unique())),
                lifecycle_stages        =("process_lifecycle_stage",
                                          lambda x: pipe_join(x.dropna().unique())),
                primary_category        =("process_category",         safe_mode),
                primary_lifecycle_stage =("process_lifecycle_stage",  safe_mode),
                payment_process_count   =("l3_process_UUID",          "nunique"),
                sample_processes        =("l3_activity_name",
                                          lambda x: pipe_join(
                                              list(x.dropna().unique())[:5])),
            )
            .reset_index()
            .rename(columns={"CTRL_ID": "Control_ID"})
        )
        a["link_type"] = label
        return a

    agg_l3 = agg_to_one_row(l3_payment, "L3 direct")
    agg_l2 = agg_to_one_row(l2_payment, "L2 fallback (Low confidence)")

    # Group 1 summary (one row per control)
    grp1 = controls[ctrl_key_cols].merge(
        pd.concat([agg_l3, agg_l2], ignore_index=True),
        on="Control_ID", how="inner"
    )

    # Group 1 detail (one row per control-process pair)
    def enrich(df, label):
        if df.empty:
            return pd.DataFrame()
        d = controls[ctrl_key_cols].merge(
            df.rename(columns={"CTRL_ID": "Control_ID"}),
            on="Control_ID", how="inner"
        )
        d["link_type"] = label
        return d

    grp1_detail = pd.concat(
        [enrich(l3_payment, "L3 direct"),
         enrich(l2_payment, "L2 fallback (Low confidence)")],
        ignore_index=True
    )

    # Group 2 summary
    grp2_ids = controls_any_holo - grp1_ids
    grp2 = controls[ctrl_key_cols][controls["Control_ID"].isin(grp2_ids)].copy()
    grp2["note"] = (
        "Linked to Holocentric processes but none appear "
        "in the payment process inventory."
    )

    # Group 2 detail (one row per linkage row)
    lk_cols_avail = [c for c in [
        "CTRL_ID","l3_activity_uuid","l3_activity_id",
        "l2_process_uuid","l2_process_id","link_level",
        "bus_unit_bcrm_id","BUS_UNIT_FLDR_LVL_3","BUS_UNIT_FLDR_LVL_4",
        "holo_value_stream","vcm_library_name",
    ] if c in linkage.columns]
    grp2_detail = controls[ctrl_key_cols].merge(
        linkage[linkage["CTRL_ID"].isin(grp2_ids)][lk_cols_avail]
        .rename(columns={"CTRL_ID": "Control_ID"}),
        on="Control_ID", how="inner"
    )
    grp2_detail["note"] = "Linked Holo process NOT in payment process inventory"

    # Group 3
    grp3_ids = set(controls["Control_ID"]) - controls_any_holo
    grp3 = controls[ctrl_key_cols][controls["Control_ID"].isin(grp3_ids)].copy()
    grp3["note"] = "No rows in the Holo linkage file for this control."

    total = len(grp1) + len(grp2) + len(grp3)
    print(f"\n  Group 1 — linked to payment    : {len(grp1):,}")
    print(f"  Group 1 detail rows            : {len(grp1_detail):,}")
    print(f"  Group 2 — linked non-payment   : {len(grp2):,}")
    print(f"  Group 2 detail rows            : {len(grp2_detail):,}")
    print(f"  Group 3 — not linked           : {len(grp3):,}")
    print(f"  Total                          : {total:,}")

    return grp1, grp1_detail, grp2, grp2_detail, grp3

# ─────────────────────────────────────────────────────────────────────────────
#  ONE BIG TABLE
# ─────────────────────────────────────────────────────────────────────────────

def build_one_big_table(grp1_detail, controls, procs):
    """
    Denormalised analytical table — Population A only.
    One row per control-process pair enriched with:
      - All key JUNO control attributes
      - All key Holo process attributes
      - Derived: CT-lifecycle alignment
      - Derived: duplicate detection (process level + category-stage level)
      - Derived: sole control identification
    """
    print("\n  Building One Big Table...")

    obt = grp1_detail.copy()

    # Add remaining JUNO OBT columns not already in detail
    juno_extra = [c for c in JUNO_OBT_COLS
                  if c not in obt.columns and c in controls.columns]
    if juno_extra:
        obt = obt.merge(
            controls[["Control_ID"] + juno_extra],
            on="Control_ID", how="left"
        )

    # Add Holo process context columns not already in detail
    holo_context = [c for c in [
        "l3_activity_channels",
        "l3_activity_customer_segments",
        "l3_activity_product_service",
        "l3_activity_description",
        "value_stream_name",
        "vcm_library_name",
    ] if c in procs.columns and c not in obt.columns]
    if holo_context and "l3_process_UUID" in obt.columns:
        obt = obt.merge(
            procs[["l3_process_UUID"] + holo_context],
            on="l3_process_UUID", how="left"
        )

    # ── Derived: CT-lifecycle alignment ───────────────────────────────────────
    def get_alignment(row):
        ct    = str(row.get("gold_control", "")).strip()
        stage = row.get("process_lifecycle_stage")
        if ct in CT_BROAD:
            return "Broad CT — applies across all lifecycle stages"
        if ct not in CT_LIFECYCLE_EXPECTED:
            return "CT not in alignment map — review"
        if pd.isna(stage) or stage is None:
            return "No lifecycle stage assigned"
        return ("Expected"
                if stage in CT_LIFECYCLE_EXPECTED[ct]
                else "Unexpected — review")

    obt["ct_expected_lifecycle_stages"] = obt["gold_control"].map(
        lambda ct: (
            " | ".join(CT_LIFECYCLE_EXPECTED.get(str(ct).strip(), []))
            if str(ct).strip() not in CT_BROAD
            else "Applies across all lifecycle stages"
        )
    )
    obt["ct_lifecycle_alignment"] = obt.apply(get_alignment, axis=1)

    # ── Derived: duplicate detection ──────────────────────────────────────────
    # Level 1: same CT + same process (strongest signal — same CT on same activity)
    if "l3_process_UUID" in obt.columns and "gold_control" in obt.columns:
        obt["controls_on_same_process_and_ct"] = (
            obt.groupby(["gold_control", "l3_process_UUID"])
            ["Control_ID"].transform("nunique")
        )
        obt["process_level_duplicate"] = obt["controls_on_same_process_and_ct"] > 1

    # Level 2: same CT + same category + same stage (broader context signal)
    if all(c in obt.columns for c in
           ["gold_control","process_category","process_lifecycle_stage"]):
        obt["controls_in_same_ct_category_stage"] = (
            obt.groupby(["gold_control","process_category","process_lifecycle_stage"])
            ["Control_ID"].transform("nunique")
        )

    # ── Derived: sole control ─────────────────────────────────────────────────
    if "controls_on_same_process_and_ct" in obt.columns:
        obt["sole_control_for_process_ct"] = (
            obt["controls_on_same_process_and_ct"] == 1
        )

    print(f"  OBT rows    : {len(obt):,}")
    print(f"  OBT columns : {len(obt.columns):,}")

    if "ct_lifecycle_alignment" in obt.columns:
        for label, count in obt["ct_lifecycle_alignment"].value_counts().items():
            print(f"    ct_lifecycle_alignment — {label}: {count:,}")
    if "process_level_duplicate" in obt.columns:
        print(f"  Process-level duplicate rows : {int(obt['process_level_duplicate'].sum()):,}")

    return obt

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

    ids1 = set(grp1["Control_ID"])
    ids2 = set(grp2["Control_ID"])
    ids3 = set(grp3["Control_ID"])

    chk("No overlap between groups",
        0, len(ids1 & ids2) + len(ids1 & ids3) + len(ids2 & ids3))
    chk("All controls assigned to exactly one group",
        len(controls), len(grp1) + len(grp2) + len(grp3))

    if "payment_categories" in grp1.columns:
        raw = set()
        for v in grp1["payment_categories"].dropna():
            for c in str(v).split(" | "):
                raw.add(c.strip())
        bad = raw - set(VALID_CATEGORIES) - {"Unclassified / Missing", ""}
        chk("All payment categories valid", "None",
            str(bad) if bad else "None")

    if "lifecycle_stages" in grp1.columns:
        raw = set()
        for v in grp1["lifecycle_stages"].dropna():
            for s in str(v).split(" | "):
                raw.add(s.strip())
        bad = raw - set(VALID_LIFECYCLE_STAGES) - {""}
        chk("All lifecycle stages valid", "None",
            str(bad) if bad else "None")

    variant_left = (procs["process_lifecycle_stage"] == "Posting, Accounting & Detection").sum()
    chk("Stage normalisation complete", 0, variant_left)
    chk("Group 2 has no payment links", 0, len(ids2 & ids1))
    chk("Group 3 has no Holo linkage",  0, len(ids3 & (ids1 | ids2)))

    return pd.DataFrame(checks)

# ─────────────────────────────────────────────────────────────────────────────
#  SUMMARY
# ─────────────────────────────────────────────────────────────────────────────

def build_summary(controls, grp1, grp1_detail, grp2, grp2_detail, grp3, obt):
    total = len(grp1) + len(grp2) + len(grp3)
    rows  = [
        ("── POPULATION ───────────────────────────────────────", ""),
        ("Total payment controls",                  len(controls)),
        ("Group 1  Linked to payment processes",    len(grp1)),
        ("  L3 direct",
         (grp1["link_type"] == "L3 direct").sum()
          if "link_type" in grp1.columns else ""),
        ("  L2 fallback (Low confidence)",
         (grp1["link_type"] == "L2 fallback (Low confidence)").sum()
          if "link_type" in grp1.columns else ""),
        ("Group 1 detail rows",                     len(grp1_detail)),
        ("Group 2  Linked non-payment Holo only",   len(grp2)),
        ("Group 2 detail rows",                     len(grp2_detail)),
        ("Group 3  Not linked to any Holo process", len(grp3)),
        ("Total (sum of groups)",                   total),
    ]
    if "payment_process_count" in grp1.columns:
        avg = grp1["payment_process_count"].astype(float).mean()
        rows.append(("Avg payment processes per Group 1 control", f"{avg:.1f}"))

    rows += [("── ONE BIG TABLE ─────────────────────────────────────", ""),
             ("OBT rows",    len(obt)),
             ("OBT columns", len(obt.columns))]

    if "process_level_duplicate" in obt.columns:
        rows.append(("OBT process-level duplicate rows",
                     int(obt["process_level_duplicate"].sum())))
    if "ct_lifecycle_alignment" in obt.columns:
        unexpected = int((obt["ct_lifecycle_alignment"] == "Unexpected — review").sum())
        rows.append(("OBT rows with unexpected CT-lifecycle alignment", unexpected))

    return pd.DataFrame(rows, columns=["Metric", "Value"])

# ─────────────────────────────────────────────────────────────────────────────
#  WRITE OUTPUT
# ─────────────────────────────────────────────────────────────────────────────

def write_outputs(grp1, grp1_detail, grp2, grp2_detail,
                  grp3, obt, summary, validation):
    print(f"\n  Writing to:\n  {OUTPUT_FILE}")
    Path(OUTPUT_FILE).parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as w:
        summary.to_excel(     w, index=False, sheet_name="summary")
        grp1.to_excel(        w, index=False, sheet_name="linked_payment")
        grp1_detail.to_excel( w, index=False, sheet_name="linked_payment_detail")
        grp2.to_excel(        w, index=False, sheet_name="linked_non_payment")
        grp2_detail.to_excel( w, index=False, sheet_name="linked_non_payment_detail")
        grp3.to_excel(        w, index=False, sheet_name="not_linked")
        obt.to_excel(         w, index=False, sheet_name="one_big_table")
        validation.to_excel(  w, index=False, sheet_name="validation_checks")

    print(f"\n  Sheets:")
    print(f"    summary                    key metrics")
    print(f"    linked_payment             {len(grp1):,} controls (one row per control)")
    print(f"    linked_payment_detail      {len(grp1_detail):,} rows (one per control-process pair)")
    print(f"    linked_non_payment         {len(grp2):,} controls (one row per control)")
    print(f"    linked_non_payment_detail  {len(grp2_detail):,} rows (one per linkage row)")
    print(f"    not_linked                 {len(grp3):,} controls")
    print(f"    one_big_table              {len(obt):,} rows, {len(obt.columns):,} columns")
    print(f"    validation_checks          integrity checks")

# ─────────────────────────────────────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main():
    print("=" * 70)
    print("  Payment Control Linkage Analysis")
    print("=" * 70)

    controls, linkage, procs                        = load_data()
    controls, linkage, procs                        = normalise(controls, linkage, procs)
    grp1, grp1_detail, grp2, grp2_detail, grp3     = build_populations(controls, linkage, procs)
    obt                                             = build_one_big_table(grp1_detail, controls, procs)
    validation                                      = run_validation(controls, grp1, grp2, grp3, procs)
    summary                                         = build_summary(controls, grp1, grp1_detail,
                                                                    grp2, grp2_detail, grp3, obt)
    write_outputs(grp1, grp1_detail, grp2, grp2_detail, grp3, obt, summary, validation)

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
