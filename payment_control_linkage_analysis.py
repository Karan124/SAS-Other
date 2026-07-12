"""
payment_control_linkage_analysis.py
─────────────────────────────────────────────────────────────────────────────
Payments Controls PoC — Control-to-Process Linkage and Coverage Analysis

Three-table join: JUNO Controls → JUNO-Holo Linkage → Holocentric Payment Processes
L3-first linkage. L2 fallback only for controls with zero L3 payment links.

Input files (update CONFIG paths before running):
  FILE_CONTROLS  — juno_payment_controls_gold.xlsx
  FILE_LINKAGE   — juno_holo_deterministic_linkage.xlsx
  FILE_PROCESSES — holocentric_payment_processes.xlsx

Output:
  Single Excel workbook — update OUTPUT_FILE path.

Run:
  python payment_control_linkage_analysis.py
"""

from pathlib import Path
from collections import Counter
from itertools import product

import pandas as pd

# ─────────────────────────────────────────────────────────────────────────────
#  CONFIG
# ─────────────────────────────────────────────────────────────────────────────

FILE_CONTROLS  = r"Z:\path\to\juno_payment_controls_gold.xlsx"
FILE_LINKAGE   = r"Z:\path\to\juno_holo_deterministic_linkage.xlsx"
FILE_PROCESSES = r"Z:\path\to\holocentric_payment_processes.xlsx"
OUTPUT_FILE    = r"Z:\path\to\payment_control_linkage_analysis.xlsx"

EXPECTED_CONTROL_COUNT = 708

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

VALID_GOLD_CONTROLS = {f"CT{i}" for i in range(1, 29)}

# Known lifecycle stage variant strings mapped to canonical form
LIFECYCLE_NORMALISATION = {
    "Posting, Accounting & Detection":    "Posting & Accounting, Detection",
    "Posting & Accounting & Detection":   "Posting & Accounting, Detection",
    "Posting & Accounting and Detection": "Posting & Accounting, Detection",
}

UNCLASSIFIED_SIGNALS = {"0", "", "nan", "none", "null", "unclassified"}

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
    if is_empty(val):
        return None
    return str(val).strip().lower()

def norm_category(val):
    if is_empty(val):
        return "Unclassified / Missing"
    s = str(val).strip()
    if s.lower() in UNCLASSIFIED_SIGNALS:
        return "Unclassified / Missing"
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
    vals = sorted(str(v) for v in values if v is not None and str(v).strip() not in ("", "nan"))
    return " | ".join(vals)

def chk(checks, name, expected, actual, note=""):
    passed = (str(expected) == str(actual))
    checks.append({
        "check":    name,
        "expected": str(expected),
        "actual":   str(actual),
        "pass":     "PASS" if passed else "FAIL",
        "note":     note,
    })
    print(f"  [{'PASS' if passed else 'FAIL'}] {name}")

# ─────────────────────────────────────────────────────────────────────────────
#  STEP 1 — LOAD
# ─────────────────────────────────────────────────────────────────────────────

def load_data():
    print("\n  Loading data...")
    controls = clean_cols(pd.read_excel(FILE_CONTROLS,  dtype=str))
    linkage  = clean_cols(pd.read_excel(FILE_LINKAGE,   dtype=str))
    procs    = clean_cols(pd.read_excel(FILE_PROCESSES,  dtype=str))
    print(f"  Controls  : {len(controls):,} rows")
    print(f"  Linkage   : {len(linkage):,}  rows")
    print(f"  Processes : {len(procs):,}  rows")
    return controls, linkage, procs

# ─────────────────────────────────────────────────────────────────────────────
#  STEP 2 — NORMALISE
# ─────────────────────────────────────────────────────────────────────────────

def normalise(controls, linkage, procs):
    print("\n  Normalising...")

    controls["Control_ID"]  = controls["Control_ID"].str.strip()
    controls["gold_control"]= controls["gold_control"].str.strip()

    linkage["CTRL_ID"]          = linkage["CTRL_ID"].str.strip()
    linkage["l3_activity_uuid"] = linkage["l3_activity_uuid"].apply(norm_uuid)
    linkage["l2_process_uuid"]  = linkage["l2_process_uuid"].apply(norm_uuid)

    procs["l3_process_UUID"]         = procs["l3_process_UUID"].apply(norm_uuid)
    procs["l2_process_UUID"]         = procs["l2_process_UUID"].apply(norm_uuid)
    procs["process_category"]        = procs["process_category"].apply(norm_category)
    procs["process_lifecycle_stage"] = procs["process_lifecycle_stage"].apply(norm_stage)

    l2_only = (linkage["l3_activity_uuid"].isna() & linkage["l2_process_uuid"].notna()).sum()
    unc_cat  = (procs["process_category"] == "Unclassified / Missing").sum()
    null_stg = procs["process_lifecycle_stage"].isna().sum()
    print(f"  L2-only linkage rows      : {l2_only:,}")
    print(f"  Unclassified categories   : {unc_cat:,}")
    print(f"  Null lifecycle stages     : {null_stg:,}")
    return controls, linkage, procs

# ─────────────────────────────────────────────────────────────────────────────
#  STEP 3 — SPLIT LINKAGE
# ─────────────────────────────────────────────────────────────────────────────

def split_linkage(linkage):
    lk_l3 = linkage[linkage["l3_activity_uuid"].notna()].copy()
    lk_l2 = linkage[
        linkage["l3_activity_uuid"].isna() &
        linkage["l2_process_uuid"].notna()
    ].copy()
    print(f"\n  Linkage rows — L3: {len(lk_l3):,}  L2-only: {len(lk_l2):,}")
    return lk_l3, lk_l2

# ─────────────────────────────────────────────────────────────────────────────
#  STEP 4 — L3 PAYMENT DETAIL
# ─────────────────────────────────────────────────────────────────────────────

def build_l3_detail(controls, lk_l3, procs):
    print("\n  Building L3 payment detail...")

    # Controls -> L3 linkage rows
    ctrl_l3 = controls[["Control_ID","CTRL_NAME","gold_control"]].merge(
        lk_l3[["CTRL_ID","l3_activity_uuid","l2_process_uuid"]],
        left_on="Control_ID", right_on="CTRL_ID", how="left"
    )

    # Inner join to payment processes on L3 UUID
    detail = (
        ctrl_l3[ctrl_l3["l3_activity_uuid"].notna()]
        .merge(
            procs[[
                "l3_process_UUID","l2_process_UUID","l2_process_id",
                "l2_process_name","l3_activity_id","l3_activity_name",
                "l3_activity_description","process_category",
                "process_lifecycle_stage","value_stream_name",
            ]],
            left_on="l3_activity_uuid",
            right_on="l3_process_UUID",
            how="inner"
        )
        .drop_duplicates(subset=["Control_ID","l3_process_UUID"])
    )

    print(f"  L3 payment pairs (deduped) : {len(detail):,}")
    print(f"  Unique controls matched    : {detail['Control_ID'].nunique():,}")
    return detail

# ─────────────────────────────────────────────────────────────────────────────
#  STEP 5 — L2 FALLBACK
# ─────────────────────────────────────────────────────────────────────────────

def build_l2_fallback(controls, lk_l2, procs, l3_detail):
    print("\n  Building L2 fallback...")

    ctrl_with_l3 = set(l3_detail["Control_ID"])
    needs_fallback = set(controls["Control_ID"]) - ctrl_with_l3
    eligible = lk_l2[lk_l2["CTRL_ID"].isin(needs_fallback)].copy()

    if eligible.empty:
        print("  No L2 fallback required.")
        return pd.DataFrame()

    expanded = eligible.merge(
        procs[["l2_process_UUID","l3_process_UUID","l3_activity_name",
               "process_category","process_lifecycle_stage"]],
        left_on="l2_process_uuid", right_on="l2_process_UUID", how="inner"
    )

    if expanded.empty:
        print("  L2 fallback: no payment process matches found.")
        return pd.DataFrame()

    fb = (
        expanded.groupby("CTRL_ID")
        .agg(
            l2_payment_categories=("process_category",
                                   lambda x: pipe_join(x.dropna().unique())),
            l2_lifecycle_stages  =("process_lifecycle_stage",
                                   lambda x: pipe_join(x.dropna().unique())),
            l2_category_count    =("process_category","nunique"),
            l2_stage_count       =("process_lifecycle_stage","nunique"),
            l2_child_process_count=("l3_process_UUID","nunique"),
        )
        .reset_index()
        .rename(columns={"CTRL_ID":"Control_ID"})
    )
    fb["link_type"]  = "L2_fallback"
    fb["confidence"] = "Low"
    fb["note"] = (
        "Aggregated from all L3 children under the linked L2 process. "
        "Do not use for individual control-to-process attribution."
    )
    print(f"  L2 fallback controls : {len(fb):,}")
    return fb

# ─────────────────────────────────────────────────────────────────────────────
#  STEP 6 — POPULATION CLASSIFICATION
# ─────────────────────────────────────────────────────────────────────────────

def classify_populations(controls, lk_l3, l3_detail, l2_fallback):
    print("\n  Classifying populations...")

    total_l3 = (
        lk_l3.groupby("CTRL_ID")["l3_activity_uuid"].nunique()
        .reset_index().rename(columns={"CTRL_ID":"Control_ID",
                                       "l3_activity_uuid":"total_l3_holo_links"})
    )
    pay_l3 = (
        l3_detail.groupby("Control_ID")["l3_process_UUID"].nunique()
        .reset_index().rename(columns={"l3_process_UUID":"payment_l3_links"})
    )

    l2_ids = set(l2_fallback["Control_ID"]) if not l2_fallback.empty else set()

    pop = controls[["Control_ID","CTRL_NAME","CTRL_DESC","gold_control"]].copy()
    pop = pop.merge(total_l3, on="Control_ID", how="left")
    pop = pop.merge(pay_l3,   on="Control_ID", how="left")
    pop["total_l3_holo_links"] = pop["total_l3_holo_links"].fillna(0).astype(int)
    pop["payment_l3_links"]    = pop["payment_l3_links"].fillna(0).astype(int)
    pop["has_l2_fallback"]     = pop["Control_ID"].isin(l2_ids)

    LABELS = {
        "A1": "A1 — All L3 links to payment processes",
        "A2": "A2 — Mixed: some payment, some non-payment L3 links",
        "B":  "B  — Holocentric-linked but no payment process match",
        "C":  "C  — No Holocentric linkage",
    }

    def code(row):
        if row["payment_l3_links"] > 0:
            return "A1" if row["total_l3_holo_links"] == row["payment_l3_links"] else "A2"
        if row["total_l3_holo_links"] > 0 or row["has_l2_fallback"]:
            return "B"
        return "C"

    pop["population_code"]  = pop.apply(code, axis=1)
    pop["population_label"] = pop["population_code"].map(LABELS)

    c = pop["population_code"].value_counts()
    print(f"  A1={c.get('A1',0)}  A2={c.get('A2',0)}  B={c.get('B',0)}  C={c.get('C',0)}"
          f"  Total={len(pop)}")
    return pop

# ─────────────────────────────────────────────────────────────────────────────
#  STEP 7 — CONTROL PROFILE
# ─────────────────────────────────────────────────────────────────────────────

def build_profile(l3_detail, l2_fallback, population):
    print("\n  Building control profiles...")

    profile_a = pd.DataFrame()
    if not l3_detail.empty:
        profile_a = (
            l3_detail.groupby("Control_ID")
            .agg(
                payment_categories      =("process_category",
                                          lambda x: pipe_join(x.dropna().unique())),
                lifecycle_stages        =("process_lifecycle_stage",
                                          lambda x: pipe_join(x.dropna().unique())),
                primary_category        =("process_category",   safe_mode),
                primary_lifecycle_stage =("process_lifecycle_stage", safe_mode),
                payment_process_count   =("l3_process_UUID",    "nunique"),
                category_count          =("process_category",   "nunique"),
                stage_count             =("process_lifecycle_stage","nunique"),
                l2_processes_covered    =("l2_process_name",
                                          lambda x: pipe_join(x.dropna().unique())),
            )
            .reset_index()
        )
        profile_a["is_multi_category"] = profile_a["category_count"] > 1
        profile_a["is_multi_stage"]    = profile_a["stage_count"] > 1
        profile_a["link_type"]  = "L3_direct"
        profile_a["confidence"] = "High"

    profile_l2 = pd.DataFrame()
    if not l2_fallback.empty:
        profile_l2 = l2_fallback.rename(columns={
            "l2_payment_categories":"payment_categories",
            "l2_lifecycle_stages":  "lifecycle_stages",
            "l2_category_count":    "category_count",
            "l2_stage_count":       "stage_count",
        }).copy()
        profile_l2["primary_category"] = profile_l2["payment_categories"].apply(
            lambda x: x.split(" | ")[0] if x else None)
        profile_l2["primary_lifecycle_stage"] = profile_l2["lifecycle_stages"].apply(
            lambda x: x.split(" | ")[0] if x else None)
        profile_l2["payment_process_count"] = profile_l2["l2_child_process_count"]
        profile_l2["is_multi_category"]     = profile_l2["category_count"] > 1
        profile_l2["is_multi_stage"]        = profile_l2["stage_count"] > 1
        profile_l2["l2_processes_covered"]  = ""

    profile = pd.concat([profile_a, profile_l2], ignore_index=True)
    profile = profile.merge(
        population[["Control_ID","CTRL_NAME","gold_control",
                    "population_code","population_label",
                    "total_l3_holo_links","payment_l3_links"]],
        on="Control_ID", how="left"
    )
    print(f"  Profiles generated : {len(profile):,}")
    return profile

# ─────────────────────────────────────────────────────────────────────────────
#  STEP 8 — COVERAGE MATRICES
# ─────────────────────────────────────────────────────────────────────────────

def build_matrices(l3_detail, controls):
    print("\n  Building coverage matrices...")
    if l3_detail.empty:
        return {}

    det = l3_detail.merge(controls[["Control_ID","gold_control"]],
                          on="Control_ID", how="left")

    def pivot(df, rows, cols, idx, col):
        g = df[["Control_ID", rows, cols]].drop_duplicates()
        m = g.groupby([rows, cols])["Control_ID"].nunique().unstack(fill_value=0)
        m = m.reindex(index=idx, columns=col, fill_value=0)
        return m

    ct_order = sorted(VALID_GOLD_CONTROLS, key=lambda x: int(x[2:]))
    all_cats  = VALID_CATEGORIES + ["Unclassified / Missing"]

    mat1 = pivot(det, "process_category", "process_lifecycle_stage",
                 all_cats, VALID_LIFECYCLE_STAGES)
    mat2 = pivot(det, "gold_control", "process_category",
                 ct_order, all_cats)
    mat3 = pivot(det, "gold_control", "process_lifecycle_stage",
                 ct_order, VALID_LIFECYCLE_STAGES)

    print(f"  Cat×Stage {mat1.shape} | CT×Cat {mat2.shape} | CT×Stage {mat3.shape}")
    return {"cat_x_stage": mat1, "ct_x_category": mat2, "ct_x_stage": mat3}

# ─────────────────────────────────────────────────────────────────────────────
#  STEP 9 — GAP ANALYSIS
# ─────────────────────────────────────────────────────────────────────────────

def build_gaps(matrices):
    print("\n  Building gap analysis...")
    gaps = []
    ct_order = sorted(VALID_GOLD_CONTROLS, key=lambda x: int(x[2:]))

    if "cat_x_stage" in matrices:
        m = matrices["cat_x_stage"]
        for cat, stg in product(VALID_CATEGORIES, VALID_LIFECYCLE_STAGES):
            n = int(m.loc[cat, stg]) if (cat in m.index and stg in m.columns) else 0
            if n == 0:
                gaps.append({"gap_type":"Category × Stage",
                             "gold_control":"", "payment_category":cat,
                             "lifecycle_stage":stg, "control_count":0,
                             "description":f"No controls: {cat} — {stg}"})

    if "ct_x_category" in matrices:
        m = matrices["ct_x_category"]
        for ct, cat in product(ct_order, VALID_CATEGORIES):
            n = int(m.loc[ct, cat]) if (ct in m.index and cat in m.columns) else 0
            if n == 0:
                gaps.append({"gap_type":"CT × Category",
                             "gold_control":ct, "payment_category":cat,
                             "lifecycle_stage":"", "control_count":0,
                             "description":f"{ct} — no controls for {cat}"})

    if "ct_x_stage" in matrices:
        m = matrices["ct_x_stage"]
        for ct, stg in product(ct_order, VALID_LIFECYCLE_STAGES):
            n = int(m.loc[ct, stg]) if (ct in m.index and stg in m.columns) else 0
            if n == 0:
                gaps.append({"gap_type":"CT × Stage",
                             "gold_control":ct, "payment_category":"",
                             "lifecycle_stage":stg, "control_count":0,
                             "description":f"{ct} — no controls for {stg}"})

    g = pd.DataFrame(gaps)
    if not g.empty:
        print(f"  Cat×Stage gaps: {(g['gap_type']=='Category × Stage').sum()}"
              f"  CT×Cat gaps: {(g['gap_type']=='CT × Category').sum()}"
              f"  CT×Stage gaps: {(g['gap_type']=='CT × Stage').sum()}")
    return g

# ─────────────────────────────────────────────────────────────────────────────
#  STEP 10 — VALIDATION
# ─────────────────────────────────────────────────────────────────────────────

def run_validation(controls, l3_detail, l2_fallback, population, procs):
    print("\n  Running validation checks...")
    checks = []

    pc = population["population_code"].value_counts()
    total = len(population)

    chk(checks, "Total controls = 708", EXPECTED_CONTROL_COUNT, total)
    chk(checks, "Population A1+A2+B+C = 708", EXPECTED_CONTROL_COUNT,
        pc.get("A1",0)+pc.get("A2",0)+pc.get("B",0)+pc.get("C",0))

    if not l3_detail.empty:
        dup = l3_detail.duplicated(subset=["Control_ID","l3_process_UUID"]).sum()
        chk(checks, "No duplicate control-process pairs", 0, dup)

        bad_cats = (set(l3_detail["process_category"].dropna().unique())
                    - set(VALID_CATEGORIES) - {"Unclassified / Missing"})
        chk(checks, "All categories are valid", "None",
            str(bad_cats) if bad_cats else "None")

        bad_stgs = (set(l3_detail["process_lifecycle_stage"].dropna().unique())
                    - set(VALID_LIFECYCLE_STAGES))
        chk(checks, "All lifecycle stages are valid", "None",
            str(bad_stgs) if bad_stgs else "None")

    variant_left = (procs["process_lifecycle_stage"] == "Posting, Accounting & Detection").sum()
    chk(checks, "Stage normalisation complete", 0, variant_left)

    gold_vals = set(controls["gold_control"].dropna().str.strip().unique())
    bad_gcs   = gold_vals - VALID_GOLD_CONTROLS
    chk(checks, "All gold controls are CT1-CT28", "None",
        str(bad_gcs) if bad_gcs else "None")

    if not l2_fallback.empty and not l3_detail.empty:
        overlap = set(l2_fallback["Control_ID"]) & set(l3_detail["Control_ID"])
        chk(checks, "L2 fallback controls have zero L3 payment links", 0,
            len(overlap), "If >0, investigate overlap between L3 and L2 fallback.")

    return pd.DataFrame(checks)

# ─────────────────────────────────────────────────────────────────────────────
#  STEP 11 — SUMMARY
# ─────────────────────────────────────────────────────────────────────────────

def build_summary(population, l3_detail, l2_fallback, matrices, gaps):
    pc  = population["population_code"].value_counts()
    rows = []
    rows += [
        ("── POPULATION ──────────────────────────────", ""),
        ("Total controls in scope",             len(population)),
        ("A1 — All L3 links to payment",        pc.get("A1",0)),
        ("A2 — Mixed payment / non-payment L3", pc.get("A2",0)),
        ("B  — No payment process match",       pc.get("B", 0)),
        ("C  — No Holocentric linkage",         pc.get("C", 0)),
        ("Controls with L2 fallback",           len(l2_fallback) if not l2_fallback.empty else 0),
    ]
    if not l3_detail.empty:
        ppc = l3_detail.groupby("Control_ID")["l3_process_UUID"].nunique()
        mc  = l3_detail.groupby("Control_ID")["process_category"].nunique()
        ms  = l3_detail.groupby("Control_ID")["process_lifecycle_stage"].nunique()
        rows += [
            ("── L3 PAYMENT LINKAGE ──────────────────────", ""),
            ("Total L3 payment control-process pairs",  len(l3_detail)),
            ("Unique payment processes linked",         l3_detail["l3_process_UUID"].nunique()),
            ("Avg payment processes per control",       f"{ppc.mean():.1f}"),
            ("Controls with multiple categories",       int((mc > 1).sum())),
            ("Controls with multiple lifecycle stages", int((ms > 1).sum())),
        ]
    if "cat_x_stage" in matrices:
        m = matrices["cat_x_stage"]
        total_cells   = len(VALID_CATEGORIES) * len(VALID_LIFECYCLE_STAGES)
        covered_cells = int((m.loc[VALID_CATEGORIES, VALID_LIFECYCLE_STAGES] > 0).values.sum())
        rows += [
            ("── COVERAGE ─────────────────────────────────", ""),
            ("Cat × Stage combinations possible", total_cells),
            ("Cat × Stage combinations covered",  covered_cells),
            ("Cat × Stage gaps",                  total_cells - covered_cells),
        ]
    if not gaps.empty:
        rows += [
            ("CT × Category gaps", int((gaps["gap_type"]=="CT × Category").sum())),
            ("CT × Stage gaps",    int((gaps["gap_type"]=="CT × Stage").sum())),
        ]
    return pd.DataFrame(rows, columns=["Metric","Value"])

# ─────────────────────────────────────────────────────────────────────────────
#  STEP 12 — WRITE OUTPUT
# ─────────────────────────────────────────────────────────────────────────────

def write_outputs(summary, population, l3_detail, profile,
                  l2_fallback, matrices, gaps, validation):
    print(f"\n  Writing outputs to:\n  {OUTPUT_FILE}")
    Path(OUTPUT_FILE).parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as w:
        summary.to_excel(w,    index=False, sheet_name="summary")
        population.to_excel(w, index=False, sheet_name="ctrl_population")
        profile.to_excel(w,    index=False, sheet_name="ctrl_payment_profile")

        if not l3_detail.empty:
            l3_detail.to_excel(w, index=False, sheet_name="ctrl_payment_detail")
        if not l2_fallback.empty:
            l2_fallback.to_excel(w, index=False, sheet_name="ctrl_l2_fallback")

        for key, label in [
            ("cat_x_stage",   "coverage_cat_x_stage"),
            ("ct_x_category", "coverage_ct_x_category"),
            ("ct_x_stage",    "coverage_ct_x_stage"),
        ]:
            if key in matrices:
                matrices[key].reset_index().to_excel(w, index=False, sheet_name=label)

        if not gaps.empty:
            gaps.to_excel(w, index=False, sheet_name="gap_analysis")

        validation.to_excel(w, index=False, sheet_name="validation_checks")

    sheets = ["summary","ctrl_population","ctrl_payment_profile","ctrl_payment_detail",
              "ctrl_l2_fallback","coverage_cat_x_stage","coverage_ct_x_category",
              "coverage_ct_x_stage","gap_analysis","validation_checks"]
    print(f"  Sheets: {' | '.join(sheets)}")

# ─────────────────────────────────────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main():
    print("=" * 70)
    print("  Payment Control Linkage & Coverage Analysis")
    print("=" * 70)

    controls, linkage, procs = load_data()
    controls, linkage, procs = normalise(controls, linkage, procs)
    lk_l3, lk_l2            = split_linkage(linkage)
    l3_detail                = build_l3_detail(controls, lk_l3, procs)
    l2_fallback              = build_l2_fallback(controls, lk_l2, procs, l3_detail)
    population               = classify_populations(controls, lk_l3, l3_detail, l2_fallback)
    profile                  = build_profile(l3_detail, l2_fallback, population)
    matrices                 = build_matrices(l3_detail, controls)
    gaps                     = build_gaps(matrices)
    validation               = run_validation(controls, l3_detail, l2_fallback, population, procs)
    summary                  = build_summary(population, l3_detail, l2_fallback, matrices, gaps)

    write_outputs(summary, population, l3_detail, profile,
                  l2_fallback, matrices, gaps, validation)

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
