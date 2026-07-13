"""
payment_control_linkage_analysis.py
─────────────────────────────────────────────────────────────────────────────
Payments Controls PoC — Control-to-Process Linkage & Coverage Analysis

Output sheets (10):
  summary                    - key metrics using distinct Control_ID counts
  linked_payment             - one row per control with payment linkage
  linked_payment_detail      - one row per control-process pair (L3 only)
  linked_non_payment         - one row per control linked to non-payment Holo
  linked_non_payment_detail  - one row per linkage row for non-payment controls
  not_linked                 - controls with no Holo linkage
  one_big_table              - denormalised analytical layer (Population A)
  misplaced_control_candidates - filtered OBT: misplaced_candidate_flag = True
  top_misplaced_examples     - top 20 ranked misplaced candidates
  validation_checks          - automated integrity checks

Run:
  python payment_control_linkage_analysis.py
"""

from pathlib import Path
from collections import Counter
import pandas as pd

# ─────────────────────────────────────────────────────────────────────────────
#  CONFIG — update file paths before running
# ─────────────────────────────────────────────────────────────────────────────

FILE_CONTROLS  = r"Z:\path\to\juno_payment_controls_gold.xlsx"
FILE_LINKAGE   = r"Z:\path\to\juno_holo_deterministic_linkage.xlsx"
FILE_PROCESSES = r"Z:\path\to\holocentric_payment_processes.xlsx"
OUTPUT_FILE    = r"Z:\path\to\payment_control_linkage_analysis.xlsx"

# ─────────────────────────────────────────────────────────────────────────────
#  VALID VALUE LISTS
# ─────────────────────────────────────────────────────────────────────────────

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

# ─────────────────────────────────────────────────────────────────────────────
#  GOLD CONTROL TITLES  (CT1-CT28)
# ─────────────────────────────────────────────────────────────────────────────

CT_TITLES = {
    "CT1":  "Validation of Human-Entered Data at Input",
    "CT2":  "Payment processing error detection",
    "CT3":  "Early Identification of Duplications and Processing Errors",
    "CT4":  "Payment processing interface and batch error resolution",
    "CT5":  "Incident response",
    "CT6":  "Master/Reference data input validation",
    "CT7":  "Service provider ongoing review",
    "CT8":  "Service provider onboarding",
    "CT9":  "Change management testing",
    "CT10": "Critical service chain mapping and risk identification",
    "CT11": "Rollback plans",
    "CT12": "System recovery capability",
    "CT13": "Business continuity plan",
    "CT14": "Logging and monitoring",
    "CT15": "Secure IT Design",
    "CT16": "Vulnerability Management",
    "CT17": "Access - Provision / deprovision (EIA)",
    "CT18": "Access - Monitoring",
    "CT19": "Patch Management",
    "CT20": "Physical Security Controls",
    "CT21": "Access - Privileged users",
    "CT22": "Mistaken Internet Payment Reports",
    "CT23": "Provision of Confirmations and Notifications",
    "CT24": "Treatment of Unauthorised and Disputed Transactions",
    "CT25": "Provision of Confirmations and Notifications",
    "CT26": "Regulatory horizon scanning",
    "CT27": "Records retention",
    "CT28": "Crisis Management planning and testing",
}

# ─────────────────────────────────────────────────────────────────────────────
#  CATEGORY-SPECIFIC CT-TO-LIFECYCLE STAGE MAPPING
# ─────────────────────────────────────────────────────────────────────────────
# Source: SME-validated payment control architecture rules.
# Key: uses gold_control_code (CT1..CT28), NOT raw numeric gold_control.
# Outer key = payment category  |  Inner key = CT code  |  Value = expected stages
# CTs absent from a category's inner dict are NOT expected for that category.

_S1 = ["Initiation & Validation & Authorisation"]
_S2 = ["Execution & Early Processing Assurance"]
_S3 = ["Clearing / Settlement"]
_S4 = ["Posting & Accounting, Detection"]
_S5 = ["Notification & Reporting"]
_S6 = ["Incident response, disputes, recovery followups"]

# Base: applies identically to all five categories for stages 1-4 and 6
_BASE = {
    "CT1":  _S1, "CT6":  _S1,
    "CT3":  _S2, "CT4":  _S2,
    "CT14": _S3,
    "CT2":  _S4, "CT27": _S4,
    "CT5":  _S6,
}

CT_CATEGORY_STAGE_EXPECTED = {
    "Customer to Customer": {
        **_BASE, "CT22": _S5, "CT23": _S5, "CT24": _S6,
    },
    "Customer to Institution": {
        **_BASE, "CT22": _S5, "CT23": _S5, "CT24": _S6,
    },
    "Institution to Customer": {
        **_BASE, "CT23": _S5,
        # CT22 and CT24 not expected (customer-initiated concepts only)
    },
    "Institution to Institution": {
        **_BASE,
        # No notification stage controls; CT22/CT23/CT24 not expected
    },
    "Supplier / Contractor / Employee Payments": {
        **_BASE,
        # No notification stage controls; CT22/CT23/CT24 not expected
    },
}

# System controls: technology/security — apply across all stages and categories
CT_SYSTEM = {"CT8", "CT12", "CT14", "CT15", "CT16", "CT18", "CT19", "CT21"}

# Governance controls: pervasive enterprise controls — all stages and categories
CT_GOVERNANCE = {
    "CT7", "CT9", "CT10", "CT11", "CT13", "CT17", "CT20", "CT25", "CT26"
}

CT_BROAD = CT_SYSTEM | CT_GOVERNANCE   # All broad CTs combined

# JUNO columns with analytical value to carry into the OBT
JUNO_OBT_COLS = [
    "Control_ID", "CTRL_NAME",
    "CTRL_NATRE", "CTRL_STUS", "CTRL_KEY_CONTRL", "CTRL_FREQ",
    "CTRL_TYP", "CTRL_ASSESS_RTNG", "CTRL_OE_RTNG", "CTRL_DE_RTNG",
    "CTRL_CTGRY_1", "CTRL_CTGRY_2", "CTRL_CATEGORY",
    "CTRL_DESC", "CTRL_EVDNCD", "CTRL_MNTRD",
    "CTRL_FLDR", "CTRL_FLDR_LVL_2", "COMMON_CTRL_REFERENCE",
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
    case-insensitive match against valid stage strings."""
    if is_empty(val):
        return None
    s = str(val).strip()
    if s in LIFECYCLE_NORMALISATION:
        return LIFECYCLE_NORMALISATION[s]
    sl = s.lower()
    for key, canonical in LIFECYCLE_NORMALISATION.items():
        if key.lower() == sl:
            return canonical
    for valid in VALID_LIFECYCLE_STAGES:
        if valid.lower() == sl:
            return valid
    return s

def norm_gold_control_code(val):
    """
    Convert raw gold_control to CT code format.
    Handles:  1 -> CT1   |   "28" -> CT28   |   "CT1" -> CT1
    Returns None for unrecognised values.
    """
    if is_empty(val):
        return None
    s = str(val).strip()
    # Already CT-prefixed (case-insensitive)
    if s.upper().startswith("CT"):
        suffix = s[2:].strip()
        try:
            n = int(suffix)
            if 1 <= n <= 28:
                return f"CT{n}"
        except ValueError:
            pass
        return None
    # Numeric
    try:
        n = int(float(s))
        if 1 <= n <= 28:
            return f"CT{n}"
    except (ValueError, TypeError):
        pass
    return None

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
    print(f"  Controls file rows : {len(controls):,}")
    print(f"  Linkage rows       : {len(linkage):,}")
    print(f"  Processes rows     : {len(procs):,}")
    return controls, linkage, procs

# ─────────────────────────────────────────────────────────────────────────────
#  NORMALISE
# ─────────────────────────────────────────────────────────────────────────────

def normalise(controls, linkage, procs):
    print("\n  Normalising...")

    # ── Controls: gold_control normalisation (THE PRIMARY BUG FIX) ───────────
    controls["Control_ID"]      = controls["Control_ID"].str.strip()
    controls["gold_control_raw"]= controls["gold_control"].copy()   # preserve original
    controls["gold_control_code"]= controls["gold_control"].apply(norm_gold_control_code)
    controls["gold_control_title"]= controls["gold_control_code"].map(CT_TITLES)

    # Duplicate Control_ID diagnostics
    dup_ids = controls[controls["Control_ID"].duplicated(keep=False)]["Control_ID"].unique()
    n_distinct = controls["Control_ID"].nunique()
    n_raw      = len(controls)
    if len(dup_ids):
        print(f"  WARNING: {len(dup_ids)} duplicate Control_ID(s) found: {list(dup_ids)[:10]}")
    print(f"  Distinct Control_IDs : {n_distinct:,}  (file rows: {n_raw:,})")

    # Deduplicate controls on Control_ID (keep first occurrence)
    controls = controls.drop_duplicates(subset=["Control_ID"], keep="first")
    print(f"  After dedup         : {len(controls):,} controls")

    # ── Linkage normalisation ─────────────────────────────────────────────────
    linkage["CTRL_ID"]           = linkage["CTRL_ID"].str.strip()
    linkage["l3_activity_uuid"]  = linkage["l3_activity_uuid"].apply(norm_uuid)
    linkage["l2_process_uuid"]   = linkage["l2_process_uuid"].apply(norm_uuid)

    # ── Processes normalisation ───────────────────────────────────────────────
    procs["l3_process_UUID"]         = procs["l3_process_UUID"].apply(norm_uuid)
    procs["l2_process_UUID"]         = procs["l2_process_UUID"].apply(norm_uuid)
    procs["process_category"]        = procs["process_category"].apply(norm_category)
    procs["process_lifecycle_stage"] = procs["process_lifecycle_stage"].apply(norm_stage)

    # Rename product/service column (contains a slash that causes issues)
    prod_col = next((c for c in procs.columns
                     if "product" in c.lower() and "service" in c.lower()), None)
    if prod_col and prod_col != "l3_activity_product_service":
        procs = procs.rename(columns={prod_col: "l3_activity_product_service"})

    # Report normalisation stats
    null_codes = controls["gold_control_code"].isna().sum()
    if null_codes:
        print(f"  WARNING: {null_codes} control(s) could not be mapped to a CT code")

    return controls, linkage, procs

# ─────────────────────────────────────────────────────────────────────────────
#  BUILD POPULATIONS
# ─────────────────────────────────────────────────────────────────────────────

def build_populations(controls, linkage, procs):
    print("\n  Building populations...")

    # Split linkage by type
    lk_l3 = linkage[linkage["l3_activity_uuid"].notna()].copy()
    lk_l2 = linkage[
        linkage["l3_activity_uuid"].isna() &
        linkage["l2_process_uuid"].notna()
    ].copy()
    print(f"  L3 linkage rows    : {len(lk_l3):,}")
    print(f"  L2-only rows       : {len(lk_l2):,}")

    # Key JUNO columns for all output sheets
    ctrl_cols = [c for c in
                 ["Control_ID","CTRL_NAME","gold_control_raw",
                  "gold_control_code","gold_control_title"] +
                 [c for c in JUNO_OBT_COLS if c not in
                  ["Control_ID","CTRL_NAME"]]
                 if c in controls.columns]

    # Holo columns for detail views
    holo_cols = [c for c in [
        "l3_process_UUID","l2_process_UUID","l2_process_name",
        "l3_activity_name","l3_activity_description",
        "process_category","process_lifecycle_stage",
        "l3_activity_product_service",
    ] if c in procs.columns]

    # ── L3 primary join ───────────────────────────────────────────────────────
    l3_payment = (
        lk_l3
        .merge(procs[holo_cols],
               left_on="l3_activity_uuid", right_on="l3_process_UUID",
               how="inner")
        .drop_duplicates(subset=["CTRL_ID","l3_process_UUID"])
    )
    ctrl_l3_pay = set(l3_payment["CTRL_ID"])
    print(f"  Controls L3 payment match : {len(ctrl_l3_pay):,}")

    # ── L2 fallback (zero L3 payment matches only) ────────────────────────────
    needs_fb  = set(controls["Control_ID"]) - ctrl_l3_pay
    l2_elig   = lk_l2[lk_l2["CTRL_ID"].isin(needs_fb)]
    l2_payment = pd.DataFrame()
    if not l2_elig.empty:
        l2_payment = (
            l2_elig
            .merge(procs[holo_cols],
                   left_on="l2_process_uuid", right_on="l2_process_UUID",
                   how="inner")
        )
    ctrl_l2_pay = set(l2_payment["CTRL_ID"]) if not l2_payment.empty else set()
    print(f"  Controls L2 fallback match: {len(ctrl_l2_pay):,}")

    controls_any_holo = set(linkage["CTRL_ID"].dropna().str.strip().unique())
    grp1_ids = ctrl_l3_pay | ctrl_l2_pay

    # Aggregate helper: one row per control
    def agg_one_row(df, label):
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
            .rename(columns={"CTRL_ID":"Control_ID"})
        )
        a["link_type"] = label
        return a

    agg_l3 = agg_one_row(l3_payment, "L3 direct")
    agg_l2 = agg_one_row(l2_payment, "L2 fallback (Low confidence)")

    grp1 = controls[ctrl_cols].merge(
        pd.concat([agg_l3, agg_l2], ignore_index=True),
        on="Control_ID", how="inner"
    )

    # Detail helper: one row per control-process pair
    def enrich(df, label):
        if df.empty:
            return pd.DataFrame()
        d = controls[ctrl_cols].merge(
            df.rename(columns={"CTRL_ID":"Control_ID"}),
            on="Control_ID", how="inner"
        )
        d["link_type"] = label
        return d

    grp1_detail = pd.concat(
        [enrich(l3_payment, "L3 direct"),
         enrich(l2_payment, "L2 fallback (Low confidence)")],
        ignore_index=True
    )

    # Group 2: Holo-linked but no payment match
    grp2_ids = controls_any_holo - grp1_ids
    grp2 = controls[ctrl_cols][controls["Control_ID"].isin(grp2_ids)].copy()
    grp2["note"] = ("Linked to Holo processes but none "
                    "appear in the payment process inventory.")

    lk_avail = [c for c in [
        "CTRL_ID","l3_activity_uuid","l3_activity_id",
        "l2_process_uuid","l2_process_id","link_level",
        "bus_unit_bcrm_id","BUS_UNIT_FLDR_LVL_3","BUS_UNIT_FLDR_LVL_4",
        "holo_value_stream","vcm_library_name",
    ] if c in linkage.columns]
    grp2_detail = controls[ctrl_cols].merge(
        linkage[linkage["CTRL_ID"].isin(grp2_ids)][lk_avail]
        .rename(columns={"CTRL_ID":"Control_ID"}),
        on="Control_ID", how="inner"
    )
    grp2_detail["note"] = "Holo process NOT in payment inventory"

    # Group 3: no Holo linkage at all
    grp3_ids = set(controls["Control_ID"]) - controls_any_holo
    grp3 = controls[ctrl_cols][controls["Control_ID"].isin(grp3_ids)].copy()
    grp3["note"] = "No rows in Holo linkage file."

    total = len(grp1) + len(grp2) + len(grp3)
    print(f"\n  Group 1 — linked payment   : {len(grp1):,}")
    print(f"  Group 1 detail rows        : {len(grp1_detail):,}")
    print(f"  Group 2 — linked non-pay   : {len(grp2):,}")
    print(f"  Group 2 detail rows        : {len(grp2_detail):,}")
    print(f"  Group 3 — not linked       : {len(grp3):,}")
    print(f"  Total                      : {total:,}")

    return grp1, grp1_detail, grp2, grp2_detail, grp3

# ─────────────────────────────────────────────────────────────────────────────
#  ONE BIG TABLE
# ─────────────────────────────────────────────────────────────────────────────

def build_one_big_table(grp1_detail, controls, procs):
    """
    Denormalised analytical table built from Population A (L3 direct rows only
    for highest-confidence alignment analysis; L2 fallback rows included but
    clearly labelled as lower confidence in link_type and misplaced_candidate_strength).

    All CT alignment checks use gold_control_code (CT1..CT28), NOT raw gold_control.
    This is the fix for the root cause bug where numeric gold_control values
    failed to match CT-prefixed keys and every row was incorrectly flagged.
    """
    print("\n  Building One Big Table...")

    obt = grp1_detail.copy()

    # Ensure gold_control_code and gold_control_title are present
    # (they come from controls via grp1_detail; add if missing)
    for col in ["gold_control_raw","gold_control_code","gold_control_title"]:
        if col not in obt.columns and col in controls.columns:
            obt = obt.merge(controls[["Control_ID",col]], on="Control_ID", how="left")

    # Add remaining JUNO columns
    juno_extra = [c for c in JUNO_OBT_COLS
                  if c not in obt.columns and c in controls.columns]
    if juno_extra:
        obt = obt.merge(controls[["Control_ID"]+juno_extra],
                        on="Control_ID", how="left")

    # Add Holo context columns
    holo_ctx = [c for c in [
        "l3_activity_channels","l3_activity_customer_segments",
        "l3_activity_product_service","value_stream_name","vcm_library_name",
    ] if c in procs.columns and c not in obt.columns]
    if holo_ctx and "l3_process_UUID" in obt.columns:
        obt = obt.merge(procs[["l3_process_UUID"]+holo_ctx],
                        on="l3_process_UUID", how="left")

    # ── CT alignment — uses gold_control_code, NOT gold_control ───────────────

    def get_category_appropriateness(row):
        ct       = str(row.get("gold_control_code") or "").strip()
        category = str(row.get("process_category") or "").strip()
        if not ct:
            return "No CT code — review"
        if ct in CT_SYSTEM:
            return "Applicable (system control — all categories)"
        if ct in CT_GOVERNANCE:
            return "Applicable (governance control — all categories)"
        cat_map = CT_CATEGORY_STAGE_EXPECTED.get(category, {})
        if not cat_map:
            return "Category not in alignment map — review"
        return "Applicable" if ct in cat_map else f"Not expected for {category} — review"

    def get_expected_stages(row):
        ct       = str(row.get("gold_control_code") or "").strip()
        category = str(row.get("process_category") or "").strip()
        if not ct:
            return "No CT code"
        if ct == "CT14":
            return "Clearing / Settlement (stage role) + all stages (system role)"
        if ct in CT_SYSTEM:
            return "All lifecycle stages (system control)"
        if ct in CT_GOVERNANCE:
            return "All lifecycle stages (governance control)"
        cat_map = CT_CATEGORY_STAGE_EXPECTED.get(category, {})
        stages  = cat_map.get(ct)
        if stages:
            return " | ".join(stages)
        return "Not expected for this category"

    def get_alignment(row):
        ct       = str(row.get("gold_control_code") or "").strip()
        stage    = row.get("process_lifecycle_stage")
        category = str(row.get("process_category") or "").strip()
        if not ct:
            return "No CT code — cannot assess"
        # CT14 dual role
        if ct == "CT14":
            if stage == "Clearing / Settlement":
                return "Expected (stage-specific role)"
            return "Expected (system-wide role)"
        if ct in CT_SYSTEM:
            return "Expected (system control — all stages)"
        if ct in CT_GOVERNANCE:
            return "Expected (governance control — all stages)"
        cat_map = CT_CATEGORY_STAGE_EXPECTED.get(category, {})
        if ct not in cat_map:
            return f"CT not expected for {category} — review"
        if pd.isna(stage) or stage is None:
            return "No lifecycle stage assigned"
        return "Expected" if stage in cat_map[ct] else "Unexpected stage — review"

    obt["ct_appropriate_for_category"] = obt.apply(get_category_appropriateness, axis=1)
    obt["ct_expected_lifecycle_stages"] = obt.apply(get_expected_stages, axis=1)
    obt["ct_lifecycle_alignment"]       = obt.apply(get_alignment, axis=1)

    # ── Duplicate detection ───────────────────────────────────────────────────
    if "l3_process_UUID" in obt.columns and "gold_control_code" in obt.columns:
        obt["controls_on_same_process_and_ct"] = (
            obt.groupby(["gold_control_code","l3_process_UUID"])
            ["Control_ID"].transform("nunique")
        )
        obt["process_level_duplicate"] = obt["controls_on_same_process_and_ct"] > 1

        if "process_category" in obt.columns and "process_lifecycle_stage" in obt.columns:
            obt["controls_in_same_ct_category_stage"] = (
                obt.groupby(["gold_control_code","process_category",
                             "process_lifecycle_stage"])
                ["Control_ID"].transform("nunique")
            )
    if "controls_on_same_process_and_ct" in obt.columns:
        obt["sole_control_for_process_ct"] = (
            obt["controls_on_same_process_and_ct"] == 1
        )

    # ── Misplaced candidate classification ────────────────────────────────────
    def classify_misplaced(row):
        """
        Returns (flag, reason, strength).

        Strong:  L3 direct + non-broad CT + wrong category OR wrong stage
        Medium:  L3 direct + CT expected for category but no lifecycle stage
        Weak:    L2 fallback + wrong category or stage (lower confidence linkage)
        Not flagged: broad/system/governance CTs, or fully expected placement
        """
        ct        = str(row.get("gold_control_code") or "").strip()
        link_type = str(row.get("link_type") or "")
        cat_appr  = str(row.get("ct_appropriate_for_category") or "")
        alignment = str(row.get("ct_lifecycle_alignment") or "")
        is_l3     = (link_type == "L3 direct")
        is_broad  = ct in CT_BROAD

        if not ct:
            return False, "No CT code — cannot classify", None

        # Broad CTs: not flagged as misplaced by design
        if is_broad:
            return False, "", None

        cat_wrong   = "Not expected" in cat_appr
        stage_wrong = "Unexpected stage" in alignment

        if cat_wrong and is_l3:
            return (True,
                    "CT not expected for this payment category",
                    "Strong")
        if stage_wrong and is_l3:
            return (True,
                    "CT at unexpected lifecycle stage for this category",
                    "Strong")
        if cat_wrong and not is_l3:
            return (True,
                    "CT not expected for this payment category "
                    "(L2 fallback — lower confidence)",
                    "Weak")
        if stage_wrong and not is_l3:
            return (True,
                    "CT at unexpected lifecycle stage "
                    "(L2 fallback — lower confidence)",
                    "Weak")

        # Medium: CT is appropriate but lifecycle stage missing — can't confirm placement
        if is_l3 and not is_broad:
            if alignment == "No lifecycle stage assigned":
                return (True,
                        "CT expected for category but lifecycle stage not assigned — "
                        "cannot confirm correct placement",
                        "Medium")

        return False, "", None

    flags    = obt.apply(classify_misplaced, axis=1)
    obt["misplaced_candidate_flag"]     = flags.apply(lambda x: x[0])
    obt["misplaced_candidate_reason"]   = flags.apply(lambda x: x[1])
    obt["misplaced_candidate_strength"] = flags.apply(lambda x: x[2])

    # ── Reorder columns for reviewer readability ──────────────────────────────
    priority_cols = [
        "Control_ID","CTRL_NAME",
        "gold_control_raw","gold_control_code","gold_control_title",
        "link_type",
        "l2_process_name","l3_activity_name",
        "process_category","process_lifecycle_stage",
        "ct_appropriate_for_category","ct_expected_lifecycle_stages",
        "ct_lifecycle_alignment",
        "misplaced_candidate_flag","misplaced_candidate_reason",
        "misplaced_candidate_strength",
        "controls_on_same_process_and_ct","process_level_duplicate",
        "controls_in_same_ct_category_stage","sole_control_for_process_ct",
    ]
    front = [c for c in priority_cols if c in obt.columns]
    rest  = [c for c in obt.columns if c not in front]
    obt   = obt[front + rest]

    # Report
    print(f"  OBT rows    : {len(obt):,}")
    print(f"  OBT columns : {len(obt.columns):,}")

    n_strong = int(obt.get("misplaced_candidate_strength","").eq("Strong").sum()
                   if "misplaced_candidate_strength" in obt.columns else 0)
    n_medium = int(obt["misplaced_candidate_strength"].eq("Medium").sum()
                   if "misplaced_candidate_strength" in obt.columns else 0)
    n_weak   = int(obt["misplaced_candidate_strength"].eq("Weak").sum()
                   if "misplaced_candidate_strength" in obt.columns else 0)
    print(f"  Misplaced candidates — Strong: {n_strong}  Medium: {n_medium}  Weak: {n_weak}")

    if "ct_lifecycle_alignment" in obt.columns:
        for label, cnt in obt["ct_lifecycle_alignment"].value_counts().items():
            print(f"    alignment: {label}: {cnt:,}")

    return obt

# ─────────────────────────────────────────────────────────────────────────────
#  MISPLACED CANDIDATES SHEET
# ─────────────────────────────────────────────────────────────────────────────

def build_misplaced_candidates(obt):
    """
    Filtered OBT: misplaced_candidate_flag = True.
    Sorted: L3 direct first, Strong first, category mismatch before stage mismatch.
    """
    if "misplaced_candidate_flag" not in obt.columns:
        return pd.DataFrame()

    mc = obt[obt["misplaced_candidate_flag"] == True].copy()

    # Sort rank
    link_rank    = {"L3 direct": 0, "L2 fallback (Low confidence)": 1}
    strength_rank= {"Strong": 0, "Medium": 1, "Weak": 2}
    reason_rank  = {
        "CT not expected for this payment category": 0,
        "CT not expected for this payment category (L2 fallback — lower confidence)": 1,
        "CT at unexpected lifecycle stage for this category": 2,
        "CT at unexpected lifecycle stage (L2 fallback — lower confidence)": 3,
    }

    mc["_link_rank"]    = mc.get("link_type",pd.Series()).map(link_rank).fillna(9)
    mc["_strength_rank"]= mc.get("misplaced_candidate_strength",
                                  pd.Series()).map(strength_rank).fillna(9)
    mc["_reason_rank"]  = mc.get("misplaced_candidate_reason",
                                  pd.Series()).map(reason_rank).fillna(9)

    mc = mc.sort_values(["_link_rank","_strength_rank","_reason_rank"]).drop(
        columns=["_link_rank","_strength_rank","_reason_rank"])

    print(f"  Misplaced candidates sheet : {len(mc):,} rows")
    return mc

# ─────────────────────────────────────────────────────────────────────────────
#  TOP MISPLACED EXAMPLES SHEET
# ─────────────────────────────────────────────────────────────────────────────

def build_top_examples(mc, n=20):
    """
    Top N diverse misplaced examples.
    Deduplicates on Control_ID to ensure variety (different controls, not same
    control appearing many times due to multiple process links).
    Ranks: L3 direct > Strong > category mismatch > stage mismatch.
    Excludes broad CTs unless no other examples exist.
    """
    if mc.empty:
        return pd.DataFrame()

    # Prefer non-broad CTs
    is_broad_mask = mc.get("gold_control_code", pd.Series()).isin(CT_BROAD)
    non_broad = mc[~is_broad_mask]
    candidates = non_broad if not non_broad.empty else mc

    # One row per Control_ID (best-ranked row per control)
    top = candidates.groupby("Control_ID").first().reset_index()

    # Re-sort
    link_rank    = {"L3 direct": 0, "L2 fallback (Low confidence)": 1}
    strength_rank= {"Strong": 0, "Medium": 1, "Weak": 2}
    reason_rank  = {
        "CT not expected for this payment category": 0,
        "CT not expected for this payment category (L2 fallback — lower confidence)": 1,
        "CT at unexpected lifecycle stage for this category": 2,
        "CT at unexpected lifecycle stage (L2 fallback — lower confidence)": 3,
    }
    top["_lr"] = top.get("link_type","").map(link_rank).fillna(9)
    top["_sr"] = top.get("misplaced_candidate_strength","").map(strength_rank).fillna(9)
    top["_rr"] = top.get("misplaced_candidate_reason","").map(reason_rank).fillna(9)
    top = top.sort_values(["_lr","_sr","_rr"]).head(n).drop(columns=["_lr","_sr","_rr"])

    print(f"  Top examples sheet         : {len(top):,} rows")
    return top

# ─────────────────────────────────────────────────────────────────────────────
#  VALIDATION
# ─────────────────────────────────────────────────────────────────────────────

def run_validation(controls_raw, controls, grp1, grp2, grp3, procs, obt):
    print("\n  Running validation checks...")
    checks = []

    def chk(name, expected, actual, note=""):
        passed = str(expected) == str(actual)
        checks.append({
            "check": name, "expected": str(expected),
            "actual": str(actual),
            "pass":   "PASS" if passed else "FAIL",
            "note":   note,
        })
        print(f"  [{'PASS' if passed else 'FAIL'}] {name}")

    n_raw      = len(controls_raw)
    n_distinct = controls["Control_ID"].nunique()
    ids1 = set(grp1["Control_ID"])
    ids2 = set(grp2["Control_ID"])
    ids3 = set(grp3["Control_ID"])

    # Population integrity
    chk("No overlap between groups",
        0, len(ids1&ids2)+len(ids1&ids3)+len(ids2&ids3))
    chk("All distinct controls assigned to exactly one group",
        n_distinct, len(grp1)+len(grp2)+len(grp3))
    if n_raw != n_distinct:
        chk("Duplicate Control_IDs in source file detected",
            "Yes", "Yes",
            f"File has {n_raw} rows but {n_distinct} distinct IDs. "
            f"Deduplication applied (kept first row per Control_ID).")

    # gold_control_code populated
    null_codes = controls["gold_control_code"].isna().sum()
    chk("All controls have gold_control_code", 0, null_codes,
        "Controls missing CT code will not appear in alignment analysis.")

    # gold_control_code values are valid
    valid_cts = set(CT_TITLES.keys())
    if "gold_control_code" in controls.columns:
        bad_codes = set(controls["gold_control_code"].dropna().unique()) - valid_cts
        chk("All gold_control_code values are CT1-CT28", "None",
            str(bad_codes) if bad_codes else "None")

    # No raw numeric gold_control used in CT alignment (sanity check)
    if "gold_control_code" in obt.columns:
        numeric_in_obt = obt["gold_control_code"].dropna().apply(
            lambda x: str(x).isdigit()
        ).sum()
        chk("No raw numeric gold_control used in CT alignment", 0, numeric_in_obt,
            "If >0 the CT alignment logic is using unnormalised values.")

    # Lifecycle normalisation complete
    variant_left = (procs["process_lifecycle_stage"] == "Posting, Accounting & Detection").sum()
    chk("Stage normalisation complete", 0, variant_left)

    # Category values valid
    if "payment_categories" in grp1.columns:
        raw = set()
        for v in grp1["payment_categories"].dropna():
            for c in str(v).split(" | "):
                raw.add(c.strip())
        bad = raw - set(VALID_CATEGORIES) - {"Unclassified / Missing",""}
        chk("All payment categories valid", "None",
            str(bad) if bad else "None")

    # Sanity check: if ALL OBT rows are flagged as "CT not expected" that
    # indicates the alignment logic is using the wrong key (e.g. numeric vs CT code)
    if "ct_lifecycle_alignment" in obt.columns and len(obt) > 0:
        not_expected_count = obt["ct_lifecycle_alignment"].str.contains(
            "not expected", case=False, na=False
        ).sum()
        all_unexpected = (not_expected_count == len(obt))
        chk("Not all OBT rows flagged as CT-not-expected (alignment sanity)",
            "False", str(all_unexpected),
            "If True, the gold_control_code normalisation may have failed.")

    return pd.DataFrame(checks)

# ─────────────────────────────────────────────────────────────────────────────
#  SUMMARY
# ─────────────────────────────────────────────────────────────────────────────

def build_summary(controls_raw, controls, grp1, grp1_detail,
                  grp2, grp2_detail, grp3, obt, mc):
    n_distinct = controls["Control_ID"].nunique()
    total      = len(grp1) + len(grp2) + len(grp3)

    rows = [
        ("── CONTROL POPULATION ──────────────────────────────────",""),
        ("Controls file rows (before dedup)",   len(controls_raw)),
        ("Distinct Control_IDs (after dedup)",  n_distinct),
        ("── POPULATION GROUPS ───────────────────────────────────",""),
        ("Group 1 — Linked to payment processes",  len(grp1)),
        ("  of which L3 direct",
         (grp1["link_type"]=="L3 direct").sum()
          if "link_type" in grp1.columns else ""),
        ("  of which L2 fallback (Low confidence)",
         (grp1["link_type"]=="L2 fallback (Low confidence)").sum()
          if "link_type" in grp1.columns else ""),
        ("Group 1 detail rows (ctrl-process pairs)", len(grp1_detail)),
        ("Group 2 — Linked to non-payment Holo only", len(grp2)),
        ("Group 2 detail rows",                   len(grp2_detail)),
        ("Group 3 — Not linked to any Holo process", len(grp3)),
        ("Total (sum of groups)",                 total),
    ]

    # OBT metrics
    rows += [
        ("── ONE BIG TABLE ───────────────────────────────────────",""),
        ("OBT rows",    len(obt)),
        ("OBT columns", len(obt.columns)),
    ]

    if "misplaced_candidate_strength" in obt.columns:
        n_strong = int(obt["misplaced_candidate_strength"].eq("Strong").sum())
        n_medium = int(obt["misplaced_candidate_strength"].eq("Medium").sum())
        n_weak   = int(obt["misplaced_candidate_strength"].eq("Weak").sum())
        n_mc_ctrls = obt[obt["misplaced_candidate_flag"]==True][
            "Control_ID"].nunique() if "misplaced_candidate_flag" in obt.columns else 0
        rows += [
            ("Misplaced candidate rows total",          len(mc)),
            ("  Strong candidates",                     n_strong),
            ("  Medium candidates",                     n_medium),
            ("  Weak candidates",                       n_weak),
            ("Distinct controls with misplaced flag",   n_mc_ctrls),
        ]

    if "process_level_duplicate" in obt.columns:
        rows.append(("Process-level duplicate rows (same CT + process)",
                     int(obt["process_level_duplicate"].sum())))

    # Distributions
    if "gold_control_code" in obt.columns:
        rows += [("── BY GOLD CONTROL CODE ─────────────────────────────────","")]
        for ct, cnt in obt["gold_control_code"].value_counts().sort_index().items():
            title = CT_TITLES.get(str(ct),"")
            rows.append((f"  {ct} — {title}", int(cnt)))

    if "process_category" in obt.columns:
        rows += [("── BY PAYMENT CATEGORY ──────────────────────────────────","")]
        for cat, cnt in obt["process_category"].value_counts().items():
            rows.append((f"  {cat}", int(cnt)))

    if "process_lifecycle_stage" in obt.columns:
        rows += [("── BY LIFECYCLE STAGE ───────────────────────────────────","")]
        for stg, cnt in obt["process_lifecycle_stage"].value_counts().items():
            rows.append((f"  {stg}", int(cnt)))

    if "misplaced_candidate_reason" in obt.columns:
        rows += [("── BY MISPLACED REASON ──────────────────────────────────","")]
        for rsn, cnt in obt["misplaced_candidate_reason"].value_counts().items():
            if rsn:
                rows.append((f"  {rsn}", int(cnt)))

    return pd.DataFrame(rows, columns=["Metric","Value"])

# ─────────────────────────────────────────────────────────────────────────────
#  WRITE OUTPUT
# ─────────────────────────────────────────────────────────────────────────────

def write_outputs(grp1, grp1_detail, grp2, grp2_detail,
                  grp3, obt, mc, top, summary, validation):
    print(f"\n  Writing to:\n  {OUTPUT_FILE}")
    Path(OUTPUT_FILE).parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as w:
        summary.to_excel(    w, index=False, sheet_name="summary")
        grp1.to_excel(       w, index=False, sheet_name="linked_payment")
        grp1_detail.to_excel(w, index=False, sheet_name="linked_payment_detail")
        grp2.to_excel(       w, index=False, sheet_name="linked_non_payment")
        grp2_detail.to_excel(w, index=False, sheet_name="linked_non_payment_detail")
        grp3.to_excel(       w, index=False, sheet_name="not_linked")
        obt.to_excel(        w, index=False, sheet_name="one_big_table")
        mc.to_excel(         w, index=False, sheet_name="misplaced_control_candidates")
        top.to_excel(        w, index=False, sheet_name="top_misplaced_examples")
        validation.to_excel( w, index=False, sheet_name="validation_checks")

    print(f"\n  Sheets (10):")
    sheets = [
        ("summary",                    "key metrics and distributions"),
        ("linked_payment",             f"{len(grp1):,} controls"),
        ("linked_payment_detail",      f"{len(grp1_detail):,} control-process pairs"),
        ("linked_non_payment",         f"{len(grp2):,} controls"),
        ("linked_non_payment_detail",  f"{len(grp2_detail):,} linkage rows"),
        ("not_linked",                 f"{len(grp3):,} controls"),
        ("one_big_table",              f"{len(obt):,} rows, {len(obt.columns):,} cols"),
        ("misplaced_control_candidates",f"{len(mc):,} candidate rows"),
        ("top_misplaced_examples",     f"{len(top):,} examples"),
        ("validation_checks",          "integrity checks"),
    ]
    for name, desc in sheets:
        print(f"    {name:<35} {desc}")

# ─────────────────────────────────────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main():
    print("=" * 70)
    print("  Payment Control Linkage Analysis")
    print("=" * 70)

    controls_raw, linkage, procs            = load_data()
    controls, linkage, procs                = normalise(controls_raw.copy(),
                                                        linkage, procs)
    grp1, grp1_detail, grp2, grp2_detail, grp3 = build_populations(
        controls, linkage, procs)
    obt        = build_one_big_table(grp1_detail, controls, procs)
    mc         = build_misplaced_candidates(obt)
    top        = build_top_examples(mc)
    validation = run_validation(controls_raw, controls,
                                grp1, grp2, grp3, procs, obt)
    summary    = build_summary(controls_raw, controls,
                               grp1, grp1_detail, grp2, grp2_detail,
                               grp3, obt, mc)
    write_outputs(grp1, grp1_detail, grp2, grp2_detail,
                  grp3, obt, mc, top, summary, validation)

    failures = validation[validation["pass"] == "FAIL"]
    if not failures.empty:
        print(f"\n  WARNING: {len(failures)} validation check(s) FAILED:")
        for _, r in failures.iterrows():
            print(f"    - {r['check']}")
            print(f"      Expected: {r['expected']}  Got: {r['actual']}")
    else:
        print("\n  All validation checks passed.")

    print("\n" + "=" * 70)
    print("  Done.")
    print("=" * 70)

if __name__ == "__main__":
    main()
