"""
build_holocentric_products.py
─────────────────────────────────────────────────────────────────────────────
Payments Controls PoC — Holocentric Product Cleaning and LLM Payload Preparation

Stage 1 of 2: Cleaning, catalogue creation, and payload preparation.
Stage 2 (LLM inference) is a separate script — this script does NOT call the LLM.

What this script does:
  1. Reads holocentric_payment_processes.xlsx
  2. Cleans product fields — removes BM code prefixes, deduplicates
  3. Merges two product columns into a single clean_product field
  4. Identifies processes where product data is missing
  5. Aggregates to one row per l3_process_UUID
  6. Builds a product catalogue from all populated product rows
  7. Generates four output files

Outputs (all in C:\\Users\\m061400\\ai-test\\big_table\\products):
  holocentric_products_cleaned.xlsx
  holocentric_missing_product_llm_payload.jsonl
  holocentric_missing_product_llm_payload.xlsx
  holocentric_product_cleaning_summary.xlsx

Run:
  python build_holocentric_products.py
"""

import json
import re
import sys
from pathlib import Path

import pandas as pd

# ─────────────────────────────────────────────────────────────────────────────
#  CONFIG
# ─────────────────────────────────────────────────────────────────────────────

INPUT_FILE  = Path(r"C:\Users\m061400\ai-test\big_table\holocentric_payment_processes.xlsx")
OUTPUT_DIR  = Path(r"C:\Users\m061400\ai-test\big_table\products")

# Expected input columns (validation)
EXPECTED_COLS = [
    "l3_process_UUID",
    "l2_process_UUID",
    "l2_process_id",
    "l2_process_name",
    "l2_process_description",
    "l3_activity_id",
    "l3_activity_name",
    "l3_activity_description",
    "l3_activity_channels",
    "l3_activity_customer_segments",
    "l3_activity_product/service",
    "l3_activity_component_products",
    "value_stream_name",
    "vcm_library_name",
    "vcm_library_type",
    "value_chain",
    "bcm",
    "task_name",
    "alphabet_app",
    "process_category",
    "process_lifecycle_stage",
    "payment_rationale",
]

# The two raw product columns
PRODUCT_COL_1 = "l3_activity_product/service"
PRODUCT_COL_2 = "l3_activity_component_products"

# Text fields to include in JSONL process_details
LLM_PROCESS_DETAIL_COLS = [
    "l3_process_UUID",
    "l2_process_UUID",
    "l2_process_id",
    "l2_process_name",
    "l2_process_description",
    "l3_activity_id",
    "l3_activity_name",
    "l3_activity_description",
    "l3_activity_channels",
    "l3_activity_customer_segments",
    "value_stream_name",
    "vcm_library_name",
    "vcm_library_type",
    "value_chain",
    "bcm",
    "task_name",
    "alphabet_app",
    "process_category",
    "process_lifecycle_stage",
    "payment_rationale",
]

# Text columns to aggregate with " | " delimiter (all text except product)
AGG_TEXT_COLS = [
    "l2_process_UUID",
    "l2_process_id",
    "l2_process_name",
    "l2_process_description",
    "l3_activity_id",
    "l3_activity_name",
    "l3_activity_description",
    "l3_activity_channels",
    "l3_activity_customer_segments",
    PRODUCT_COL_1,
    PRODUCT_COL_2,
    "value_stream_name",
    "vcm_library_name",
    "vcm_library_type",
    "value_chain",
    "bcm",
    "task_name",
    "alphabet_app",
    "process_category",
    "process_lifecycle_stage",
    "payment_rationale",
]

# Placeholder strings that count as missing product
MISSING_PLACEHOLDERS = {
    "n/a", "na", "none", "null", "unknown",
    "not applicable", "not available", "not specified",
    "tbd", "tbc", "-",
}

# BM prefix pattern: matches "BM 04.11.02 -" and variants
# Handles: hyphen (-), en-dash (–), em-dash (—), with/without spaces
BM_PREFIX_RE = re.compile(
    r"^BM\s+[\d]+\.[\d]+\.?[\d]*\s*[-\u2013\u2014]\s*",
    re.IGNORECASE,
)

# ─────────────────────────────────────────────────────────────────────────────
#  HELPER FUNCTIONS
# ─────────────────────────────────────────────────────────────────────────────

def safe_str(val) -> str:
    """Convert any value to a clean string. NaN/None → empty string."""
    if pd.isna(val):
        return ""
    s = str(val)
    # Replace tabs, carriage returns, newlines with spaces
    s = s.replace("\t", " ").replace("\r", " ").replace("\n", " ")
    # Collapse repeated whitespace
    s = re.sub(r" {2,}", " ", s)
    return s.strip()


def strip_bm_prefix(raw_value: str) -> str:
    """
    Remove BM code prefix from a single product value.
    'BM 04.11.02 - Margin Lending' → 'Margin Lending'
    Handles dash variants: - / – / —
    """
    return BM_PREFIX_RE.sub("", raw_value.strip()).strip()


def clean_product_string(raw: str) -> list[str]:
    """
    Parse a semicolon-delimited product string, strip BM prefixes,
    and return a list of unique clean product names.
    Empty or missing input returns [].
    """
    if not raw:
        return []
    parts = [strip_bm_prefix(p) for p in raw.split(";")]
    # Remove blanks
    return [p for p in parts if p]


def is_missing_product(clean_product: str) -> bool:
    """
    Return True if clean_product should be treated as missing.
    Covers: blank, whitespace-only, and known placeholder strings.
    """
    if not clean_product or not clean_product.strip():
        return True
    return clean_product.strip().lower() in MISSING_PLACEHOLDERS


def merge_product_lists(list1: list[str], list2: list[str]) -> list[str]:
    """
    Merge two product lists, deduplicating case-insensitively
    while preserving the first-seen readable form.
    """
    seen_lower = set()
    merged = []
    for item in list1 + list2:
        key = item.lower().strip()
        if key and key not in seen_lower:
            seen_lower.add(key)
            merged.append(item.strip())
    return merged


def agg_pipe(series: pd.Series) -> str:
    """Aggregate a series of strings with ' | ', dropping blanks."""
    unique_vals = []
    seen = set()
    for v in series:
        s = safe_str(v)
        if s and s not in seen:
            seen.add(s)
            unique_vals.append(s)
    return " | ".join(unique_vals)


def agg_comma(series: pd.Series) -> str:
    """Aggregate a series of comma-separated product strings, deduplicating."""
    all_products = []
    for v in series:
        s = safe_str(v)
        if s:
            for p in [x.strip() for x in s.split(",")]:
                if p:
                    all_products.append(p)
    merged = merge_product_lists(all_products, [])
    return ", ".join(merged)


# ─────────────────────────────────────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main():
    print("=" * 70)
    print("  Holocentric Product Cleaning & LLM Payload Preparation")
    print("=" * 70)

    # ── Validate input file exists ────────────────────────────────────────────
    if not INPUT_FILE.exists():
        print(f"\n  ERROR: Input file not found:\n  {INPUT_FILE}")
        sys.exit(1)

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    # ── Load input ────────────────────────────────────────────────────────────
    print(f"\n  Loading: {INPUT_FILE.name}")
    df = pd.read_excel(INPUT_FILE, engine="openpyxl", dtype=str)
    df.columns = df.columns.str.strip()
    print(f"  Rows loaded       : {len(df):,}")
    print(f"  Columns           : {len(df.columns)}")

    # ── Validate expected columns ─────────────────────────────────────────────
    missing_cols = [c for c in EXPECTED_COLS if c not in df.columns]
    if missing_cols:
        print(f"\n  ERROR: Missing expected columns:")
        for c in missing_cols:
            print(f"    - {c}")
        sys.exit(1)
    print(f"  Column validation : PASS (all {len(EXPECTED_COLS)} expected columns present)")

    # ── Normalise all text ────────────────────────────────────────────────────
    print("\n  Normalising text fields...")
    for col in df.columns:
        df[col] = df[col].apply(safe_str)

    # ── Handle blank l3_process_UUID ──────────────────────────────────────────
    blank_uuid = df["l3_process_UUID"].eq("").sum()
    if blank_uuid > 0:
        print(f"\n  WARNING: {blank_uuid:,} row(s) have blank l3_process_UUID "
              f"and will be excluded from all outputs.")
    df = df[df["l3_process_UUID"].ne("")].copy()
    print(f"  Rows with valid UUID: {len(df):,}")

    # ── Clean product fields ──────────────────────────────────────────────────
    print("\n  Cleaning product fields...")

    def build_clean_product(row) -> tuple[str, str]:
        """Returns (clean_product, product_source)"""
        raw1 = row[PRODUCT_COL_1]
        raw2 = row[PRODUCT_COL_2]
        list1 = clean_product_string(raw1)
        list2 = clean_product_string(raw2)
        merged = merge_product_lists(list1, list2)
        clean = ", ".join(merged)
        # Determine source
        has1 = bool(list1)
        has2 = bool(list2)
        if has1 and has2:
            source = "Both product/service and component_products"
        elif has1:
            source = "l3_activity_product/service"
        elif has2:
            source = "l3_activity_component_products"
        else:
            source = "Missing"
        return clean, source

    results = df.apply(build_clean_product, axis=1)
    df["clean_product"]  = results.apply(lambda x: x[0])
    df["product_source"] = results.apply(lambda x: x[1])
    df["has_product"]    = df["clean_product"].apply(
        lambda x: not is_missing_product(x))
    df["requires_llm_product_inference"] = ~df["has_product"]

    has_count  = df["has_product"].sum()
    miss_count = df["requires_llm_product_inference"].sum()
    print(f"  Rows with product : {has_count:,}")
    print(f"  Rows missing product: {miss_count:,}")

    # ── Aggregate to one row per l3_process_UUID ──────────────────────────────
    print("\n  Aggregating to one row per l3_process_UUID...")

    agg_dict = {}
    for col in AGG_TEXT_COLS:
        if col in df.columns:
            agg_dict[col] = agg_pipe
    # Product aggregation
    agg_dict["clean_product"]  = agg_comma
    agg_dict["product_source"] = "first"
    agg_dict["has_product"]    = "any"
    agg_dict["requires_llm_product_inference"] = lambda x: not x.any()

    aggregated = (
        df.groupby("l3_process_UUID", sort=False)
        .agg(agg_dict)
        .reset_index()
    )

    # Recompute has_product and missing flag post-aggregation
    aggregated["has_product"] = aggregated["clean_product"].apply(
        lambda x: not is_missing_product(x))
    aggregated["requires_llm_product_inference"] = ~aggregated["has_product"]
    aggregated["product_source"] = aggregated.apply(
        lambda r: r["product_source"] if r["has_product"] else "Missing",
        axis=1
    )

    print(f"  Aggregated rows   : {len(aggregated):,}")

    # Reorder columns: derived columns first
    derived_cols = [
        "l3_process_UUID",
        "clean_product",
        "has_product",
        "requires_llm_product_inference",
        "product_source",
    ]
    remaining = [c for c in aggregated.columns if c not in derived_cols]
    aggregated = aggregated[derived_cols + remaining]

    # ── Build product catalogue ───────────────────────────────────────────────
    print("\n  Building product catalogue...")

    all_products = []
    for val in aggregated.loc[aggregated["has_product"], "clean_product"]:
        for p in [x.strip() for x in val.split(",")]:
            if p and not is_missing_product(p):
                all_products.append(p)

    # Deduplicate case-insensitively, preserve readable form
    seen = {}
    for p in all_products:
        key = p.lower()
        if key not in seen:
            seen[key] = p
    product_catalogue = sorted(seen.values(), key=lambda x: x.lower())
    print(f"  Distinct products in catalogue: {len(product_catalogue):,}")

    # ── Rows requiring LLM inference ──────────────────────────────────────────
    missing_df = aggregated[aggregated["requires_llm_product_inference"]].copy()
    print(f"  Processes requiring LLM inference: {len(missing_df):,}")

    # ── Build JSONL payload ───────────────────────────────────────────────────
    print("\n  Building LLM payload...")

    jsonl_path = OUTPUT_DIR / "holocentric_missing_product_llm_payload.jsonl"
    records_written = 0

    with open(jsonl_path, "w", encoding="utf-8") as f:
        for _, row in missing_df.iterrows():
            # Build process_details from available columns
            process_details = {}
            for col in LLM_PROCESS_DETAIL_COLS:
                if col in row.index:
                    val = row[col]
                    process_details[col] = val if val and val != "" else None

            payload = {
                "l3_process_UUID": row["l3_process_UUID"],
                "task": (
                    "Infer the most likely product or products for this "
                    "Holocentric L3 payment process using only the allowed "
                    "product catalogue."
                ),
                "allowed_product_catalogue": product_catalogue,
                "process_details": process_details,
                "expected_response_format": {
                    "inferred_products": ["Product name from allowed_product_catalogue"],
                    "confidence": "High | Medium | Low",
                    "rationale": (
                        "Brief explanation based on the process description fields"
                    ),
                    "requires_human_review": True,
                },
            }
            f.write(json.dumps(payload, ensure_ascii=False) + "\n")
            records_written += 1

    print(f"  JSONL records written: {records_written:,}")

    # ── Build Excel review version of JSONL ───────────────────────────────────
    review_rows = []
    preview_limit = 30
    catalogue_preview = ", ".join(product_catalogue[:preview_limit])
    if len(product_catalogue) > preview_limit:
        catalogue_preview += f" ... (+{len(product_catalogue)-preview_limit} more)"

    for _, row in missing_df.iterrows():
        process_details = {}
        for col in LLM_PROCESS_DETAIL_COLS:
            if col in row.index:
                process_details[col] = row[col] if row[col] != "" else None

        payload = {
            "l3_process_UUID": row["l3_process_UUID"],
            "task": (
                "Infer the most likely product or products for this "
                "Holocentric L3 payment process using only the allowed "
                "product catalogue."
            ),
            "allowed_product_catalogue": product_catalogue,
            "process_details": process_details,
            "expected_response_format": {
                "inferred_products": ["Product name from allowed_product_catalogue"],
                "confidence": "High | Medium | Low",
                "rationale": "Brief explanation based on the process description fields",
                "requires_human_review": True,
            },
        }

        review_rows.append({
            "l3_process_UUID":               row["l3_process_UUID"],
            "l2_process_name":               row.get("l2_process_name", ""),
            "l3_activity_name":              row.get("l3_activity_name", ""),
            "l3_activity_description":       row.get("l3_activity_description", ""),
            "process_category":              row.get("process_category", ""),
            "process_lifecycle_stage":       row.get("process_lifecycle_stage", ""),
            "alphabet_app":                  row.get("alphabet_app", ""),
            "allowed_product_catalogue_count":   len(product_catalogue),
            "allowed_product_catalogue_preview": catalogue_preview,
            "llm_payload_json":              json.dumps(payload, ensure_ascii=False),
        })

    review_df = pd.DataFrame(review_rows)

    # ── Build summary stats ───────────────────────────────────────────────────

    overall_summary = pd.DataFrame([
        ("Input file",                          INPUT_FILE.name),
        ("Input row count",                     len(df)),
        ("Unique l3_process_UUID count",        len(aggregated)),
        ("L3 processes with product",           int(aggregated["has_product"].sum())),
        ("L3 processes missing product",        int(aggregated["requires_llm_product_inference"].sum())),
        ("Distinct product catalogue count",    len(product_catalogue)),
    ], columns=["Metric", "Value"])

    # By process_category
    by_category = (
        aggregated.groupby("process_category")
        .agg(
            total=("l3_process_UUID", "count"),
            has_product=("has_product", "sum"),
            missing_product=("requires_llm_product_inference", "sum"),
        )
        .reset_index()
    )

    # By lifecycle_stage
    by_stage = (
        aggregated.groupby("process_lifecycle_stage")
        .agg(
            total=("l3_process_UUID", "count"),
            has_product=("has_product", "sum"),
            missing_product=("requires_llm_product_inference", "sum"),
        )
        .reset_index()
    )

    # By product_source
    by_source = (
        aggregated.groupby("product_source")
        .agg(count=("l3_process_UUID", "count"))
        .reset_index()
    )

    # Product catalogue sheet
    catalogue_df = pd.DataFrame(
        {"product_name": product_catalogue}
    )

    # ── Write outputs ─────────────────────────────────────────────────────────
    print("\n  Writing outputs...")

    # 1. holocentric_products_cleaned.xlsx
    cleaned_path = OUTPUT_DIR / "holocentric_products_cleaned.xlsx"
    aggregated.to_excel(cleaned_path, index=False, engine="openpyxl")
    print(f"  [1/4] {cleaned_path.name}  ({len(aggregated):,} rows)")

    # 2. JSONL (already written above)
    print(f"  [2/4] {jsonl_path.name}  ({records_written:,} records)")

    # 3. holocentric_missing_product_llm_payload.xlsx
    review_path = OUTPUT_DIR / "holocentric_missing_product_llm_payload.xlsx"
    review_df.to_excel(review_path, index=False, engine="openpyxl")
    print(f"  [3/4] {review_path.name}  ({len(review_df):,} rows)")

    # 4. holocentric_product_cleaning_summary.xlsx
    summary_path = OUTPUT_DIR / "holocentric_product_cleaning_summary.xlsx"
    with pd.ExcelWriter(summary_path, engine="openpyxl") as w:
        overall_summary.to_excel(w, index=False, sheet_name="overall_summary")
        by_category.to_excel(    w, index=False, sheet_name="by_process_category")
        by_stage.to_excel(       w, index=False, sheet_name="by_lifecycle_stage")
        by_source.to_excel(      w, index=False, sheet_name="by_product_source")
        catalogue_df.to_excel(   w, index=False, sheet_name="product_catalogue")
    print(f"  [4/4] {summary_path.name}")

    print(f"\n  All outputs written to:\n  {OUTPUT_DIR}")
    print("\n" + "=" * 70)
    print("  Done.")
    print("=" * 70)


if __name__ == "__main__":
    main()
