"""
run_product_inference_llm.py
─────────────────────────────────────────────────────────────────────────────
Payments Controls PoC — LLM Product Inference for Missing Holocentric Products

Stage 2 of 2. Requires Stage 1 outputs from build_holocentric_products.py.

Reads:
  holocentric_missing_product_llm_payload.jsonl  (one record per missing process)
  holocentric_products_cleaned.xlsx              (full cleaned dataset for merge)

Writes:
  holocentric_product_inference_results.xlsx     (LLM results, one row per UUID)
  holocentric_products_complete.xlsx             (cleaned + inferred, full dataset)
  holocentric_product_inference_checkpoint.jsonl (checkpoint for resume)

Before running (PowerShell):
  az account set --subscription 6c72e6c5-ed48-4030-b29c-34e2849c9288
  $env:REQUESTS_CA_BUNDLE = "C:\\Users\\m061400\\ai-test\\cacert.pem"
  $env:SSL_CERT_FILE      = "C:\\Users\\m061400\\ai-test\\cacert.pem"
  Remove-Item Env:AZURE_CA_BUNDLE -ErrorAction SilentlyContinue
  python run_product_inference_llm.py
  python run_product_inference_llm.py --force   # ignore checkpoint and restart
"""

import argparse
import json
import os
import random
import re
import time
from pathlib import Path

import pandas as pd
from azure.identity import DefaultAzureCredential, get_bearer_token_provider
from openai import AzureOpenAI

# ─────────────────────────────────────────────────────────────────────────────
#  CONFIG
# ─────────────────────────────────────────────────────────────────────────────

AZURE_ENDPOINT  = "https://ai.eng.azure.srv.westpac.com.au"
API_VERSION     = "2024-10-21"
MODEL           = "gpt-5.4"
REASONING_EFFORT = "medium"
MAX_COMPLETION_TOKENS = 8000

PRODUCTS_DIR = Path(r"C:\Users\m061400\ai-test\big_table\products")

JSONL_FILE    = PRODUCTS_DIR / "holocentric_missing_product_llm_payload.jsonl"
CLEANED_FILE  = PRODUCTS_DIR / "holocentric_products_cleaned.xlsx"
CHECKPOINT    = PRODUCTS_DIR / "holocentric_product_inference_checkpoint.jsonl"
RESULTS_FILE  = PRODUCTS_DIR / "holocentric_product_inference_results.xlsx"
COMPLETE_FILE = PRODUCTS_DIR / "holocentric_products_complete.xlsx"

BATCH_SIZE    = 3     # processes per LLM call
RETRY_COUNT   = 3
RETRY_BASE    = 5     # seconds base for exponential backoff
INTER_BATCH_SLEEP = 0.5

INPUT_PRICE_USD_PER_M  = 1.75
OUTPUT_PRICE_USD_PER_M = 14.00
AUD_USD_RATE           = 0.65

# ─────────────────────────────────────────────────────────────────────────────
#  IMPROVED SYSTEM PROMPT
# ─────────────────────────────────────────────────────────────────────────────

SYSTEM_PROMPT = """
You are a payment process product classification specialist for an Australian bank.

Your task is to infer which product(s) from a defined catalogue apply to
payment processes that have no existing product assignment.

CRITICAL RULES — apply to every response without exception:
1. You MUST only assign products that appear EXACTLY in the allowed_product_catalogue.
2. Do NOT invent, abbreviate, reword or paraphrase product names.
3. Copy product names character-for-character from the catalogue.
4. If no product can be confidently inferred, return an empty array.
5. Do not guess. Accuracy is more important than completeness.
6. requires_human_review must always be true.
""".strip()

# ─────────────────────────────────────────────────────────────────────────────
#  PROMPT RULES (user message template)
# ─────────────────────────────────────────────────────────────────────────────

PROMPT_RULES = """
CLASSIFICATION INSTRUCTIONS
────────────────────────────────────────────────────────────────────────

CONTEXT
All processes in this batch have no existing product assignment.
Use the evidence hierarchy below to determine the most applicable
product(s) from the allowed_product_catalogue.

EVIDENCE HIERARCHY
Apply sources in this order. Use the first source that provides clear,
direct evidence. Do not skip ahead to weaker sources if a stronger one
provides sufficient evidence.

  Priority 1 — Application / System  (field: alphabet_app)
  The application or platform name is the strongest indicator of product.
  If the application platform directly corresponds to a product in the
  catalogue, assign that product.
  Example: if alphabet_app contains "BPAY" and "BPAY" appears in the
  catalogue, assign it. If the app name maps to a product family covering
  multiple catalogue entries, assign all applicable entries.

  Priority 2 — Business Capability  (field: bcm)
  BCM (Business Capability Map) names often directly identify a product
  domain. Use to confirm or narrow to a specific product or product family.

  Priority 3 — Value Chain  (field: value_chain)
  Indicates the broader business area. Use to identify the product family
  when Priority 1 and 2 do not provide a specific product match.

  Priority 4 — Task Detail  (field: task_name)
  Explicit task descriptions may name products directly.
  Only use where task_name contains direct product references.

  Priority 5 — L2 Process Context  (fields: l2_process_name, l2_process_description)
  The parent L2 process may provide product context not visible in the L3
  activity fields. Use where the L2 name or description clearly identifies
  a product.

  Priority 6 — L3 Activity  (fields: l3_activity_name, l3_activity_description)
  Use only as a last resort. Only where the activity name or description
  explicitly names or directly and unambiguously implies a product that
  exists in the catalogue.

MULTI-PRODUCT MAPPING
Assign multiple products where the evidence clearly supports more than one:
  - Application serves multiple distinct products
  - BCM or value chain covers a product family with multiple catalogue entries
  - Task or activity description explicitly references multiple products
Return ALL applicable products in the inferred_products array.
Do NOT limit to one product if multiple are clearly supported.

CONFIDENCE LEVELS
  High   — Product directly identified from alphabet_app, bcm, or value_chain
            with no inference required. Product name or clear equivalent
            found in catalogue.
  Medium — Product inferred from task_name, l2_process context, or clear
            operational activity context. Reasonable but not explicit.
  Low    — Product inferred from general l3 description only.
            Limited direct evidence. Assign only if evidence is still clear.

NULL CASE — when to return empty products
  If no product can be confidently determined from any available field:
  - inferred_products: []
  - confidence: "Low"
  - rationale: "Insufficient evidence to identify a catalogued product
    from the available fields."
  Do NOT force an assignment. A blank is correct when evidence is absent
  or ambiguous.

OUTPUT FORMAT
Return valid JSON only. No markdown. No prose outside the JSON.
Return exactly one object per process, as a JSON array in the same
order as the input processes.

[
  {
    "l3_process_UUID": "<copied exactly from input>",
    "inferred_products": ["<exact string from catalogue>", ...],
    "confidence": "High | Medium | Low",
    "rationale": "<1-2 sentences citing the specific field(s) used>",
    "requires_human_review": true
  }
]

WORKED EXAMPLES (for calibration — do not copy, use as style guide)

Example A — High confidence via alphabet_app
  alphabet_app: "BPAY Group Platform"
  bcm: "Bill Payment Processing"
  Inferred: ["BPAY"] (if in catalogue)
  Confidence: High
  Rationale: "alphabet_app explicitly references the BPAY platform,
  which maps directly to the BPAY product in the catalogue."

Example B — Medium confidence via l2_process_name
  alphabet_app: ""
  l2_process_name: "Manage Home Loan Repayments"
  l3_activity_name: "Apply Repayment to Account"
  Inferred: ["Investor Home Loans", "Owner Occupied Home Loans"]
    (if both in catalogue and L2 context spans both)
  Confidence: Medium
  Rationale: "l2_process_name references Home Loan Repayments. Both
  Investor and Owner Occupied Home Loan products are in the catalogue
  and the L2 process spans both."

Example C — No confident match
  alphabet_app: "Generic Document Management System"
  bcm: "Document Processing"
  l3_activity_name: "Archive Completed Records"
  Inferred: []
  Confidence: Low
  Rationale: "No field provides sufficient evidence to identify a
  specific product from the catalogue. Activity appears generic
  and applicable across all products."

HARD RULES (no exceptions):
  - Every item in inferred_products MUST match exactly a string in
    the allowed_product_catalogue for this batch.
  - Return empty array rather than guessing.
  - requires_human_review must always be true.
  - Return exactly one object per input process, in the same order.
  - Do not include any text outside the JSON array.
""".strip()


# ─────────────────────────────────────────────────────────────────────────────
#  HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def is_empty(val) -> bool:
    if val is None:
        return True
    return str(val).strip().lower() in ("", "nan", "none", "null")


def parse_json_response(text: str) -> list:
    """Parse LLM response — expects a JSON array."""
    text = (text or "").strip()
    # Strip markdown code fences if present
    text = re.sub(r"^```(?:json)?\s*", "", text)
    text = re.sub(r"\s*```$", "", text)
    text = text.strip()
    try:
        result = json.loads(text)
        if isinstance(result, list):
            return result
        if isinstance(result, dict) and "mappings" in result:
            return result["mappings"]
        if isinstance(result, dict):
            return [result]
        raise ValueError(f"Unexpected JSON structure: {type(result)}")
    except json.JSONDecodeError:
        # Try to extract array
        m = re.search(r"\[.*\]", text, re.DOTALL)
        if m:
            return json.loads(m.group(0))
        raise


def validate_products(inferred: list, catalogue_set: set) -> tuple[list, list]:
    """
    Separate valid (in catalogue) from hallucinated (not in catalogue) products.
    Returns (valid_products, hallucinated_products).
    """
    valid, hallucinated = [], []
    for p in inferred:
        if str(p).strip() in catalogue_set:
            valid.append(str(p).strip())
        else:
            hallucinated.append(str(p).strip())
    return valid, hallucinated


# ─────────────────────────────────────────────────────────────────────────────
#  CHECKPOINT
# ─────────────────────────────────────────────────────────────────────────────

def load_checkpoint(path: Path) -> set:
    """Return set of already-processed l3_process_UUIDs."""
    done = set()
    if not path.exists():
        return done
    with open(path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                rec = json.loads(line)
                if "l3_process_UUID" in rec:
                    done.add(rec["l3_process_UUID"])
            except json.JSONDecodeError:
                continue
    return done


def write_checkpoint(result: dict, path: Path) -> None:
    """Append a single result record to the checkpoint JSONL."""
    try:
        with open(path, "a", encoding="utf-8") as f:
            f.write(json.dumps(result, ensure_ascii=True) + "\n")
    except (OSError, TypeError) as e:
        print(f"  WARNING: Checkpoint write failed for {result.get('l3_process_UUID')}: {e}")


# ─────────────────────────────────────────────────────────────────────────────
#  LLM CALL
# ─────────────────────────────────────────────────────────────────────────────

def call_llm(client: AzureOpenAI, batch: list, catalogue: list,
             batch_no: int) -> tuple[list, dict]:
    """
    Send a batch of processes to the LLM for product inference.
    Returns (parsed_results, usage_dict).
    """
    # Build catalogue section (once per batch)
    catalogue_text = "\n".join(f"  - {p}" for p in catalogue)

    # Build process section
    process_blocks = []
    for i, rec in enumerate(batch, 1):
        details = rec.get("process_details", {})
        block = {
            "process_number": i,
            "l3_process_UUID": rec.get("l3_process_UUID", ""),
        }
        # Include all process detail fields, skipping empty ones
        for k, v in details.items():
            if v and str(v).strip() not in ("", "None", "nan"):
                block[k] = v
        process_blocks.append(block)

    user_message = (
        PROMPT_RULES
        + f"\n\nALLOWED PRODUCT CATALOGUE ({len(catalogue)} products):\n"
        + catalogue_text
        + f"\n\nPROCESSES TO CLASSIFY (batch {batch_no}, {len(batch)} process(es)):\n"
        + json.dumps(process_blocks, indent=2, ensure_ascii=False)
        + f"\n\nReturn a JSON array of exactly {len(batch)} object(s) in order."
    )

    kwargs = {
        "model":                 MODEL,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user",   "content": user_message},
        ],
        "max_completion_tokens": MAX_COMPLETION_TOKENS,
        "response_format":       {"type": "json_object"},
    }
    if MODEL.startswith("gpt-5"):
        kwargs["reasoning_effort"] = REASONING_EFFORT

    for attempt in range(1, RETRY_COUNT + 1):
        try:
            t0       = time.time()
            response = client.chat.completions.create(**kwargs)
            latency  = int((time.time() - t0) * 1000)
            text     = response.choices[0].message.content or ""

            usage = {}
            if hasattr(response, "usage") and response.usage:
                u = response.usage
                usage = {
                    "input_tokens":  getattr(u, "prompt_tokens", 0),
                    "output_tokens": getattr(u, "completion_tokens", 0),
                }
                details = getattr(u, "completion_tokens_details", None)
                if details:
                    usage["reasoning_tokens"] = getattr(details, "reasoning_tokens", 0)
            usage["latency_ms"] = latency

            parsed = parse_json_response(text)
            return parsed, usage

        except Exception as exc:
            if attempt < RETRY_COUNT:
                sleep_t = min(RETRY_BASE * (2 ** (attempt - 1)) + random.uniform(0, 2),
                              60)
                print(f"    Attempt {attempt} failed: {exc}. Retrying in {sleep_t:.1f}s...")
                time.sleep(sleep_t)
            else:
                raise


# ─────────────────────────────────────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="LLM product inference for missing Holocentric products."
    )
    parser.add_argument("--force", action="store_true",
                        help="Ignore checkpoint and reprocess from scratch.")
    args = parser.parse_args()

    print("=" * 70)
    print("  Holocentric LLM Product Inference")
    print(f"  Model            : {MODEL}")
    print(f"  Reasoning effort : {REASONING_EFFORT}")
    print(f"  Batch size       : {BATCH_SIZE}")
    print(f"  Input JSONL      : {JSONL_FILE.name}")
    print("=" * 70)

    # ── Validate inputs ───────────────────────────────────────────────────────
    if not JSONL_FILE.exists():
        print(f"\n  ERROR: JSONL payload not found:\n  {JSONL_FILE}")
        print("  Run build_holocentric_products.py first.")
        return

    # ── Load JSONL payload ────────────────────────────────────────────────────
    print("\n  Loading JSONL payload...")
    all_records = []
    catalogue   = []

    with open(JSONL_FILE, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            rec = json.loads(line)
            all_records.append(rec)
            if not catalogue and rec.get("allowed_product_catalogue"):
                catalogue = rec["allowed_product_catalogue"]

    print(f"  Records to process   : {len(all_records):,}")
    print(f"  Product catalogue    : {len(catalogue):,} products")

    if not all_records:
        print("  No records to process. Exiting.")
        return

    catalogue_set = set(catalogue)

    # ── Checkpoint / resume ───────────────────────────────────────────────────
    if args.force and CHECKPOINT.exists():
        CHECKPOINT.unlink()
        print("\n  --force: checkpoint deleted. Starting from scratch.")

    done_uuids = load_checkpoint(CHECKPOINT)
    if done_uuids:
        print(f"  Checkpoint found: {len(done_uuids):,} already processed.")

    pending = [r for r in all_records
               if r.get("l3_process_UUID") not in done_uuids]

    if not pending:
        print("  All records already processed. Regenerating outputs only...")
    else:
        print(f"  Remaining: {len(pending):,}")

        # ── Init client ───────────────────────────────────────────────────────
        print("\n  Initialising Azure OpenAI client...")
        client = init_client()
        print("  Client ready.\n")

        # ── Process batches ───────────────────────────────────────────────────
        total_batches = (len(pending) + BATCH_SIZE - 1) // BATCH_SIZE
        total_in_tok, total_out_tok = 0, 0
        run_start = time.time()

        for batch_idx in range(0, len(pending), BATCH_SIZE):
            batch      = pending[batch_idx:batch_idx + BATCH_SIZE]
            batch_no   = batch_idx // BATCH_SIZE + 1
            pct        = batch_no / total_batches * 100
            bar        = "#" * int(pct / 5) + "." * (20 - int(pct / 5))
            elapsed    = time.time() - run_start
            eta_s      = ((elapsed / batch_no) * (total_batches - batch_no)
                          if batch_no > 1 else 0)
            eta_str    = (f"{int(eta_s//3600)}h {int((eta_s%3600)//60)}m"
                          if eta_s > 60 else f"{int(eta_s)}s")
            cost_usd   = ((total_in_tok / 1e6 * INPUT_PRICE_USD_PER_M) +
                          (total_out_tok / 1e6 * OUTPUT_PRICE_USD_PER_M))
            cost_aud   = cost_usd / AUD_USD_RATE

            print(f"  [{bar}] {pct:5.1f}%  "
                  f"Batch {batch_no}/{total_batches}  "
                  f"ETA {eta_str}  Cost A${cost_aud:.2f}")

            try:
                results, usage = call_llm(client, batch, catalogue, batch_no)
            except Exception as e:
                print(f"    ERROR in batch {batch_no}: {e}")
                print("    Skipping batch. Re-run to retry.")
                continue

            # Pair results with input UUIDs
            batch_uuids = [r.get("l3_process_UUID", "") for r in batch]

            for i, rec in enumerate(batch):
                uuid = rec.get("l3_process_UUID", "")
                # Get matching result (by position or by UUID)
                result_obj = None
                if i < len(results):
                    r = results[i]
                    # Accept if UUID matches or is absent in response
                    if isinstance(r, dict):
                        resp_uuid = r.get("l3_process_UUID", "")
                        if not resp_uuid or resp_uuid == uuid:
                            result_obj = r

                if result_obj is None:
                    print(f"    WARNING: No result for UUID {uuid} in batch {batch_no}")
                    checkpoint_rec = {
                        "l3_process_UUID":     uuid,
                        "inferred_products":   [],
                        "inferred_products_valid": [],
                        "hallucinated_products": [],
                        "confidence":          "Low",
                        "rationale":           "No result returned by model for this process.",
                        "requires_human_review": True,
                        "inference_status":    "error_no_result",
                        "batch_no":            batch_no,
                    }
                else:
                    raw_products = result_obj.get("inferred_products", [])
                    if not isinstance(raw_products, list):
                        raw_products = []
                    valid, hallucinated = validate_products(raw_products, catalogue_set)

                    if hallucinated:
                        print(f"    WARNING: Hallucinated products for {uuid}: "
                              f"{hallucinated}")

                    checkpoint_rec = {
                        "l3_process_UUID":         uuid,
                        "inferred_products":        raw_products,
                        "inferred_products_valid":  valid,
                        "hallucinated_products":    hallucinated,
                        "confidence":              result_obj.get("confidence", "Low"),
                        "rationale":               result_obj.get("rationale", ""),
                        "requires_human_review":   True,
                        "inference_status":        (
                            "hallucination_detected" if hallucinated
                            else "no_product_inferred" if not valid
                            else "success"
                        ),
                        "batch_no": batch_no,
                    }

                write_checkpoint(checkpoint_rec, CHECKPOINT)

            in_tok  = usage.get("input_tokens", 0)
            out_tok = usage.get("output_tokens", 0)
            total_in_tok  += in_tok
            total_out_tok += out_tok
            rt = (f" | {usage.get('reasoning_tokens')} thinking"
                  if usage.get("reasoning_tokens") else "")
            print(f"         {len(batch)} processes  "
                  f"{usage.get('latency_ms')}ms  "
                  f"{in_tok} in / {out_tok} out{rt}")

            if INTER_BATCH_SLEEP > 0:
                time.sleep(INTER_BATCH_SLEEP)

        final_cost = ((total_in_tok / 1e6 * INPUT_PRICE_USD_PER_M +
                       total_out_tok / 1e6 * OUTPUT_PRICE_USD_PER_M) / AUD_USD_RATE)
        print(f"\n  This run: {len(pending):,} processes  "
              f"{total_in_tok:,} in  {total_out_tok:,} out  A${final_cost:.2f}")

    # ── Build results DataFrame from checkpoint ───────────────────────────────
    print("\n  Building results from checkpoint...")
    result_rows = []
    with open(CHECKPOINT, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                rec = json.loads(line)
                result_rows.append({
                    "l3_process_UUID":         rec.get("l3_process_UUID", ""),
                    "inferred_products":       ", ".join(rec.get("inferred_products_valid", [])),
                    "inferred_products_raw":   ", ".join(rec.get("inferred_products", [])),
                    "hallucinated_products":   ", ".join(rec.get("hallucinated_products", [])),
                    "inference_confidence":    rec.get("confidence", ""),
                    "inference_rationale":     rec.get("rationale", ""),
                    "requires_human_review":   rec.get("requires_human_review", True),
                    "inference_status":        rec.get("inference_status", ""),
                    "batch_no":                rec.get("batch_no", ""),
                })
            except json.JSONDecodeError:
                continue

    results_df = pd.DataFrame(result_rows)
    print(f"  Result rows: {len(results_df):,}")

    # Summary stats
    if not results_df.empty:
        success      = (results_df["inference_status"] == "success").sum()
        no_product   = (results_df["inference_status"] == "no_product_inferred").sum()
        hallucinated = (results_df["inference_status"] == "hallucination_detected").sum()
        error        = (results_df["inference_status"] == "error_no_result").sum()
        print(f"  Success (product inferred) : {success:,}")
        print(f"  No product inferred        : {no_product:,}")
        print(f"  Hallucination detected     : {hallucinated:,}")
        print(f"  Error (no result)          : {error:,}")

    # ── Merge with cleaned dataset ────────────────────────────────────────────
    complete_df = None
    if CLEANED_FILE.exists():
        print(f"\n  Loading cleaned dataset for merge...")
        cleaned = pd.read_excel(CLEANED_FILE, engine="openpyxl", dtype=str)
        cleaned.columns = cleaned.columns.str.strip()
        print(f"  Cleaned rows: {len(cleaned):,}")

        if not results_df.empty:
            # Merge inferred products into cleaned dataset
            merged = cleaned.merge(
                results_df[["l3_process_UUID","inferred_products",
                             "inference_confidence","inference_rationale",
                             "inference_status","hallucinated_products"]],
                on="l3_process_UUID", how="left"
            )

            # Fill clean_product with inferred where currently missing
            def resolve_product(row):
                existing = str(row.get("clean_product", "")).strip()
                inferred = str(row.get("inferred_products", "")).strip()
                if existing and existing.lower() not in (
                        "", "nan", "none", "null"):
                    return existing, "original"
                if inferred:
                    return inferred, "llm_inferred"
                return "", "missing_after_inference"

            resolved = merged.apply(resolve_product, axis=1)
            merged["final_product"]        = resolved.apply(lambda x: x[0])
            merged["final_product_source"] = resolved.apply(lambda x: x[1])

            complete_df = merged
            print(f"  Complete dataset rows: {len(complete_df):,}")
    else:
        print(f"\n  WARNING: Cleaned file not found at {CLEANED_FILE}")
        print("  Skipping merge — writing results only.")

    # ── Write outputs ─────────────────────────────────────────────────────────
    print(f"\n  Writing outputs to:\n  {PRODUCTS_DIR}")

    # 1. Results
    with pd.ExcelWriter(RESULTS_FILE, engine="openpyxl") as w:
        results_df.to_excel(w, index=False, sheet_name="inference_results")

        # Summary sheet
        if not results_df.empty:
            summary = pd.DataFrame([
                ("Total processes sent to LLM",     len(results_df)),
                ("Product successfully inferred",    int((results_df["inference_status"] == "success").sum())),
                ("No product inferred (blank)",      int((results_df["inference_status"] == "no_product_inferred").sum())),
                ("Hallucination detected (removed)", int((results_df["inference_status"] == "hallucination_detected").sum())),
                ("Error — no result from model",     int((results_df["inference_status"] == "error_no_result").sum())),
            ], columns=["Metric","Value"])
            summary.to_excel(w, index=False, sheet_name="summary")

    print(f"  [1/2] {RESULTS_FILE.name}  ({len(results_df):,} rows)")

    # 2. Complete dataset
    if complete_df is not None:
        complete_df.to_excel(COMPLETE_FILE, index=False, engine="openpyxl")
        print(f"  [2/2] {COMPLETE_FILE.name}  ({len(complete_df):,} rows)")
    else:
        print(f"  [2/2] Skipped — cleaned file not available for merge")

    print("\n" + "=" * 70)
    print("  Done.")
    print("=" * 70)


def init_client() -> AzureOpenAI:
    token_provider = get_bearer_token_provider(
        DefaultAzureCredential(),
        "https://cognitiveservices.azure.com/.default"
    )
    return AzureOpenAI(
        azure_endpoint=AZURE_ENDPOINT,
        api_version=API_VERSION,
        azure_ad_token_provider=token_provider,
    )


if __name__ == "__main__":
    main()
