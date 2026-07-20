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

BATCH_SIZE    = 1     # one process per call — simplest and most reliable
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
a payment process that has no existing product assignment.

You must always produce a meaningful output. A blank result is only acceptable
when you have exhausted all evidence sources and can provide a clear reason.

Return valid JSON only. No markdown. No text outside the JSON object.
""".strip()

# ─────────────────────────────────────────────────────────────────────────────
#  PROMPT RULES (user message template)
# ─────────────────────────────────────────────────────────────────────────────

PROMPT_RULES = """
CLASSIFICATION INSTRUCTIONS
────────────────────────────────────────────────────────────────────────

CONTEXT
The process below has no existing product assignment.
You must determine the most applicable product(s) from the
allowed_product_catalogue using the evidence hierarchy below.

YOU MUST ALWAYS PRODUCE AN OUTPUT.
Do not return empty inferred_products without providing either:
  (a) a closest_alternative from the catalogue with Low confidence, or
  (b) a clear not_assigned_reason explaining why no product applies.

EVIDENCE HIERARCHY
Apply in order. Stop when sufficient evidence is found.

  Priority 1 — Alphabet Application  (field: alphabet_app)
  The application or platform name is the strongest product indicator.
  Match application names to products in the catalogue.
  Example: "BPAY Group Platform" → "BPAY" if in catalogue.
  If the app serves a product family, assign all matching catalogue entries.

  Priority 2 — Business Capability / VCM  (fields: bcm, vcm_library_name)
  BCM and VCM names often directly identify a product domain.
  Example: "Mortgage Lending" → "Investor Home Loans", "Owner Occupied Home Loans"
  Example: "Merchant Acquiring" → "In-store and Mobile Point of Sale (PoS)"

  Priority 3 — Value Chain  (field: value_chain)
  Indicates the broader product area when Priority 1-2 are not conclusive.

  Priority 4 — Task Detail  (field: task_name)
  Tasks may name products, systems or payment types directly.
  Look for product names embedded in task descriptions.

  Priority 5 — L2 Process Context  (fields: l2_process_name, l2_process_description)
  The parent process description often identifies the product family.
  Example: "Settle Loan" with Mortgage Lending context → home loan products.

  Priority 6 — L3 Activity  (fields: l3_activity_name, l3_activity_description)
  Last resort. Only use where the activity explicitly names a catalogued product.

MULTI-PRODUCT ASSIGNMENT
Assign multiple products where evidence clearly supports more than one:
  - App serves multiple distinct products
  - BCM / value chain covers a product family with multiple catalogue entries
  - Task or activity explicitly references multiple products
Return ALL applicable products. Do not limit to one if multiple are supported.

CLOSEST ALTERNATIVE RULE
If you cannot identify an exact product from the catalogue, but the evidence
points toward a product area, select the closest matching product(s) from the
catalogue and set confidence to Low.
Example: evidence suggests "Trade Finance" and "Trade Finance" is in the catalogue
— assign it even if the match is inferred rather than explicit.

NOT ASSIGNED RULE
Only leave inferred_products empty if:
  - The process is genuinely product-agnostic (e.g. a generic IT or governance
    activity with no payment product context)
  - AND no plausible closest alternative exists in the catalogue
  In this case, populate not_assigned_reason with a clear explanation.

CONFIDENCE LEVELS
  High   — Product directly and explicitly identified from alphabet_app, bcm,
            vcm_library_name or value_chain. No inference required.
  Medium — Product inferred from task_name, l2_process context, or clear
            payment processing activity context. Reasonable match.
  Low    — Product inferred from general l3 description only, or selected
            as closest alternative rather than exact match.

HARD RULES
  - Every item in inferred_products MUST match EXACTLY a string in
    the allowed_product_catalogue.
  - requires_human_review must always be true.
  - Do NOT invent, abbreviate or paraphrase product names.
  - Copy product names character-for-character from the catalogue.

OUTPUT FORMAT
Return a single JSON object. No markdown. No text outside the JSON.

{
  "l3_process_UUID": "<copied exactly from input>",
  "inferred_products": ["<exact string from catalogue>"],
  "confidence": "High | Medium | Low",
  "rationale": "<1-2 sentences citing the specific field(s) used>",
  "closest_alternative": "<single catalogue product if inferred_products is
    empty but a plausible match exists — otherwise leave blank>",
  "not_assigned_reason": "<explanation if inferred_products is empty and no
    closest alternative exists — otherwise leave blank>",
  "requires_human_review": true
}

WORKED EXAMPLES

Example A — High confidence, multi-product (vcm_library_name)
  vcm_library_name: "Mortgage Lending"
  value_chain: "Mortgage Lending"
  l2_process_name: "Settle Loan"
  inferred_products: ["Investor Home Loans", "Owner Occupied Home Loans"]
  confidence: High
  rationale: "vcm_library_name is Mortgage Lending. Both Investor and Owner
  Occupied Home Loans are in the catalogue and the Mortgage Lending value
  chain encompasses both product subtypes."

Example B — High confidence, single product (alphabet_app)
  alphabet_app: "BPAY Group Platform"
  inferred_products: ["Domestic Payments - Payer Initiated"]
  confidence: High
  rationale: "alphabet_app references the BPAY platform, which maps to
  Domestic Payments - Payer Initiated in the catalogue."

Example C — Medium confidence (l2 context)
  alphabet_app: ""
  value_chain: "Credit Cards"
  l2_process_name: "Process Credit Card Transactions"
  inferred_products: ["Business Liability Credit Cards", "Personal Liability Credit Cards"]
  confidence: Medium
  rationale: "value_chain is Credit Cards. Both Business and Personal
  Liability Credit Card products in the catalogue apply to credit card
  transaction processing."

Example D — Low confidence, closest alternative
  alphabet_app: "FX Settlement System"
  value_chain: "Markets"
  inferred_products: ["FX Derivatives"]
  confidence: Low
  closest_alternative: "FX Derivatives"
  rationale: "alphabet_app references an FX system and the catalogue
  contains FX Derivatives as the closest matching product. Assigned as
  closest alternative — exact product match not confirmed."

Example E — Not assigned (genuinely product-agnostic)
  l3_activity_name: "Archive Completed Records"
  value_chain: "Operations Support"
  bcm: "Document Management"
  inferred_products: []
  closest_alternative: ""
  not_assigned_reason: "Activity is generic document archiving with no
  payment product context. No catalogue product applies."
  confidence: Low
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

def call_llm(client: AzureOpenAI, record: dict, catalogue: list,
             batch_no: int) -> tuple[dict, dict]:
    """
    Send a SINGLE process to the LLM for product inference.
    Returns (parsed_result_dict, usage_dict).
    Batch size = 1 avoids array parsing issues.
    """
    catalogue_text = "\n".join(f"  - {p}" for p in catalogue)
    details = record.get("process_details", {})
    process_block = {"l3_process_UUID": record.get("l3_process_UUID", "")}
    for k, v in details.items():
        if v and str(v).strip() not in ("", "None", "nan"):
            process_block[k] = v

    user_message = (
        PROMPT_RULES
        + f"\n\nALLOWED PRODUCT CATALOGUE ({len(catalogue)} products):\n"
        + catalogue_text
        + "\n\nPROCESS TO CLASSIFY:\n"
        + json.dumps(process_block, indent=2, ensure_ascii=False)
        + "\n\nReturn a single JSON object for this process."
    )

    kwargs = {
        "model":   MODEL,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user",   "content": user_message},
        ],
        "max_completion_tokens": MAX_COMPLETION_TOKENS,
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
                d = getattr(u, "completion_tokens_details", None)
                if d:
                    usage["reasoning_tokens"] = getattr(d, "reasoning_tokens", 0)
            usage["latency_ms"] = latency

            parsed = parse_json_response(text)
            # parse_json_response returns list or dict; unwrap if list
            if isinstance(parsed, list) and len(parsed) == 1:
                parsed = parsed[0]
            if not isinstance(parsed, dict):
                raise ValueError(f"Expected a JSON object, got: {type(parsed)}")
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

        # One call per process (BATCH_SIZE=1 eliminates array parsing issues)
        for rec_idx, rec in enumerate(pending):
            process_no = rec_idx + 1
            uuid       = rec.get("l3_process_UUID", "")
            pct        = process_no / total_batches * 100
            bar        = "#" * int(pct / 5) + "." * (20 - int(pct / 5))
            elapsed    = time.time() - run_start
            eta_s      = ((elapsed / process_no) * (total_batches - process_no)
                          if process_no > 1 else 0)
            eta_str    = (f"{int(eta_s//3600)}h {int((eta_s%3600)//60)}m"
                          if eta_s > 60 else f"{int(eta_s)}s")
            cost_aud   = ((total_in_tok / 1e6 * INPUT_PRICE_USD_PER_M +
                           total_out_tok / 1e6 * OUTPUT_PRICE_USD_PER_M)
                          / AUD_USD_RATE)

            print(f"  [{bar}] {pct:5.1f}%  "
                  f"{process_no}/{total_batches}  "
                  f"ETA {eta_str}  Cost A${cost_aud:.2f}")

            try:
                result_obj, usage = call_llm(client, rec, catalogue, process_no)
            except Exception as e:
                print(f"    ERROR for {uuid}: {e}")
                write_checkpoint({
                    "l3_process_UUID":          uuid,
                    "inferred_products":         [],
                    "inferred_products_valid":   [],
                    "hallucinated_products":     [],
                    "confidence":               "Low",
                    "rationale":               f"Model error: {e}",
                    "closest_alternative":      "",
                    "not_assigned_reason":      f"Script error — re-run to retry: {e}",
                    "requires_human_review":    True,
                    "inference_status":         "error",
                    "batch_no":                 process_no,
                }, CHECKPOINT)
                continue

            raw_products = result_obj.get("inferred_products", [])
            if not isinstance(raw_products, list):
                raw_products = []
            valid, hallucinated = validate_products(raw_products, catalogue_set)

            if hallucinated:
                print(f"    WARNING: Hallucinated products removed for {uuid}: "
                      f"{hallucinated}")

            # If inferred_products had hallucinations but we also have a
            # closest_alternative, try to salvage it
            closest_alt = str(result_obj.get("closest_alternative", "") or "").strip()
            if closest_alt and closest_alt not in catalogue_set:
                print(f"    WARNING: closest_alternative not in catalogue "
                      f"for {uuid}: '{closest_alt}' — cleared")
                closest_alt = ""

            status = (
                "hallucination_detected" if hallucinated and not valid
                else "success_with_hallucination_removed" if hallucinated
                else "no_product_inferred" if not valid and not closest_alt
                else "closest_alternative_used" if not valid and closest_alt
                else "success"
            )

            checkpoint_rec = {
                "l3_process_UUID":          uuid,
                "inferred_products":         raw_products,
                "inferred_products_valid":   valid,
                "hallucinated_products":     hallucinated,
                "confidence":               result_obj.get("confidence", "Low"),
                "rationale":               result_obj.get("rationale", ""),
                "closest_alternative":      closest_alt,
                "not_assigned_reason":      str(
                    result_obj.get("not_assigned_reason", "") or "").strip(),
                "requires_human_review":    True,
                "inference_status":         status,
                "batch_no":                 process_no,
            }
            write_checkpoint(checkpoint_rec, CHECKPOINT)

            in_tok  = usage.get("input_tokens", 0)
            out_tok = usage.get("output_tokens", 0)
            total_in_tok  += in_tok
            total_out_tok += out_tok
            rt = (f" | {usage.get('reasoning_tokens')} thinking"
                  if usage.get("reasoning_tokens") else "")
            prod_summary = (", ".join(valid) if valid
                            else f"[alt: {closest_alt}]" if closest_alt
                            else "[none]")
            print(f"         {uuid[:20]}...  "
                  f"{usage.get('latency_ms')}ms  "
                  f"{in_tok}in/{out_tok}out{rt}  → {prod_summary}")

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
                valid    = rec.get("inferred_products_valid", [])
                closest  = str(rec.get("closest_alternative", "") or "").strip()
                # final_inferred: validated products, or closest alternative if empty
                final    = valid if valid else ([closest] if closest else [])
                result_rows.append({
                    "l3_process_UUID":         rec.get("l3_process_UUID", ""),
                    "inferred_products":       ", ".join(final),
                    "inferred_products_raw":   ", ".join(rec.get("inferred_products", [])),
                    "hallucinated_products":   ", ".join(rec.get("hallucinated_products", [])),
                    "closest_alternative":     closest,
                    "not_assigned_reason":     str(rec.get("not_assigned_reason", "") or "").strip(),
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
                ("Total processes sent to LLM",
                 len(results_df)),
                ("Product assigned (exact match)",
                 int((results_df["inference_status"] == "success").sum())),
                ("Product assigned (hallucination removed, still valid)",
                 int((results_df["inference_status"] == "success_with_hallucination_removed").sum())),
                ("Closest alternative used (exact not found)",
                 int((results_df["inference_status"] == "closest_alternative_used").sum())),
                ("No product inferred (blank — see not_assigned_reason)",
                 int((results_df["inference_status"] == "no_product_inferred").sum())),
                ("Hallucination detected, no valid product",
                 int((results_df["inference_status"] == "hallucination_detected").sum())),
                ("Error — model call failed",
                 int((results_df["inference_status"].isin(["error","error_no_result"])).sum())),
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
