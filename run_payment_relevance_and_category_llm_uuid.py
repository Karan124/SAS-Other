r"""
run_payment_relevance_and_category_llm_uuid.py
---------------------------------------------
Runs two-step LLM classification over UUID-based Holocentric process payloads:
1. Determine whether the process is a payments process / payments-enabling process.
2. If payment-related, map to valid payment category/categories.

Default input:
Z:\Enterprise Risk Insights\23 _sas_batch_\Controls-PoC\processes\holo_process_samples_aggregated_for_llm_uuid.json

Before running:
az login
az account set --subscription 6c72e6c5-ed48-4030-b29c-34e2849c9288
cd C:\Users\m061400\ai-test
$env:REQUESTS_CA_BUNDLE="C:\Users\m061400\ai-test\cacert.pem"
$env:SSL_CERT_FILE="C:\Users\m061400\ai-test\cacert.pem"
Remove-Item Env:AZURE_CA_BUNDLE -ErrorAction SilentlyContinue
"""

import argparse
import json
import os
import re
import subprocess
import time
from datetime import datetime
from pathlib import Path

import pandas as pd
from azure.identity import AzureCliCredential, get_bearer_token_provider
from openai import AzureOpenAI

ENDPOINT = "https://ai.eng.azure.srv.westpac.com.au"
API_VERSION = "2024-10-21"
MODEL = "gpt-5.4"
DEFAULT_INPUT_FILE = r"Z:\Enterprise Risk Insights\23 _sas_batch_\Controls-PoC\processes\holo_process_samples_aggregated_for_llm_uuid.json"
BATCH_SIZE = 5
MAX_COMPLETION_TOKENS = 12000
REASONING_EFFORT = "high"
RETRY_COUNT = 2
RETRY_SLEEP_SECONDS = 3

VALID_CATEGORIES = [
    "Customer to Customer", "Customer to Institution", "Institution to Customer",
    "Institution to Institution", "Supplier / Contractor / Employee Payments",
]

SYSTEM_PROMPT = """You are a payments process assessment and category mapping analyst for an Australian ADI.
Return valid JSON only. Do not include markdown or prose outside JSON.
Do not invent facts. Be conservative where evidence is sparse.
Every rationale must include at least one short quoted phrase from the supplied process payload."""

PROMPT_RULES = """
Role
For each process in the batch, perform two steps.

Step 1 — Payment relevance assessment:
Determine whether the process is actually a payments process or payments-enabling process based only on the supplied evidence.

A process is payment-related if it directly executes, supports, controls, monitors, reconciles, settles, clears, posts, amends, cancels, routes, releases, validates, or governs payment movement, payment instructions, payment systems, payment exceptions, payment settlement, or payment operations.

Step 2 — Payment category mapping:
If the process is payment-related, determine which payment category or categories it maps to using only the five valid payment categories and the mapping rules.

If the process is not payment-related, do not map it to a payment category.

Identifier Rule
Each process includes l3_process_UUID. This is the unique process identifier and must be copied exactly into the output.
Do not use l3_activity_id or process_id as the unique join key because the same activity ID may appear in multiple process contexts.

Valid Payment Categories
Use only:
- Customer to Customer
- Customer to Institution
- Institution to Customer
- Institution to Institution
- Supplier / Contractor / Employee Payments

Category Definitions
- Customer to Customer — individual / merchant / business pays individual / merchant / business.
- Customer to Institution — customer pays the bank or deposits/repays funds into the institution.
- Institution to Customer — the bank or institution disburses or pays funds to the customer.
- Institution to Institution — one bank / financial institution pays, settles with, clears with, or exchanges obligations with another institution.
- Supplier / Contractor / Employee Payments — the bank pays suppliers, contractors, or employees, including internal payroll and vendor payments.

Minimum Mapping Requirement
Category mapping requires at least one of:
(a) explicit payer-payee relationship
(b) explicit payment action, such as settle, disburse, process payment
(c) explicit payment system/process reference, such as RTGS, BPAY, clearing, RITS, NPP, Direct Entry, PEXA, SWIFT
(d) explicit system, third-party, or governance reference in the context of payment processing.

If the process is payment-related but no payer-payee category is defensible, set is_payment_process=true, mapped_categories=[], primary_category=null, and explain why the category is not defensible.

Mapping Rules
R1 — Map only where a payer-payee relationship is clear or strongly implied by a named payment system/process.
R2 — Process must operate in a category-specific payment scenario, not generic enterprise activity.
R3 — Process must create, control, execute, clear, settle, post, amend, cancel, route, release, correct, reconcile or materially affect payment movement.
R4 — Multiple categories are allowed only where explicitly supported across different payer-payee relationships. Always choose a primary_category if any category is mapped.
R5 — For systems, schemes and third parties, determine category by ultimate payer-payee relationship, not intermediary processors.
R6 — L3 activity name is a primary signal, but naming alone must not drive mapping.

Priority Guidance
- Interbank settlement, clearing networks, counterparties, schemes, external banks, RBA, RITS, OFIs, settlement instructions or obligations between payment providers -> Institution to Institution.
- Customer repayments, customer deposits, ATM deposits, BPAY, card repayments, customers sending funds into the bank -> Customer to Institution.
- Disbursements, refunds, interest payments, loan settlements, welfare, benefit or tax payments, or payroll processed by the bank on behalf of a client institution -> Institution to Customer.
- Transfers between customers, merchants or businesses, including card payments, P2P, internal transfers, direct debit or customer-originated movement between non-bank parties -> Customer to Customer.
- Vendor payments, internal staff payroll, supplier/contractor/employee payments internal to the bank -> Supplier / Contractor / Employee Payments.

Confidence Guidance
Use only High, Medium or Low.
- payment_process_confidence measures confidence that the process is payment-related.
- category_confidence_score measures confidence in the payer-payee category mapping.
A process can be clearly payment-related but still have Low category confidence if payer/payee evidence is missing.

Output exactly:
{
  "mappings": [
    {
      "l3_process_UUID": "string copied exactly from input processes[].l3_process_UUID",
      "process_id": "string copied from input processes[].process_id",
      "l2_process_name": "string or null",
      "l3_activity_name": "string or null",
      "is_payment_process": true,
      "payment_process_confidence": "High | Medium | Low",
      "payment_process_rationale": "concise explanation grounded in quoted process text",
      "mapped_categories": ["valid categories only, or empty array"],
      "primary_category": "valid category or null",
      "category_confidence_score": "High | Medium | Low",
      "rule_hits": ["R1", "R2", "R3", "R4", "R5", "R6"],
      "category_mapping_rationale": "concise explanation grounded in quoted process text; explain why chosen categories apply and close alternatives are less appropriate"
    }
  ],
  "notes": ["optional batch-level caveats"]
}

Rules
- Emit exactly one mapping entry for every input process.
- Use only l3_process_UUID values present in the input batch.
- Every payment_process_rationale and category_mapping_rationale must include quoted source text.
- If is_payment_process=false, mapped_categories must be [], primary_category must be null, and category_confidence_score should be Low.
"""


def run_shell_command(command: str):
    return subprocess.run(command, capture_output=True, text=True, shell=True, check=True)


def preflight_environment_check():
    print("\nPre-flight environment/auth check")
    print("-" * 72)
    print(f"REQUESTS_CA_BUNDLE = {os.getenv('REQUESTS_CA_BUNDLE')}")
    print(f"SSL_CERT_FILE      = {os.getenv('SSL_CERT_FILE')}")
    print(f"HTTP_PROXY         = {os.getenv('HTTP_PROXY')}")
    print(f"HTTPS_PROXY        = {os.getenv('HTTPS_PROXY')}")
    print(f"AZURE_CA_BUNDLE    = {os.getenv('AZURE_CA_BUNDLE')}")
    if os.getenv("AZURE_CA_BUNDLE"):
        print("WARNING: AZURE_CA_BUNDLE is set. Recommended to remove it for Azure CLI auth.")
    run_shell_command("az account show")
    print("az account show    = OK")


def init_client():
    credential = AzureCliCredential()
    token_provider = get_bearer_token_provider(credential, "https://cognitiveservices.azure.com/.default")
    return AzureOpenAI(azure_endpoint=ENDPOINT, api_version=API_VERSION, azure_ad_token_provider=token_provider)


def load_processes(input_file: str):
    with open(input_file, "r", encoding="utf-8") as f:
        data = json.load(f)
    if isinstance(data, dict) and "processes" in data:
        return data["processes"]
    if isinstance(data, list):
        return data
    raise ValueError("Input must be a JSON list or object with 'processes'.")


def chunk_list(items, size):
    for i in range(0, len(items), size):
        yield i // size + 1, items[i:i + size]


def parse_json_response(text: str):
    text = (text or "").strip()
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        m = re.search(r"\{.*\}", text, flags=re.DOTALL)
        if m:
            return json.loads(m.group(0))
        raise


def validate_output(parsed, batch):
    warnings = []
    input_ids = {str(p.get("l3_process_UUID")) for p in batch}
    mappings = parsed.get("mappings", [])
    output_ids = {str(m.get("l3_process_UUID")) for m in mappings}
    if input_ids - output_ids:
        warnings.append(f"Missing UUID mappings: {sorted(input_ids - output_ids)}")
    if output_ids - input_ids:
        warnings.append(f"Unknown UUIDs returned: {sorted(output_ids - input_ids)}")
    if len(mappings) != len(batch):
        warnings.append(f"Expected {len(batch)} mappings but received {len(mappings)}")
    for m in mappings:
        for c in (m.get("mapped_categories") or []):
            if c not in VALID_CATEGORIES:
                warnings.append(f"Invalid category {c!r} for UUID {m.get('l3_process_UUID')}")
        primary = m.get("primary_category")
        if primary is not None and primary not in VALID_CATEGORIES:
            warnings.append(f"Invalid primary_category {primary!r} for UUID {m.get('l3_process_UUID')}")
        if m.get("is_payment_process") is False and (m.get("mapped_categories") or []):
            warnings.append(f"is_payment_process=false but categories returned for UUID {m.get('l3_process_UUID')}")
    return warnings


def call_llm(client, batch, batch_no):
    prompt = PROMPT_RULES + "\n\nProcesses to assess\n" + json.dumps(batch, ensure_ascii=False, indent=2)
    kwargs = {
        "model": MODEL,
        "messages": [{"role": "system", "content": SYSTEM_PROMPT}, {"role": "user", "content": prompt}],
        "max_completion_tokens": MAX_COMPLETION_TOKENS,
        "response_format": {"type": "json_object"},
    }
    if MODEL.startswith("gpt-5") and REASONING_EFFORT:
        kwargs["reasoning_effort"] = REASONING_EFFORT
    for attempt in range(1, RETRY_COUNT + 1):
        try:
            t0 = time.time()
            response = client.chat.completions.create(**kwargs)
            latency_ms = int((time.time() - t0) * 1000)
            text = response.choices[0].message.content or ""
            parsed = parse_json_response(text)
            usage_obj = getattr(response, "usage", None)
            usage = {}
            if usage_obj:
                usage = {
                    "input_tokens": getattr(usage_obj, "prompt_tokens", None),
                    "output_tokens": getattr(usage_obj, "completion_tokens", None),
                    "total_tokens": getattr(usage_obj, "total_tokens", None),
                }
                details = getattr(usage_obj, "completion_tokens_details", None)
                if details:
                    usage["reasoning_tokens"] = getattr(details, "reasoning_tokens", None)
            return parsed, {"batch_no": batch_no, "latency_ms": latency_ms, "usage": usage, "raw_response": text}
        except Exception as exc:
            if attempt < RETRY_COUNT:
                print(f"    Attempt {attempt} failed: {exc}. Retrying in {RETRY_SLEEP_SECONDS}s...")
                time.sleep(RETRY_SLEEP_SECONDS)
            else:
                raise


def flatten_results(batch_results):
    rows = []
    for result in batch_results:
        usage = result.get("usage", {}) or {}
        for m in result.get("mappings", []):
            rows.append({
                "batch_no": result.get("batch_no"),
                "l3_process_UUID": m.get("l3_process_UUID"),
                "process_id": m.get("process_id"),
                "l2_process_name": m.get("l2_process_name"),
                "l3_activity_name": m.get("l3_activity_name"),
                "is_payment_process": m.get("is_payment_process"),
                "payment_process_confidence": m.get("payment_process_confidence"),
                "payment_process_rationale": m.get("payment_process_rationale"),
                "mapped_categories": " | ".join(m.get("mapped_categories", []) or []),
                "primary_category": m.get("primary_category"),
                "category_confidence_score": m.get("category_confidence_score"),
                "rule_hits": " | ".join(m.get("rule_hits", []) or []),
                "category_mapping_rationale": m.get("category_mapping_rationale"),
                "input_tokens": usage.get("input_tokens"),
                "output_tokens": usage.get("output_tokens"),
                "total_tokens": usage.get("total_tokens"),
                "reasoning_tokens": usage.get("reasoning_tokens"),
            })
    return pd.DataFrame(rows)


def write_outputs(input_file, batch_results, raw_records, warnings):
    input_path = Path(input_file)
    output_dir = input_path.parent
    stem = input_path.stem.replace("_aggregated_for_llm_uuid", "")
    json_out = output_dir / f"{stem}_llm_payment_relevance_category_uuid_high.json"
    jsonl_out = output_dir / f"{stem}_llm_payment_relevance_category_uuid_high_raw.jsonl"
    xlsx_out = output_dir / f"{stem}_llm_payment_relevance_category_uuid_high.xlsx"
    combined = {
        "run_ts": datetime.now().isoformat(timespec="seconds"),
        "config": {"endpoint": ENDPOINT, "api_version": API_VERSION, "model": MODEL, "batch_size": BATCH_SIZE, "reasoning_effort": REASONING_EFFORT, "max_completion_tokens": MAX_COMPLETION_TOKENS},
        "warnings": warnings,
        "mappings": [],
        "batch_metadata": [],
    }
    for result in batch_results:
        combined["mappings"].extend(result.get("mappings", []))
        combined["batch_metadata"].append({"batch_no": result.get("batch_no"), "latency_ms": result.get("latency_ms"), "usage": result.get("usage"), "warnings": result.get("warnings", []), "notes": result.get("notes", [])})
    with open(json_out, "w", encoding="utf-8") as f:
        json.dump(combined, f, ensure_ascii=False, indent=2)
    with open(jsonl_out, "w", encoding="utf-8") as f:
        for record in raw_records:
            f.write(json.dumps(record, ensure_ascii=False) + "\n")
    with pd.ExcelWriter(xlsx_out, engine="openpyxl") as writer:
        flatten_results(batch_results).to_excel(writer, index=False, sheet_name="mappings")
        pd.DataFrame(combined["batch_metadata"]).to_excel(writer, index=False, sheet_name="batch_metadata")
        pd.DataFrame({"warning": warnings}).to_excel(writer, index=False, sheet_name="warnings")
    print("\nOutputs saved:")
    print(f"  JSON : {json_out}")
    print(f"  JSONL: {jsonl_out}")
    print(f"  Excel: {xlsx_out}")


def main():
    parser = argparse.ArgumentParser(description="Run payment relevance and category mapping using l3_process_UUID.")
    parser.add_argument("input_file", nargs="?", default=DEFAULT_INPUT_FILE)
    parser.add_argument("--batch-size", type=int, default=BATCH_SIZE)
    args = parser.parse_args()
    if not os.path.exists(args.input_file):
        raise FileNotFoundError(f"Input file not found: {args.input_file}")
    print("=" * 72)
    print("Payment relevance + category LLM runner using l3_process_UUID")
    print(f"Model                 : {MODEL}")
    print(f"Reasoning effort      : {REASONING_EFFORT}")
    print(f"Max completion tokens : {MAX_COMPLETION_TOKENS}")
    print(f"Input file            : {args.input_file}")
    print(f"Batch size            : {args.batch_size}")
    print("=" * 72)
    processes = load_processes(args.input_file)
    missing_uuid = [p.get("process_id") for p in processes if not p.get("l3_process_UUID")]
    if missing_uuid:
        raise ValueError("Some input processes are missing l3_process_UUID. Re-run the updated aggregation script first.")
    print(f"Loaded {len(processes):,} process payloads")
    preflight_environment_check()
    print("\nInitialising Azure OpenAI client...")
    client = init_client()
    print("Client ready")
    batch_results, raw_records, all_warnings = [], [], []
    total_batches = (len(processes) + args.batch_size - 1) // args.batch_size
    for batch_no, batch in chunk_list(processes, args.batch_size):
        print(f"\nBatch {batch_no}/{total_batches}: {len(batch)} process(es)")
        parsed, meta = call_llm(client, batch, batch_no)
        warnings = validate_output(parsed, batch)
        result = {"batch_no": batch_no, "mappings": parsed.get("mappings", []), "notes": parsed.get("notes", []), "warnings": warnings, "latency_ms": meta.get("latency_ms"), "usage": meta.get("usage")}
        batch_results.append(result)
        raw_records.append({"batch_no": batch_no, "input_l3_process_UUIDs": [p.get("l3_process_UUID") for p in batch], "parsed_response": parsed, "raw_response": meta.get("raw_response"), "usage": meta.get("usage"), "latency_ms": meta.get("latency_ms"), "warnings": warnings})
        all_warnings.extend([f"Batch {batch_no}: {w}" for w in warnings])
        print(f"  Returned mappings: {len(result['mappings'])}")
        print(f"  Latency: {result['latency_ms']}ms")
        print(f"  Usage: {result['usage']}")
        if warnings:
            print("  Warnings:")
            for warning in warnings:
                print(f"    - {warning}")
    write_outputs(args.input_file, batch_results, raw_records, all_warnings)
    print("\nDone.")

if __name__ == "__main__":
    main()
