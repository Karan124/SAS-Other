r"""
run_lifecycle_mapping_llm.py
───────────────────────────────────────────────────────────────────────────────
Payments Controls PoC — Payment Process Lifecycle Stage Mapper

Input:
  Two files joined by l3_process_UUID:
  1. PROCESS_INPUT_FILE  — original process candidates JSON (all source fields)
  2. CATEGORY_OUTPUT_FILE — output from the category mapping run (payment_process_type,
                            primary_category, etc.)
  Only processes where payment_process_type = "Direct" are sent to the LLM.

Output:
  JSONL checkpoint  — incremental, written per batch (enables resume)
  JSON summary      — full structured results
  Excel             — mappings, lifecycle_eligible, not_eligible, sme_review,
                      summary, batch_metadata sheets

Before running (PowerShell):
  az account set --subscription 6c72e6c5-ed48-4030-b29c-34e2849c9288
  $env:REQUESTS_CA_BUNDLE = "C:\path\to\cacert.pem"
  $env:SSL_CERT_FILE      = "C:\path\to\cacert.pem"
  Remove-Item Env:AZURE_CA_BUNDLE -ErrorAction SilentlyContinue
  python run_lifecycle_mapping_llm.py
  python run_lifecycle_mapping_llm.py --force   # ignore checkpoint, restart
"""

import argparse
import json
import os
import random
import re
import time
from collections import Counter
from datetime import datetime
from pathlib import Path

import pandas as pd
from azure.identity import DefaultAzureCredential, get_bearer_token_provider
from openai import AzureOpenAI

# ──────────────────────────────────────────────────────────────────────────────
#  CONFIG — update file paths before running
# ──────────────────────────────────────────────────────────────────────────────

ENDPOINT    = "https://ai.eng.azure.srv.westpac.com.au"
API_VERSION = "2024-10-21"
MODEL       = "gpt-5.4"

# Input files — update these paths
PROCESS_INPUT_FILE   = r"Z:\path\to\process_candidates_aggregated.json"
CATEGORY_OUTPUT_FILE = r"Z:\path\to\payment_category_output.json"
OUTPUT_DIR           = r"Z:\path\to\output"

BATCH_SIZE             = 3
MAX_COMPLETION_TOKENS  = 16000
REASONING_EFFORT       = "medium"

RETRY_COUNT       = 3
RETRY_BASE_SLEEP  = 5
RETRY_MAX_SLEEP   = 60
INTER_BATCH_SLEEP = 0.5

INPUT_PRICE_USD_PER_M  = 1.75
OUTPUT_PRICE_USD_PER_M = 14.00
AUD_USD_RATE           = 0.65

VALID_LIFECYCLE_STAGES = [
    "Initiation & Validation & Authorisation",
    "Execution & Early Processing Assurance",
    "Clearing / Settlement",
    "Posting & Accounting, Detection",
    "Notification & Reporting",
    "Incident response, disputes, recovery followups",
]

# Source fields to carry into the lifecycle payload from the original input
SOURCE_FIELDS = [
    "l2_process_name", "l2_process_description",
    "l3_activity_name", "l3_activity_description",
    "l3_activity_channels", "l3_activity_customer_segments",
    "l3_activity_product_service", "tasks", "systems",
    "third_parties", "value_stream_name", "vcm_library_name",
]

# ──────────────────────────────────────────────────────────────────────────────
#  SYSTEM PROMPT
# ──────────────────────────────────────────────────────────────────────────────

SYSTEM_PROMPT = (
    "You are a payments process-to-lifecycle mapping analyst for an Australian ADI. "
    "All input processes have already been classified as Direct payment processes. "
    "Your only task is to determine lifecycle eligibility and assign the correct "
    "lifecycle stage from the six valid stages. "
    "Return valid JSON only. No markdown. No prose outside the JSON object. "
    "Do not invent facts. All reasoning must be grounded in supplied fields only. "
    "Every lifecycle_rationale must include at least one quoted phrase from the process."
)

# ──────────────────────────────────────────────────────────────────────────────
#  PROMPT
# ──────────────────────────────────────────────────────────────────────────────

PROMPT_RULES = """
Role

All input processes are pre-classified Direct payment processes from the
payment category mapping run. You do not need to re-assess payment relevance.

Your task for each process:
  1. Determine lifecycle eligibility — does the activity actually perform
     a payment lifecycle outcome, or is it product/account/facility/contract
     lifecycle management that happens to be tagged as Direct?
  2. If eligible, assign the closest defensible primary lifecycle stage.
  3. Provide a clear rationale grounded in quoted text from the process.

CRITICAL ROUTING RULE:

  Lifecycle eligible   → assign primary_lifecycle_stage from the six valid stages
  Not lifecycle eligible → primary_lifecycle_stage = null, explain why clearly

──────────────────────────────────────────────────────────────────────────────
STEP 1 — LIFECYCLE ELIGIBILITY CHECK
──────────────────────────────────────────────────────────────────────────────

Even though all inputs are Direct, some activities manage product, account,
facility, contract or customer lifecycle rather than the payment lifecycle.
Setting up capability for future payments is NOT the same as processing a
payment. Classify is_lifecycle_eligible=false where the activity primarily:

NOT LIFECYCLE ELIGIBLE (common cases):
  - Prepares contract documentation, terms and conditions, or facility documents
  - Sets up or activates accounts, cards, products or facilities
  - Onboards customers, brokers or third parties
  - Manages product or account lifecycle (maintenance, modification, closure)
  - Issues cards, PINs, plastics or welcome communications
  - Creates or configures records without processing a payment event
  - Manages procurement, purchasing or invoice administration without
    processing an actual payment transaction or payment file
  - Manages physical facility access or security intelligence without direct
    operational linkage to payment execution, settlement or incident response
  - Performs credit assessment, credit approval or facility approval
    where no specific payment instruction is directly authorised for release

LIFECYCLE ELIGIBLE: activities that directly perform one of these outcomes:
  - Receive or capture a payment instruction
  - Validate payment, account, beneficiary or biller details
  - Check funds availability, limits, sanctions or fraud before payment release
  - Authorise or approve a payment for release
  - Execute, release, route, transmit or batch a payment into processing
  - Clear or settle a payment obligation between parties
  - Post, apply, account for or reconcile a payment outcome
  - Notify a customer or counterparty of a payment outcome or status
  - Investigate, reverse, recall, recover, reissue or remediate a payment failure

IMPORTANT DISTINCTION:
  Direct Debit instruction SETUP that captures a future payment authority
  → Lifecycle eligible (Initiation & Validation & Authorisation)
  Loan facility SETUP before any drawdown/disbursement occurs
  → NOT lifecycle eligible
  Loan DRAWDOWN processing or disbursement payment execution
  → Lifecycle eligible

──────────────────────────────────────────────────────────────────────────────
STEP 2 — LIFECYCLE STAGE ASSIGNMENT (lifecycle-eligible processes only)
──────────────────────────────────────────────────────────────────────────────

Identify the primary operational outcome of the L3 activity, then assign the
closest defensible stage. Where multiple stages appear, assign the dominant
operational outcome as primary and list others as secondary.

VALID LIFECYCLE STAGES (use only these exact strings):
  - Initiation & Validation & Authorisation
  - Execution & Early Processing Assurance
  - Clearing / Settlement
  - Posting & Accounting, Detection
  - Notification & Reporting
  - Incident response, disputes, recovery followups

Stage Definitions:

  Initiation & Validation & Authorisation
  The payment instruction is captured, received, validated and approved before
  execution begins. Covers: payment request capture, beneficiary/account/biller
  validation, available funds check, limit check, fraud/AML/sanctions check
  before release, maker-checker approval, customer authentication, direct debit
  instruction setup, drawdown authorisation and payment approval workflows.
  Exclude: product/account/customer onboarding, contract/document validation,
  credit approval unless the activity directly validates or authorises payment release.

  Execution & Early Processing Assurance
  An authorised payment is released into processing systems, gateways, queues,
  files or rails before settlement completes. Covers: payment release, routing,
  batching, formatting, file generation, transmission, gateway exchange, early
  processing, early rejects, payment file failures and routing errors before
  settlement.

  Clearing / Settlement
  Payment obligations are cleared, exchanged, settled or discharged between
  parties. Covers: interbank settlement, scheme settlement, RTGS/RITS settlement,
  correspondent banking settlement, internal settlement, ESA/RBA settlement,
  completion of settlement obligations.
  Exclude: storing settlement instructions or setting up settlement accounts.

  Posting & Accounting, Detection
  Payment outcomes are recorded, applied to accounts, reconciled or detected as
  exceptions after processing or settlement. Covers: debit/credit posting, balance
  updates, suspense/nostro/GL entries, applying repayments to loan/card/deposit
  accounts, post-processing reconciliation, detection of posting or accounting breaks.

  Notification & Reporting
  Completed payment outcomes or statuses are communicated to customers,
  counterparties or stakeholders. Covers: payment confirmation, payment advice,
  remittance advice, drawdown/funding confirmation, settlement confirmation,
  payslip/payment advice, payment status notification.
  Exclude: product setup communications, welcome letters, contract dispatch,
  generic management reporting, reconciliation reporting to leadership.

  Incident response, disputes, recovery followups
  Payment failures, disputes, exceptions or recoveries are investigated, corrected,
  reversed, recalled, returned, reprocessed, remediated or closed. Covers: failed
  payment investigation, duplicate/incorrect/delayed payment resolution, rejected
  payment repair, unauthorised/disputed transaction handling, recall, return,
  reversal, reissue, refund, recovery, re-run, suspense resolution, mistaken
  payment recovery.

Boundary Guidance:
  Before release of authorised payment    → Initiation & Validation & Authorisation
  Released/routed before settlement       → Execution & Early Processing Assurance
  Settlement obligation completed         → Clearing / Settlement
  Posted/reconciled after settlement      → Posting & Accounting, Detection
  Payment outcome communicated            → Notification & Reporting
  Payment issue investigated/recovered    → Incident response, disputes, recovery followups

Where the activity spans multiple stages, assign the dominant operational outcome
as primary_lifecycle_stage. Include others in secondary_lifecycle_stages.

──────────────────────────────────────────────────────────────────────────────
SUPPORTING CONTEXT RULES
──────────────────────────────────────────────────────────────────────────────

USING CATEGORY CONTEXT
The primary_category field from the category mapping run indicates the payer-payee
relationship. Use it to clarify ambiguous lifecycle activity interpretations:
  - Customer to Institution + reconciliation activity → likely Posting & Accounting
  - Institution to Customer + notification activity → Notification & Reporting
  - Institution to Institution + settlement → Clearing / Settlement
  - Supplier/Contractor/Employee + disbursement → Execution or Posting

USING CHANNEL, SEGMENT AND PRODUCT/SERVICE CONTEXT
These fields support but do not determine lifecycle stage:
  - l3_activity_channels: indicates where the process occurs (ATM, Branch, Digital,
    GTS Direct Connectivity, Corporate Online)
  - l3_activity_customer_segments: indicates parties (Consumer, Commercial, Institutional)
  - l3_activity_product_service: indicates product context (Home Loans, BPAY,
    Transaction Accounts, Trade Finance, FX Derivatives)

Product context guidance:
  - BPAY: repayment/payment capture → Initiation; processing → Execution
  - Home Loans / Mortgage: drawdown authorisation → Initiation; disbursement → Execution;
    repayment posting → Posting & Accounting; confirmation → Notification
  - FX / SWIFT: authorisation → Initiation; transmission → Execution;
    settlement → Clearing / Settlement; confirmation → Notification
  - ATM: withdrawal/deposit processing → Execution; reconciliation → Posting

PAYMENT RAIL NEUTRALITY
Do not determine lifecycle stage solely from a payment rail or mechanism.
SWIFT, RTGS, RITS, NPP, BPAY, Direct Entry, EFT, Visa, Mastercard, OFI, ESA
are rails and settlement mechanisms — not lifecycle stages.
The activity performed determines the stage:
  - "Authorise Outbound Payments - SWIFT" → Initiation & Validation & Authorisation
    (if validating/approving before release)
  - "Execute Outbound Payment - SWIFT" → Execution & Early Processing Assurance
    (if releasing/transmitting payment)
  - "Settle SWIFT Obligations" → Clearing / Settlement

INBOUND / OUTBOUND RULE
Do not derive lifecycle stage from inbound/outbound alone.
Use the actual activity description, operational outcome and product context.

LENDING, FUNDING AND DRAWDOWN
  - Mortgage drawdown authorisation → Initiation & Validation & Authorisation
  - Loan disbursement/settlement → Execution & Early Processing Assurance
  - Loan repayment posting → Posting & Accounting, Detection
  - Disbursement/repayment confirmation → Notification & Reporting
  - Failed disbursement investigation → Incident response, disputes, recovery followups
  - Credit assessment, facility approval with no direct payment release → NOT eligible

──────────────────────────────────────────────────────────────────────────────
WORKED EXAMPLES
──────────────────────────────────────────────────────────────────────────────

Example 1 — Initiation & Validation & Authorisation
  L3 Activity: Validate BPAY Payment Instruction
  Description: "Customer submits BPAY payment; system validates biller code,
  customer reference number and available funds before authorisation"
  primary_category: Customer to Institution
  is_lifecycle_eligible: true
  primary_lifecycle_stage: Initiation & Validation & Authorisation
  Why: Activity "validates biller code... and available funds before authorisation"
  — this is pre-execution validation of a payment instruction. Execution is less
  appropriate because no payment has been released yet. Posting is less appropriate
  because no settlement has occurred.

Example 2 — Execution & Early Processing Assurance
  L3 Activity: Release Approved Payments to NPP Gateway
  Description: "Approved payment files are released to NPP gateway for processing
  and routed to receiving institutions"
  primary_category: Customer to Customer
  is_lifecycle_eligible: true
  primary_lifecycle_stage: Execution & Early Processing Assurance
  Why: "Released to NPP gateway for processing and routed" — payment has been
  authorised and is now being transmitted. Initiation is less appropriate —
  authorisation already completed. Clearing/Settlement is less appropriate —
  settlement has not yet occurred.

Example 3 — Clearing / Settlement
  L3 Activity: Settle Daily Interbank Obligations via RITS
  Description: "Net interbank settlement obligations are submitted to RITS and
  ESA accounts are debited and credited to complete settlement"
  primary_category: Institution to Institution
  is_lifecycle_eligible: true
  primary_lifecycle_stage: Clearing / Settlement
  Why: "ESA accounts are debited and credited to complete settlement" — the
  obligation between institutions is being discharged. Execution is less
  appropriate — this is post-transmission settlement, not payment release.

Example 4 — Posting & Accounting, Detection
  L3 Activity: Apply Direct Debit Repayments to Loan Accounts
  Description: "Direct debit repayments received are applied to customer loan
  account balances and general ledger entries are posted"
  primary_category: Customer to Institution
  is_lifecycle_eligible: true
  primary_lifecycle_stage: Posting & Accounting, Detection
  Why: "Applied to customer loan account balances and general ledger entries
  are posted" — this is post-settlement accounting. Execution is less appropriate
  — the payment has already been processed. Notification is less appropriate —
  no communication to customer occurs here.

Example 5 — Notification & Reporting
  L3 Activity: Send Payment Confirmation to Customer
  Description: "Customer receives confirmation of successful payment execution
  including payment reference and settlement details"
  primary_category: Customer to Institution
  is_lifecycle_eligible: true
  primary_lifecycle_stage: Notification & Reporting
  Why: "Receives confirmation of successful payment execution" — completed outcome
  is being communicated. Posting is less appropriate — accounting has already
  occurred upstream. Execution is less appropriate — payment has already settled.

Example 6 — Incident response, disputes, recovery followups
  L3 Activity: Investigate and Reissue Failed Direct Debit Payments
  Description: "Failed direct debit payments are identified, root cause investigated
  and payments are reissued or returned to originator"
  primary_category: Customer to Institution
  is_lifecycle_eligible: true
  primary_lifecycle_stage: Incident response, disputes, recovery followups
  Why: "Failed direct debit payments are identified, root cause investigated and
  payments are reissued or returned" — this is payment failure remediation.
  Execution is less appropriate — original execution already failed.
  Posting is less appropriate — no successful outcome to post.

Example 7 — NOT lifecycle eligible (contract/document lifecycle)
  L3 Activity: Prepare Loan Contract Documentation
  L2 Process: Prepare Product and Service Arrangement
  is_lifecycle_eligible: false
  primary_lifecycle_stage: null
  Why: "Prepare contract documentation" involves preparing terms, conditions and
  facility documents. This is product/contract lifecycle management — it does not
  receive, validate, authorise, execute, settle, post, notify or investigate a
  payment transaction. Setting up documentation for future payment capability is
  not payment processing.

Example 8 — NOT lifecycle eligible (account/facility setup)
  L3 Activity: Set-up Product Account / Facility
  is_lifecycle_eligible: false
  primary_lifecycle_stage: null
  Why: Creating and configuring accounts or facilities belongs to product/account
  lifecycle management. No payment instruction is received, validated, executed,
  settled, posted or investigated here.

Example 9 — NOT lifecycle eligible (procurement administration)
  L3 Activity: Process Supplier Invoice for Approval
  L2 Process: Manage Invoice
  is_lifecycle_eligible: false
  primary_lifecycle_stage: null
  Why: Invoice approval workflow is procurement lifecycle management. This activity
  manages supplier invoices for administrative approval — it does not directly
  process, execute, settle, reconcile or investigate a payment transaction or
  payment file. If a subsequent activity explicitly processes the payment run
  (e.g. generates payment file, executes EFT batch), that activity would be eligible.

Example 10 — Eligible despite generic L2 process name (Manage Data)
  L3 Activity: Reconcile Payment Transaction Data Against Settlement Records
  L2 Process: Manage Data (06.02.09)
  is_lifecycle_eligible: true
  primary_lifecycle_stage: Posting & Accounting, Detection
  Why: Despite a generic L2 name, the L3 activity "reconcile payment transaction
  data against settlement records" directly performs a payment reconciliation
  outcome. L3 activity name takes precedence over L2 process name per field
  precedence rules.

──────────────────────────────────────────────────────────────────────────────
IDENTIFIER RULE
──────────────────────────────────────────────────────────────────────────────

Each process includes l3_process_UUID. Copy it exactly into every output mapping.
Do not alter, trim, reformat or truncate.
If UUID is missing, use process_id exactly as supplied.

──────────────────────────────────────────────────────────────────────────────
FIELD PRECEDENCE
──────────────────────────────────────────────────────────────────────────────

When fields conflict, prefer in this order:
  1. l3_activity_name
  2. l3_activity_description
  3. tasks / task detail
  4. l3_activity_product_service
  5. l2_process_name / l2_process_description
  6. primary_category (from category run)
  7. channel / segment context
  8. payment rail / settlement mechanism (lowest priority — see Rail Neutrality)

Do not override a clear L3 activity name with a generic L2 process name.

──────────────────────────────────────────────────────────────────────────────
OUTPUT SCHEMA
──────────────────────────────────────────────────────────────────────────────

Return valid JSON only. No markdown. No prose outside the JSON object.

{
  "mappings": [
    {
      "l3_process_UUID": "copied exactly from input",
      "process_id": "copied from input",
      "l2_process_name": "string or null",
      "l3_activity_name": "string or null",
      "primary_category": "echoed from input",
      "is_lifecycle_eligible": true | false,
      "lifecycle_ineligible_reason": "explanation if false, else null",
      "primary_lifecycle_stage": "valid stage string or null",
      "secondary_lifecycle_stages": [],
      "confidence_score": "High | Medium | Low",
      "operational_outcome": "one sentence describing what the activity does",
      "why_this_stage": "why the primary stage fits this operational outcome",
      "why_not_adjacent_stages": "why the nearest alternative stages are less appropriate",
      "lifecycle_rationale": "concise rationale citing at least one quoted phrase from the process"
    }
  ],
  "notes": ["batch-level observations"]
}

──────────────────────────────────────────────────────────────────────────────
HARD RULES
──────────────────────────────────────────────────────────────────────────────

- Emit exactly one mapping entry per input process.
- Use only l3_process_UUIDs present in the input batch.
- Use only the six valid lifecycle stage strings exactly as listed.
- If is_lifecycle_eligible=false: primary_lifecycle_stage must be null.
- If is_lifecycle_eligible=true: primary_lifecycle_stage must be assigned.
- Every lifecycle_rationale must include at least one quoted phrase from the
  process payload (l3_activity_name, l3_activity_description, tasks, or systems).
- Do not force lifecycle stage for product/account/facility/contract/customer
  lifecycle management activities.
- Do not derive lifecycle stage from payment rail or mechanism alone.
- Do not derive lifecycle stage from inbound/outbound direction alone.
- where_not_adjacent_stages is mandatory for every lifecycle-eligible process.
- Do not invent facts. All reasoning must be grounded in supplied fields.
"""


# ──────────────────────────────────────────────────────────────────────────────
#  CHECKPOINT / RESUME
# ──────────────────────────────────────────────────────────────────────────────

def get_output_paths(output_dir: str) -> dict:
    ts     = datetime.now().strftime("%Y%m%d_%H%M%S")
    base   = Path(output_dir)
    base.mkdir(parents=True, exist_ok=True)
    suffix = f"lifecycle_mapping_{REASONING_EFFORT}"
    return {
        "dir":   base,
        "jsonl": base / f"{suffix}_raw.jsonl",
        "json":  base / f"{suffix}.json",
        "xlsx":  base / f"{suffix}.xlsx",
    }


def load_checkpoint(jsonl_path: Path) -> set:
    processed = set()
    if not jsonl_path.exists():
        return processed
    with open(jsonl_path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                rec = json.loads(line)
                for uuid in rec.get("input_l3_process_UUIDs", []):
                    processed.add(str(uuid))
            except json.JSONDecodeError:
                continue
    return processed


def write_batch_checkpoint(record: dict, jsonl_path: Path) -> None:
    safe = {k: v for k, v in record.items() if k != "raw_response"}
    try:
        with open(jsonl_path, "a", encoding="utf-8") as f:
            f.write(json.dumps(safe, ensure_ascii=True) + "\n")
    except (OSError, TypeError, ValueError) as e:
        print(f"  WARNING: Could not write batch {record.get('batch_no')} "
              f"to checkpoint: {e}")


def reconstruct_from_jsonl(jsonl_path: Path) -> tuple:
    batch_results, lookup, warnings = [], {}, []
    with open(jsonl_path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                rec = json.loads(line)
            except json.JSONDecodeError:
                continue
            batch_no = rec.get("batch_no")
            mappings = rec.get("parsed_response", {}).get("mappings", [])
            batch_results.append({
                "batch_no":   batch_no,
                "mappings":   mappings,
                "notes":      rec.get("parsed_response", {}).get("notes", []),
                "warnings":   rec.get("warnings", []),
                "latency_ms": rec.get("latency_ms"),
                "usage":      rec.get("usage"),
            })
            lookup[batch_no] = rec.get("source_review_fields", {})
            warnings.extend([f"Batch {batch_no}: {w}" for w in rec.get("warnings", [])])
    return batch_results, lookup, warnings


# ──────────────────────────────────────────────────────────────────────────────
#  DATA LOADING
# ──────────────────────────────────────────────────────────────────────────────

def load_direct_processes() -> list:
    """
    Load source process records and join with category output.
    Returns list of enriched process dicts for Direct processes only.
    """
    # Load original process candidates
    print(f"  Loading process source: {PROCESS_INPUT_FILE}")
    with open(PROCESS_INPUT_FILE, "r", encoding="utf-8") as f:
        raw = json.load(f)
    processes = raw["processes"] if isinstance(raw, dict) and "processes" in raw else raw
    process_lookup = {str(p.get("l3_process_UUID", p.get("process_id", ""))): p
                      for p in processes}
    print(f"  Loaded {len(processes):,} source processes")

    # Load category output
    print(f"  Loading category output: {CATEGORY_OUTPUT_FILE}")
    with open(CATEGORY_OUTPUT_FILE, "r", encoding="utf-8") as f:
        cat_raw = json.load(f)
    cat_mappings = cat_raw.get("mappings", [])
    print(f"  Loaded {len(cat_mappings):,} category mappings")

    # Join and filter to Direct
    enriched = []
    for cat in cat_mappings:
        if cat.get("payment_process_type") != "Direct":
            continue
        uuid = str(cat.get("l3_process_UUID", cat.get("process_id", "")))
        source = process_lookup.get(uuid, {})

        # Build enriched payload — source fields take priority for descriptions,
        # category fields provide classification context
        record = {}

        # From source (original process fields)
        for field in SOURCE_FIELDS:
            val = source.get(field, "")
            record[field] = str(val).strip() if val and str(val).strip() not in ("nan", "None", "") else None

        # From category output (classification context)
        record["l3_process_UUID"]      = uuid
        record["process_id"]           = cat.get("process_id")
        record["payment_process_type"] = cat.get("payment_process_type")
        record["primary_category"]     = cat.get("primary_category")
        record["mapped_categories"]    = cat.get("mapped_categories")
        record["payment_process_rationale"] = cat.get("payment_process_rationale")

        # Prefer source descriptions over category-echoed ones if richer
        if not record.get("l3_activity_description") and cat.get("l3_activity_description"):
            record["l3_activity_description"] = cat.get("l3_activity_description")
        if not record.get("l2_process_name") and cat.get("l2_process_name"):
            record["l2_process_name"] = cat.get("l2_process_name")

        enriched.append(record)

    print(f"  Direct processes to classify: {len(enriched):,}")
    return enriched


# ──────────────────────────────────────────────────────────────────────────────
#  CORE UTILITIES
# ──────────────────────────────────────────────────────────────────────────────

def preflight_check():
    print("\n  Pre-flight environment check")
    print("  " + "─" * 58)
    for var in ("REQUESTS_CA_BUNDLE", "SSL_CERT_FILE",
                "HTTP_PROXY", "HTTPS_PROXY", "AZURE_CA_BUNDLE"):
        print(f"  {var:<22} = {os.getenv(var) or '(not set)'}")
    if os.getenv("AZURE_CA_BUNDLE"):
        print("  WARNING: AZURE_CA_BUNDLE is set — recommend removing.")


def init_client() -> AzureOpenAI:
    token_provider = get_bearer_token_provider(
        DefaultAzureCredential(),
        "https://cognitiveservices.azure.com/.default",
    )
    return AzureOpenAI(
        azure_endpoint=ENDPOINT,
        api_version=API_VERSION,
        azure_ad_token_provider=token_provider,
    )


def chunk_list(items: list, size: int):
    for i in range(0, len(items), size):
        yield i // size + 1, items[i:i + size]


def parse_json_response(text: str) -> dict:
    text = (text or "").strip()
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        m = re.search(r"\{.*\}", text, flags=re.DOTALL)
        if m:
            return json.loads(m.group(0))
        raise


def get_usage(response) -> dict:
    usage = {}
    if hasattr(response, "usage") and response.usage:
        u = response.usage
        usage["input_tokens"]  = getattr(u, "prompt_tokens", None)
        usage["output_tokens"] = getattr(u, "completion_tokens", None)
        usage["total_tokens"]  = getattr(u, "total_tokens", None)
        details = getattr(u, "completion_tokens_details", None)
        if details:
            usage["reasoning_tokens"] = getattr(details, "reasoning_tokens", None)
    return usage


# ──────────────────────────────────────────────────────────────────────────────
#  VALIDATION
# ──────────────────────────────────────────────────────────────────────────────

def validate_output(parsed: dict, batch: list) -> list:
    warnings = []
    input_ids  = {str(p.get("l3_process_UUID", p.get("process_id", ""))) for p in batch}
    mappings   = parsed.get("mappings", [])
    output_ids = {str(m.get("l3_process_UUID", m.get("process_id", ""))) for m in mappings}

    missing = sorted(input_ids - output_ids)
    extra   = sorted(output_ids - input_ids)
    if missing:
        warnings.append(f"Missing UUIDs in output: {missing}")
    if extra:
        warnings.append(f"Unknown UUIDs returned: {extra}")
    if len(mappings) != len(batch):
        warnings.append(f"Expected {len(batch)} mappings, received {len(mappings)}")

    for m in mappings:
        uid   = m.get("l3_process_UUID", m.get("process_id", ""))
        stage = m.get("primary_lifecycle_stage")
        elig  = m.get("is_lifecycle_eligible")

        if stage and stage not in VALID_LIFECYCLE_STAGES:
            warnings.append(f"Invalid lifecycle stage '{stage}' for UUID {uid}")

        if elig is True and not stage:
            warnings.append(f"is_lifecycle_eligible=true but no stage for UUID {uid}")

        if elig is False and stage:
            warnings.append(
                f"is_lifecycle_eligible=false but stage '{stage}' returned for UUID {uid}"
            )

    return warnings


# ──────────────────────────────────────────────────────────────────────────────
#  AUGMENTATION
# ──────────────────────────────────────────────────────────────────────────────

def verify_rationale_citation(rationale: str, process: dict) -> bool:
    if not rationale:
        return False
    source_parts = []
    for field in SOURCE_FIELDS + ["l3_activity_name", "l3_activity_description"]:
        val = process.get(field)
        if val:
            source_parts.append(str(val))
    source_text = " ".join(source_parts).lower()
    quoted = re.findall(r'"([^"]{4,})"', rationale)
    if not quoted:
        return False
    return any(phrase.lower() in source_text for phrase in quoted)


def augment_mapping(mapping: dict, process: dict) -> dict:
    m = dict(mapping)

    # Citation verification
    rationale = m.get("lifecycle_rationale", "") or ""
    m["citation_verified"] = verify_rationale_citation(rationale, process)

    # SME review flag
    m["sme_review_flag"] = (
        not m["citation_verified"]
        or m.get("confidence_score") == "Low"
        or not m.get("why_not_adjacent_stages")
    )

    return m


def check_batch_distribution(mappings: list) -> list:
    warnings = []
    stages = [m.get("primary_lifecycle_stage") for m in mappings
              if m.get("primary_lifecycle_stage")]
    if len(stages) < 2:
        return warnings
    for stage, count in Counter(stages).items():
        if count / len(mappings) > 0.8:
            warnings.append(
                f"DISTRIBUTION WARNING: '{stage}' = {count}/{len(mappings)} in batch. "
                f"Review for model drift."
            )
    return warnings


# ──────────────────────────────────────────────────────────────────────────────
#  LLM CALL
# ──────────────────────────────────────────────────────────────────────────────

def call_llm(client: AzureOpenAI, batch: list, batch_no: int) -> tuple:
    prompt = (
        PROMPT_RULES
        + "\n\nProcesses to assess\n"
        + json.dumps(batch, ensure_ascii=False, indent=2)
    )
    kwargs = {
        "model":                 MODEL,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user",   "content": prompt},
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
            parsed   = parse_json_response(text)
            usage    = get_usage(response)
            return parsed, {"batch_no": batch_no, "latency_ms": latency,
                            "usage": usage, "raw_response": text}
        except Exception as exc:
            if attempt < RETRY_COUNT:
                sleep_t = min(RETRY_BASE_SLEEP * (2 ** (attempt - 1)), RETRY_MAX_SLEEP)
                sleep_t += random.uniform(0, 2)
                print(f"    Attempt {attempt} failed: {exc}. Retrying in {sleep_t:.1f}s...")
                time.sleep(sleep_t)
            else:
                raise


# ──────────────────────────────────────────────────────────────────────────────
#  OUTPUT
# ──────────────────────────────────────────────────────────────────────────────

def flatten_results(batch_results: list, lookup: dict) -> pd.DataFrame:
    rows = []
    for result in batch_results:
        usage    = result.get("usage") or {}
        batch_no = result.get("batch_no")
        src_map  = lookup.get(batch_no, {})

        for m in result.get("mappings", []):
            uuid = str(m.get("l3_process_UUID", m.get("process_id", "")))
            src  = src_map.get(uuid, {})

            rows.append({
                # Identifiers
                "l3_process_UUID":          m.get("l3_process_UUID"),
                "process_id":               m.get("process_id"),
                # Review context from source
                "l2_process_name":          src.get("l2_process_name", m.get("l2_process_name")),
                "l2_process_description":   src.get("l2_process_description", ""),
                "l3_activity_name":         src.get("l3_activity_name", m.get("l3_activity_name")),
                "l3_activity_description":  src.get("l3_activity_description", ""),
                "l3_activity_product_service": src.get("l3_activity_product_service", ""),
                # Category context (from category run)
                "primary_category":         m.get("primary_category"),
                # Lifecycle results
                "is_lifecycle_eligible":    m.get("is_lifecycle_eligible"),
                "lifecycle_ineligible_reason": m.get("lifecycle_ineligible_reason"),
                "primary_lifecycle_stage":  m.get("primary_lifecycle_stage"),
                "secondary_lifecycle_stages": " | ".join(m.get("secondary_lifecycle_stages") or []),
                "confidence_score":         m.get("confidence_score"),
                "operational_outcome":      m.get("operational_outcome"),
                "why_this_stage":           m.get("why_this_stage"),
                "why_not_adjacent_stages":  m.get("why_not_adjacent_stages"),
                "lifecycle_rationale":      m.get("lifecycle_rationale"),
                # Augmentation
                "citation_verified":        m.get("citation_verified"),
                "sme_review_flag":          m.get("sme_review_flag"),
                # Token usage
                "batch_no":                 batch_no,
                "input_tokens":             usage.get("input_tokens"),
                "output_tokens":            usage.get("output_tokens"),
                "reasoning_tokens":         usage.get("reasoning_tokens"),
            })
    return pd.DataFrame(rows)


def build_summary(df: pd.DataFrame) -> pd.DataFrame:
    total     = len(df)
    eligible  = df["is_lifecycle_eligible"].sum() if "is_lifecycle_eligible" in df else 0
    not_elig  = total - int(eligible)
    sme_flags = df["sme_review_flag"].sum() if "sme_review_flag" in df else 0
    no_cite   = (~df["citation_verified"]).sum() if "citation_verified" in df else 0

    stage_dist = {}
    if "primary_lifecycle_stage" in df:
        stage_dist = (df[df["primary_lifecycle_stage"].notna()]
                      ["primary_lifecycle_stage"].value_counts().to_dict())

    conf_dist = {}
    if "confidence_score" in df:
        conf_dist = df["confidence_score"].value_counts().to_dict()

    rows = [
        ("Total Direct processes",      total),
        ("Lifecycle eligible",           int(eligible)),
        ("Not lifecycle eligible",       not_elig),
        ("SME review flagged",           int(sme_flags)),
        ("Citation unverified",          int(no_cite)),
        ("──────────────────────────",  "──────────────"),
    ]
    for stage, cnt in stage_dist.items():
        rows.append((f"Stage: {stage}", cnt))
    rows.append(("──────────────────────────", "──────────────"))
    for conf, cnt in conf_dist.items():
        rows.append((f"Confidence: {conf}", cnt))

    return pd.DataFrame(rows, columns=["Metric", "Count"])


def write_outputs(batch_results: list, lookup: dict, all_warnings: list,
                  paths: dict, config: dict) -> None:
    df_all = flatten_results(batch_results, lookup)

    # JSON
    combined = {
        "run_ts":        datetime.now().isoformat(timespec="seconds"),
        "config":        config,
        "warnings":      all_warnings,
        "mappings":      [m for r in batch_results for m in r.get("mappings", [])],
        "batch_metadata":[{
            "batch_no":   r.get("batch_no"),
            "latency_ms": r.get("latency_ms"),
            "usage":      r.get("usage"),
            "warnings":   r.get("warnings", []),
        } for r in batch_results],
    }
    with open(paths["json"], "w", encoding="utf-8") as f:
        json.dump(combined, f, ensure_ascii=True, indent=2)

    # Excel
    with pd.ExcelWriter(paths["xlsx"], engine="openpyxl") as writer:
        df_all.to_excel(writer, index=False, sheet_name="mappings")

        if "is_lifecycle_eligible" in df_all.columns:
            df_all[df_all["is_lifecycle_eligible"] == True].to_excel(
                writer, index=False, sheet_name="lifecycle_eligible")
            df_all[df_all["is_lifecycle_eligible"] == False].to_excel(
                writer, index=False, sheet_name="not_eligible")

        if "sme_review_flag" in df_all.columns:
            df_all[df_all["sme_review_flag"] == True].to_excel(
                writer, index=False, sheet_name="sme_review")

        build_summary(df_all).to_excel(writer, index=False, sheet_name="summary")

        pd.DataFrame(combined["batch_metadata"]).to_excel(
            writer, index=False, sheet_name="batch_metadata")

        pd.DataFrame({"warning": all_warnings}).to_excel(
            writer, index=False, sheet_name="warnings")

    print(f"\n  Outputs saved:")
    print(f"  JSON  → {paths['json']}")
    print(f"  Excel → {paths['xlsx']}")
    elig_count = (df_all["is_lifecycle_eligible"] == True).sum() if "is_lifecycle_eligible" in df_all.columns else 0
    sme_count  = (df_all["sme_review_flag"] == True).sum() if "sme_review_flag" in df_all.columns else 0
    print(f"\n  Lifecycle eligible : {elig_count}/{len(df_all)}")
    print(f"  SME review queue   : {sme_count}/{len(df_all)}")


# ──────────────────────────────────────────────────────────────────────────────
#  MAIN
# ──────────────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Lifecycle stage mapping for Direct payment processes."
    )
    parser.add_argument("--force", action="store_true",
                        help="Ignore checkpoint and reprocess from scratch.")
    args = parser.parse_args()

    paths = get_output_paths(OUTPUT_DIR)

    print("=" * 72)
    print("  Payment Process Lifecycle Mapper")
    print(f"  Model             : {MODEL}")
    print(f"  Reasoning effort  : {REASONING_EFFORT}")
    print(f"  Batch size        : {BATCH_SIZE}")
    print(f"  Process input     : {PROCESS_INPUT_FILE}")
    print(f"  Category output   : {CATEGORY_OUTPUT_FILE}")
    print(f"  JSONL checkpoint  : {paths['jsonl']}")
    print("=" * 72)

    # Load and join processes
    all_processes = load_direct_processes()
    if not all_processes:
        print("  No Direct processes found. Check input files.")
        return

    preflight_check()

    # Checkpoint
    if args.force and paths["jsonl"].exists():
        paths["jsonl"].unlink()
        print("\n  --force: checkpoint deleted. Starting from scratch.")

    processed_uuids = load_checkpoint(paths["jsonl"])
    if processed_uuids:
        print(f"\n  Checkpoint found: {len(processed_uuids):,} already processed.")

    processes = [
        p for p in all_processes
        if str(p.get("l3_process_UUID", p.get("process_id", ""))) not in processed_uuids
    ]

    if not processes:
        print("  All processes completed. Regenerating outputs...")
        batch_results, lookup, all_warnings = reconstruct_from_jsonl(paths["jsonl"])
        write_outputs(batch_results, lookup, all_warnings, paths, {
            "model": MODEL, "reasoning_effort": REASONING_EFFORT,
            "batch_size": BATCH_SIZE,
        })
        print("\n  Done.")
        return

    total_batches = (len(processes) + BATCH_SIZE - 1) // BATCH_SIZE
    print(f"\n  Remaining : {len(processes):,} | Completed: {len(processed_uuids):,}")
    print(f"  Batches   : {total_batches:,}\n")

    print("  Initialising Azure OpenAI client...")
    client = init_client()
    print("  Client ready.\n")

    batch_results, all_warnings_run = [], []
    total_in_tok, total_out_tok = 0, 0
    run_start = time.time()

    for batch_idx, (batch_no, batch) in enumerate(chunk_list(processes, BATCH_SIZE), 1):
        pct     = batch_idx / total_batches * 100
        bar     = "#" * int(pct / 5) + "." * (20 - int(pct / 5))
        elapsed = time.time() - run_start
        eta_s   = (elapsed / batch_idx) * (total_batches - batch_idx) if batch_idx > 1 else 0
        eta_str = (f"{int(eta_s//3600)}h {int((eta_s%3600)//60)}m"
                   if eta_s > 60 else f"{int(eta_s)}s")
        cost_aud = ((total_in_tok / 1e6 * INPUT_PRICE_USD_PER_M +
                     total_out_tok / 1e6 * OUTPUT_PRICE_USD_PER_M) / AUD_USD_RATE)

        print(f"  [{bar}] {pct:5.1f}%  Batch {batch_idx}/{total_batches}  "
              f"ETA {eta_str}  Cost A${cost_aud:.2f}")

        parsed, meta = call_llm(client, batch, batch_no)

        struct_warnings = validate_output(parsed, batch)

        # Augment mappings
        proc_lookup = {str(p.get("l3_process_UUID", p.get("process_id", ""))): p
                       for p in batch}
        augmented = []
        for m in parsed.get("mappings", []):
            uid = str(m.get("l3_process_UUID", m.get("process_id", "")))
            augmented.append(augment_mapping(m, proc_lookup.get(uid, {})))
        parsed["mappings"] = augmented

        dist_warnings  = check_batch_distribution(augmented)
        batch_warnings = struct_warnings + dist_warnings

        result = {
            "batch_no":   batch_no,
            "mappings":   augmented,
            "notes":      parsed.get("notes", []),
            "warnings":   batch_warnings,
            "latency_ms": meta.get("latency_ms"),
            "usage":      meta.get("usage"),
        }
        batch_results.append(result)

        # Source review fields for JSONL reconstruction
        source_review = {
            str(p.get("l3_process_UUID", p.get("process_id", ""))): {
                k: str(p.get(k, "") or "") for k in
                ["l2_process_name", "l2_process_description",
                 "l3_activity_name", "l3_activity_description",
                 "l3_activity_product_service"]
            }
            for p in batch
        }
        raw_record = {
            "batch_no":               batch_no,
            "input_l3_process_UUIDs": [str(p.get("l3_process_UUID",
                                           p.get("process_id", ""))) for p in batch],
            "source_review_fields":   source_review,
            "parsed_response":        parsed,
            "usage":                  meta.get("usage"),
            "latency_ms":             meta.get("latency_ms"),
            "warnings":               batch_warnings,
        }
        write_batch_checkpoint(raw_record, paths["jsonl"])
        all_warnings_run.extend([f"Batch {batch_no}: {w}" for w in batch_warnings])

        usage      = result["usage"] or {}
        in_tok     = usage.get("input_tokens") or 0
        out_tok    = usage.get("output_tokens") or 0
        total_in_tok  += in_tok
        total_out_tok += out_tok
        sme_count  = sum(1 for m in augmented if m.get("sme_review_flag"))
        rt = (f" | {usage.get('reasoning_tokens')} thinking"
              if usage.get("reasoning_tokens") else "")
        print(f"         {len(augmented)} mappings  "
              f"{result['latency_ms']}ms  "
              f"{in_tok} in / {out_tok} out{rt}  "
              f"SME: {sme_count}")
        if batch_warnings:
            for w in batch_warnings:
                print(f"         WARNING: {w}")

        if INTER_BATCH_SLEEP > 0:
            time.sleep(INTER_BATCH_SLEEP)

    # Final outputs from complete JSONL
    print("\n  Generating final outputs from complete JSONL...")
    all_batch_results, all_lookup, all_warnings_full = reconstruct_from_jsonl(paths["jsonl"])

    final_cost_aud = ((total_in_tok / 1e6 * INPUT_PRICE_USD_PER_M +
                       total_out_tok / 1e6 * OUTPUT_PRICE_USD_PER_M) / AUD_USD_RATE)
    print(f"  This run: {len(processes):,} processes  "
          f"{total_in_tok:,} input  {total_out_tok:,} output  "
          f"A${final_cost_aud:.2f}")

    write_outputs(all_batch_results, all_lookup, all_warnings_full, paths, {
        "model": MODEL, "reasoning_effort": REASONING_EFFORT,
        "batch_size": BATCH_SIZE,
        "process_input_file":   PROCESS_INPUT_FILE,
        "category_output_file": CATEGORY_OUTPUT_FILE,
    })
    print("\n  Done.")


if __name__ == "__main__":
    main()
