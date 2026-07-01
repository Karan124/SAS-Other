r"""
run_payment_relevance_and_category_llm_uuid.py
───────────────────────────────────────────────────────────────────────────────
Payments Controls PoC — Process Payment Relevance & Category Classifier
Uses l3_process_UUID as the stable primary join key.

Changes from previous version:
  - reasoning_effort   : high     → medium   (structured classification, not open-ended)
  - BATCH_SIZE         : 5        → 3        (more reasoning budget per process)
  - MAX_COMPLETION_TOKENS: 12000  → 16000    (safety margin for reasoning tokens)
  - Retry logic        : flat sleep → exponential backoff with jitter
  - Credential         : AzureCliCredential → DefaultAzureCredential (portable)
  - Prompt             : improved v2 — Step 0 exclusion gate, conservative bias,
                         contrastive reasoning, worked examples, detailed rules
  - Augmentations added (post-processing, no extra LLM calls):
      1. objective_confidence  — derived from rule hit count (R1/R2/R3 = strong)
      2. citation_verified     — quoted phrases in rationale verified against source
      3. field_coverage_pct/tier — how many source fields were populated
      4. confidence_conflict   — stated High but objective Low → flag
      5. sme_review_flag       — any concern triggers SME review
      6. Batch distribution warning — >70% same category → potential model drift
  - New Excel sheet: "sme_review_queue" — all flagged rows in one place

Before running (PowerShell):
  az account set --subscription 6c72e6c5-ed48-4030-b29c-34e2849c9288
  $env:REQUESTS_CA_BUNDLE = "C:\Users\m061400\ai-test\cacert.pem"
  $env:SSL_CERT_FILE      = "C:\Users\m061400\ai-test\cacert.pem"
  Remove-Item Env:AZURE_CA_BUNDLE -ErrorAction SilentlyContinue
  python run_payment_relevance_and_category_llm_uuid.py [input_file]
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
#  CONFIG
# ──────────────────────────────────────────────────────────────────────────────

ENDPOINT    = "https://ai.eng.azure.srv.westpac.com.au"
API_VERSION = "2024-10-21"
MODEL       = "gpt-5.4"

DEFAULT_INPUT_FILE = (
    r"Z:\Enterprise Risk Insights\23 _sas_batch_\Controls-PoC\processes"
    r"\holo_process_samples_aggregated_for_llm_uuid.json"
)

BATCH_SIZE             = 3        # was 5 — smaller = more reasoning per process
MAX_COMPLETION_TOKENS  = 16000    # was 12000 — reasoning tokens need headroom
REASONING_EFFORT       = "medium" # was "high" — structured classification, not agentic

RETRY_COUNT      = 3
RETRY_BASE_SLEEP = 5    # seconds — base for exponential backoff
RETRY_MAX_SLEEP  = 60   # seconds — cap

VALID_CATEGORIES = [
    "Customer to Customer",
    "Customer to Institution",
    "Institution to Customer",
    "Institution to Institution",
    "Supplier / Contractor / Employee Payments",
]

VALID_PROCESS_TYPES = ["Direct", "Enabling", "Non-payment"]

# Fields used to calculate source field coverage (augmentation 3)
# Adjust to match the actual keys in your input JSON payload
RELEVANT_SOURCE_FIELDS = [
    "l2_process_name",
    "l3_activity_name",
    "description",
    "tasks",
    "systems",
    "third_parties",
    "governance_context",
    "value_stream_names",
    "vcm_library_names",
    "business_capabilities_bcm",
    "l3_activity_channels",            # NEW — where the process originates/takes place
    "l3_activity_customer_segments",   # NEW — parties involved in the process
]

# Fields specifically used as payer-payee corroborating evidence.
# Used to populate channel_segment_used cross-check (augmentation).
PAYER_PAYEE_EVIDENCE_FIELDS = [
    "l3_activity_channels",
    "l3_activity_customer_segments",
]

# ──────────────────────────────────────────────────────────────────────────────
#  SYSTEM PROMPT
# ──────────────────────────────────────────────────────────────────────────────

SYSTEM_PROMPT = (
    "You are a payments process assessment and category mapping analyst for an "
    "Australian ADI (Authorised Deposit-taking Institution).\n"
    "Return valid JSON only. No markdown. No prose outside the JSON object.\n"
    "Do not invent facts. Do not use knowledge outside the supplied process payload.\n"
    "Classify processes as Direct, Enabling, or Non-payment. "
    "Payment-enabling processes (IT, BCP, governance, change, supplier) are "
    "payment-related even without explicit payment keywords — classify them.\n"
    "Every rationale must include at least one quoted phrase from the supplied payload."
)

# ──────────────────────────────────────────────────────────────────────────────
#  PROMPT — IMPROVED V2
#  Combines:
#    - Two-step structure + l3_process_UUID from the previous version
#    - Detailed R1-R6 rules with include/exclude criteria from the original prompt
#    - New: Step 0 exclusion gate, conservative bias, contrastive reasoning,
#           7 worked examples, explicit confidence thresholds
# ──────────────────────────────────────────────────────────────────────────────

PROMPT_RULES = """
Role

For each process in the batch, perform three steps in sequence.

Core principle: Classify all processes as Direct, Enabling, or Non-payment.
Only Non-payment processes are excluded from category mapping. Do not stop
evaluation due to absence of explicit payment keywords — many legitimate
payment-related processes at a bank support payment infrastructure without
naming a specific payment system in their description.

For all payment-related processes (Direct or Enabling), exactly one
primary_category must be assigned. Do not return null primary_category or
empty mapped_categories for any payment-related process. If no category is
directly supported, assign the closest alternative category with a Low or
Medium confidence score and explain the basis for the inference.

Do not invent facts. All inference must be grounded in available fields
(description, tasks, systems, channels, customer segments, L2 context).

──────────────────────────────────────────────────────────────────────────────
STEP 0 — NON-PAYMENT EXCLUSION  (narrow filter — apply only with clear evidence)
──────────────────────────────────────────────────────────────────────────────

Mark is_payment_process=false, payment_process_type="Non-payment", and
set exclusion_gate_applied=true ONLY where the process has CLEAR EVIDENCE
of no linkage to any payment system, payment processing, payment governance,
or payment-scoped operations AND the process is unambiguously one of:

  • Pure strategy or portfolio planning with no payment system reference
  • Pure product design/catalogue management with no payment processing link
  • Pure marketing or communications unrelated to payment execution
  • Pure HR/workforce activity that does NOT involve paying anyone
    (training, performance reviews, recruitment) —
    NOTE: payroll, staff payment runs, and employee disbursements are
    NOT excluded — they are a valid payment category
  • Pure procurement administration with no payment execution link
    (vendor selection, contract negotiation) — actual vendor PAYMENT
    processing is not excluded
  • Pure legal/regulatory affairs with no payment system scope
  • Pure financial reporting/management accounting with no payment
    transaction processing link

Do NOT use this gate for: IT security, ITGC, change management, access
management, monitoring, governance, business continuity, resilience, or
supplier oversight processes. These are payment-ENABLING by nature at a bank
and must be classified as Enabling in Step 1, not excluded here.
Do not stop at Step 0 solely because a process does not mention a specific
named payment system. An IT security or BCP process at a bank inherently
applies to payment systems — that is the context of the institution.

──────────────────────────────────────────────────────────────────────────────
STEP 1 — THREE-TIER CLASSIFICATION
──────────────────────────────────────────────────────────────────────────────

Classify each process as one of three types. Set payment_process_type to
exactly one of: "Direct", "Enabling", or "Non-payment".

DIRECT — is_payment_process=true, payment_process_type="Direct"
  Process directly executes, processes, settles, reconciles, posts, routes,
  authorises, amends, cancels, or reports on payment transactions.
  Evidence: explicit payment action verbs, named payment systems, payer-payee
  relationships, payment instruments (Direct Debit, BPAY, RTGS, SWIFT, etc.).

ENABLING — is_payment_process=true, payment_process_type="Enabling"
  Process supports the infrastructure, security, governance, continuity, or
  third-party relationships that enable payment operations to function.
  Include all of the following at a bank, even when no named payment system
  appears in the description:
    - IT security / ITGC (vulnerability management, pen testing, access
      management, patch management, secure design) — these protect payment
      system infrastructure
    - Business continuity / resilience / disaster recovery — these ensure
      payment operations continue during disruptions
    - Change management for technology or operations — changes at a bank
      affect payment systems
    - Third-party and supplier governance for technology or operational
      service providers — these providers support payment execution
    - Incident and crisis management — payment incidents are managed here
    - Risk oversight and assurance for operational or technology risk —
      payment risk is an operational risk at a bank
  Enabling processes DO proceed to Step 2 for category mapping.
  Assign category based on the dominant payment flow the process most
  likely supports, using segment, channel, and L2 context as signals.

NON-PAYMENT — is_payment_process=false, payment_process_type="Non-payment"
  Apply only where there is clear evidence of no linkage to payments.
  See Step 0 for the explicit exclusion list. Do not use this classification
  for IT, BCP, governance, or change management processes.

Confidence for payment_process_confidence:
  High   — Explicit evidence: named payment system, payment action verb,
            or internal disbursement reference.
  Medium — Payment relevance strongly implied by context, segment, channel,
            or L2 process name, but no explicit payment keyword.
  Low    — Payment relevance inferred from the nature of the process type
            (e.g. IT security at a bank classified as Enabling). Acceptable —
            these are flagged for SME review.

If payment_process_type="Non-payment":
  Set is_payment_process=false, mapped_categories=[], primary_category=null,
  category_confidence_score=Low. Do not proceed to Step 2.

──────────────────────────────────────────────────────────────────────────────
STEP 2 — PAYMENT CATEGORY MAPPING  (all Direct and Enabling processes)
──────────────────────────────────────────────────────────────────────────────

MANDATORY: For every process where is_payment_process=true (whether Direct
or Enabling), you MUST assign exactly one primary_category. Do not return
null primary_category or empty mapped_categories for any payment-related
process.

If direct evidence does not support a category:
  1. Use customer segment and channel to determine the most likely category.
  2. Use L2 process name and context to infer the dominant payment flow.
  3. If still ambiguous, assign the closest alternative category with
     category_confidence_score=Low and explain the inference basis.
  4. For IT/BCP/governance/change/supplier processes (Enabling type):
     - If "Institutional" segment or interbank/infrastructure context present
       → Institution to Institution
     - If "Consumer" or "Commercial" segment or retail context present
       → Customer to Customer or Customer to Institution (based on flow)
     - If L2 context is lending/disbursement → Institution to Customer
     - If L2 context is payroll/vendor → Supplier/Contractor/Employee
     - If no segment/channel/L2 signal at all → select the category
       that best represents the dominant payment use case at the bank
       for this type of operational process, at Low confidence.

Valid Payment Categories (use only these exact strings):
  - Customer to Customer
  - Customer to Institution
  - Institution to Customer
  - Institution to Institution
  - Supplier / Contractor / Employee Payments

Category Definitions:

  Customer to Customer
  Individual, merchant, or business pays individual, merchant, or business.
  The bank is an intermediary, not a principal party.
  Signals: PayID, P2P payments, card payments, direct debit, NPP/Osko,
  Electronic Fund Transfer, internal transfers between customers.

  Customer to Institution
  Customer pays the bank or deposits/repays funds into the institution.
  Signals: Loan repayments, card repayments, ATM deposits, contactless
  payments, cheques deposited, BPAY bill payments to the bank.

  Institution to Customer
  The bank or institution disburses or pays funds to the customer.
  Signals: Loan disbursements, interest payments, file-based bulk payments,
  SWIFT international transfers, FX remittances, tax payments,
  superannuation, welfare/benefit payments, payroll processed on behalf
  of a client institution.

  Institution to Institution
  One financial institution pays, settles, clears, or exchanges obligations
  with another institution.
  Signals: Interbank clearing, deal settlements, correspondent banking,
  Visa/Mastercard settlement, POS/merchant settlements, batch merchant
  settlements, gateway/e-commerce settlement, RITS/RTGS.

  Supplier / Contractor / Employee Payments
  The bank pays its own suppliers, contractors, or employees.
  Signals: Vendor payments, internal staff payroll, contractor invoices,
  supplier expense reimbursements.

Note: Signals are illustrative examples. They do not override process evidence.

──────────────────────────────────────────────────────────────────────────────
USING CHANNEL AND CUSTOMER SEGMENT AS PAYER-PAYEE EVIDENCE
──────────────────────────────────────────────────────────────────────────────

Each process may include two additional fields that help determine the
payer-payee relationship:

  l3_activity_channels
  Where the process takes place or originates (e.g. "Digital (Online and
  Mobile)", "Branch", "Relationship Manager", "ATM", "GTS Direct
  Connectivity"). Use this to corroborate WHO is initiating or touching
  the process, not to determine category alone.

  l3_activity_customer_segments
  The parties/customer types involved (e.g. "Consumer", "Consumer - High
  Net Worth", "Commercial", "Small Business", "Institutional"). Use this
  to help identify WHO the payer or payee is.

How to use these fields:
  - "Institutional" segment + a settlement/clearing context → supports
    Institution to Institution.
  - "Consumer" or "Commercial" segment + a deposit/repayment/ATM/BPAY
    context → supports Customer to Institution.
  - "Consumer" segment + a disbursement/refund/interest context →
    supports Institution to Customer.
  - "Consumer", "Commercial", or "Small Business" segment + a transfer/
    P2P/card/direct-debit context between non-bank parties → supports
    Customer to Customer.
  - Channel alone (e.g. "Branch", "ATM", "Digital") does NOT determine
    category — it only indicates where the interaction happens. Combine
    channel with segment and process action to determine the category.
  - These fields STRENGTHEN a mapping already suggested by R1-R3 evidence.
    They must not be the SOLE basis for a mapping — always corroborate
    with the process name, description, or tasks. Cite the specific
    channel or segment value used in your rationale when it contributes
    to the decision.
  - If these fields are absent or generic (e.g. blank, or list every
    segment with no distinguishing pattern), do not rely on them — fall
    back to R1-R6 evidence only.

Minimum Mapping Requirement

Category mapping requires at least one of:
  (a) Explicit payer-payee relationship
  (b) Explicit payment action - settle, disburse, process payment, etc.
  (c) Explicit payment system/process reference - RTGS, BPAY, clearing, etc.
  (d) Explicit system, third-party, or governance reference in the context
      of payment processing
  (e) A combination of customer segment + channel + process action that
      together unambiguously imply a payer-payee relationship

  (f) Payment instrument type (e.g. direct debit, BPAY, EFT, cheque)
      combined with transaction direction (inbound vs outbound) and L2/L3
      process context — together these can unambiguously imply a category
      even when no explicit payer-payee wording exists
  (g) For Enabling processes: the nature of the operational function at a
      bank combined with customer segment, channel, or L2 process context

For Direct processes with no category evidence: return empty and explain.
For Enabling processes: a primary_category is always required — use the
closest defensible alternative based on L2 context and segment signals.

──────────────────────────────────────────────────────────────────────────────
INSTRUMENT AND DIRECTION INFERENCE  (apply when L3 activity is generic)
──────────────────────────────────────────────────────────────────────────────

Do NOT rely only on explicit payer-payee wording in the L3 activity name.
If the L3 activity is a generic payment processing step (e.g. receive,
validate, process inbound payment), you MUST apply the following:

  Step A — Identify payment instrument / type
  This is the strongest single signal.
  e.g. direct debit, BPAY, EFT, cheque, NPP, RTGS, SWIFT, PEXA, Visa, etc.

  Step B — Identify transaction direction
  Inbound (funds coming into the bank or a customer account) vs outbound
  (funds leaving the bank or a customer account).

  Step C — Identify L2 process context
  e.g. lending, supplier payments, customer transfers, payroll, clearing.

  Step D — Determine economic purpose of the payment flow
  e.g. repayment, disbursement, transfer, settlement, payroll.

Deterministic signal priority mapping:
  - Direct Debit / BPAY                       → Customer to Institution
  - Inbound + lending / loan context           → Customer to Institution
  - Outbound + lending / loan context          → Institution to Customer
  - Consumer EFT / transfer                    → Customer to Customer
  - Clearing / settlement / interbank context  → Institution to Institution
  - Payroll / vendor / supplier outbound       → Supplier / Contractor / Employee Payments

Generic processing activities MUST inherit classification from the dominant
payment use case and cannot be assessed in isolation. Only apply contextual
inference when:
  - The process clearly sits within a defined payment use case, AND
  - There is no conflicting strong signal

Do NOT return empty mapped_categories if:
  - Direction is known, AND
  - A strong signal (instrument or L2 process context) exists
Return empty mapped_categories only when:
  - Direction is unclear, AND
  - No identifiable payment use case or instrument exists

Mapping Rules

R1 — Direct Payer-Payee Relevance
Map only where a specific payer-to-payee relationship is clearly identifiable
or unambiguously inferable from a named payment system, use case, or the
combination of customer segment and process action.
  Include: payer and payee are identifiable; relationship matches a category;
           inference where directly derived from a named payment system
           (e.g. BPAY implies Customer to Institution) or from segment +
           channel + action combined.
  Exclude: no identifiable payer-payee; cannot determine who pays whom;
           general activities with no payment relationship.

R2 — Category-Specific Payment Context
Process must operate in a specific payment scenario, not generic enterprise
activity.
  Include: description indicates who pays whom; references payment product
           processing; aligns with incomplete/inaccurate/duplicate/unauthorised
           payment or payment-compliance-failure scenarios; ITGC/governance
           processes explicitly scoped to a payment system.
  Exclude: strategy, planning, product design, marketing, pure HR (excluding
           payroll), pure onboarding, general reporting without direct
           payment execution/control context; no funds movement; upstream
           activity before any payment exists.

R3 — Payment Instruction or Execution Link
Process must create, control, act upon a payment instruction, or enable,
secure, control, or recover systems/processes that execute payments.
  Include: creates payment instruction; approves/releases payments; processes
           payments between defined parties; system/third-party/governance/
           access management directly tied to a payment system; monitoring
           directly influencing execution outcomes; disburses funds to
           employees/contractors/suppliers; activities supporting payment
           systems or lifecycle stages (e.g. IT, security, supplier,
           resilience) even where no direct payment action exists.
  Exclude: only where there is no identifiable linkage to any payment system,
           platform, or payment process.

R4 — Multi-Category Applicability
Multiple categories only where explicitly used across different payer-payee
relationships in a single activity.

Priority logic (apply in order):
  1. Prefer the category most explicitly supported by the process text.
  2. Prefer the category with the clearest payer-payee relationship in the
     l3 activity name, description, tasks, systems, third parties, channel,
     customer segment, or context.
  3. L3 activity name is a primary signal and overrides generic descriptions
     where it clearly indicates a payment action or relationship.
  4. Interbank settlement, clearing networks, counterparties, schemes, external
     banks, RBA, RITS, OFIs, or "Institutional" segment in a settlement
     context -> Institution to Institution.
  5. Customer repayments, deposits, ATM deposits, BPAY, card repayments,
     customers sending funds into the bank, or "Consumer"/"Commercial"
     segment in a repayment/deposit context -> Customer to Institution.
  6. Disbursements, refunds, interest payments, loan settlements, welfare,
     benefit/tax payments, payroll on behalf of a client institution, or
     "Consumer" segment in a disbursement context -> Institution to Customer.
  7. Transfers between customers/merchants/businesses, P2P, internal transfers,
     card payments, direct debit, customer-originated movement, or segment
     indicating two non-bank parties -> Customer to Customer.
  8. Vendor payments, internal staff payroll, supplier/contractor/employee
     payments internal to the bank -> Supplier / Contractor / Employee.
  9. If genuinely across multiple payment flows and no category dominates,
     return all defensible categories and explain why no single primary fits.
  Always set primary_category when any category is mapped.

R5 — Category-Aligned Third-Party or Scheme Participation
Category always determined by ultimate payer-payee, not intermediary processor.
  Include: systems/third parties participating in payment between defined
           payer and payee; scheme/network in execution.
  Exclude: third-party governance only (contracts, reviews) with no direct
           payment execution/control link.

R6 — Functional Naming Test
Naming supports but must never solely drive mapping. Supports conclusions
from R1-R3. L3 activity name is the primary naming signal.
  Positive signals: Process, Disburse, Settle, Pay, Transfer (with payer/payee).
  Negative signals (alone, with no other evidence): Manage, Develop, Design,
  Oversee.

Confidence for category_confidence_score:
  High   - Explicit payer-payee AND explicit payment action or system, or
            strong segment+channel+action corroboration. Category fit is
            direct and unambiguous.
  Medium - Defensible but relies on contextual inference, or multiple
            fields (including segment/channel) needed to determine category.
  Low    - Tentative, limited text, or key fields missing. Still defensible
            enough to flag as a possible mapping for review.

If payment-related (Direct) but no category is defensible:
  Set mapped_categories=[], primary_category=null, category_confidence_score=Low.
  Explain why no category can be assigned, and state the closest alternative
  that was considered in the mapping_rationale.

If payment-related (Enabling) but no direct category evidence exists:
  Use the closest alternative category. Assign with Low confidence. Cite the
  basis (segment, channel, L2 context, or operational function inference).
  primary_category must not be null for Enabling processes.

Contrastive reasoning requirement:
  For every mapping_rationale, state why the chosen category applies AND
  why the closest alternative does not. This is mandatory.

Workflow (per process):
  1. Review all supplied fields - l3_process_UUID, process_id, l2_process_name,
     l3_activity_name, description, tasks, systems, third_parties,
     governance_context, l3_activity_channels, l3_activity_customer_segments,
     and all other narrative context.
  2. Apply Step 0 scope check. Only exclude if clearly no plausible payment
     signal AND matches an excluded category exactly as defined.
  3. Test Minimum Mapping Requirement. If not satisfied, return empty mapping.
  4. Apply R1 - payer-payee relationship present (including via segment/channel)?
  5. Apply R2 - category-specific payment context, not generic enterprise?
  6. Apply R3 - directly creates, controls, or acts on payment instruction
     or payment system?
  7. Apply R5 where systems, schemes, or third parties are involved.
  8. Apply R6 as naming support only.
  9. Emit category/categories. Apply R4 if multiple are defensible.
  10. Assign category_confidence_score.
  11. Include mapping_rationale with quoted text, contrastive reasoning, and
      citation of channel/segment if used.
  12. If key fields are missing, state this explicitly and be conservative
      about confidence (not about inclusion).
  13. If no direct payment action exists, identify the payment system or
      lifecycle stage supported, and map the category based on that linkage.

──────────────────────────────────────────────────────────────────────────────
WORKED EXAMPLES  (use for interpretation guidance only - do not copy)
──────────────────────────────────────────────────────────────────────────────

Example 1 - Customer to Institution (High confidence)
  l3_activity_name: Receive Incoming Funds
  Description: "Customer sends funds into bank account"
  l3_activity_customer_segments: "BM 03.01.01 - Consumer"
  Result: mapped_categories=["Customer to Institution"], confidence=High
  Why: "Customer sends funds into bank account" explicitly identifies customer
  as payer and institution as payee, corroborated by Consumer segment. R1 and
  R2 satisfied. Customer to Customer is less appropriate because the payee is
  the bank. Institution to Customer is less appropriate because the bank is
  receiving, not paying.

Example 2 - Institution to Customer (High confidence)
  l3_activity_name: Disburse Funds
  Description: "Institution disburses funds to the customer"
  Result: mapped_categories=["Institution to Customer"], confidence=High
  Why: "Institution disburses funds to the customer" explicitly identifies the
  institution as payer and customer as payee. R1 and R3 satisfied.

Example 3 - Supplier / Contractor / Employee Payments (High confidence)
  l3_activity_name: Create Payroll Payment Batch
  Description: "Periodic staff payroll runs"
  Result: Supplier/Contractor/Employee, confidence=High
  Why: "Payroll Payment Batch" and "staff payroll runs" directly support
  employee payments. R2 and R3 satisfied. This is NOT excluded at Step 0
  because payroll/staff payment is an explicit valid payment category.

Example 4 - Institution to Institution (High confidence)
  l3_activity_name: Settle Interbank Obligation
  Description: "Settlement between banks through clearing network"
  l3_activity_customer_segments: "BM 03.03.01 - Institutional"
  Result: Institution to Institution, confidence=High
  Why: "Interbank Obligation" and "clearing network", corroborated by
  Institutional segment. R1, R3, and R5 satisfied.

Example 5 - Multiple categories (Medium confidence)
  l3_activity_name: Generate Payment Advice
  Description: "Advice applies to both customers and counterparties"
  Result: ["Institution to Customer", "Institution to Institution"],
  primary=Institution to Customer, confidence=Medium
  Why: Text explicitly supports both audiences. Multiple categories under R4.

Example 6 - ITGC process scoped to payments (now correctly included)
  l3_activity_name: Manage Privileged Access to SWIFT Interface
  Description: "Review and approve privileged access requests to SWIFT
  messaging infrastructure used for international payment settlement"
  Result: is_payment_process=true, payment_process_confidence=High,
  mapped_categories=["Institution to Institution"], confidence=Medium
  Why: Explicitly scoped to SWIFT (named payment system) and tied to
  international payment settlement. Step 0 does not exclude this - it is
  access governance directly tied to a payment system. Category mapped to
  Institution to Institution because SWIFT settlement context implies
  interbank obligation, though specific payer-payee is inferred from
  context rather than explicitly stated (Medium confidence).

Example 7 - Payment-related, no defensible category (Low confidence)
  l3_activity_name: Monitor Payment Exception Queue
  Description: "Team monitors exception queue for flagged payment transactions"
  Result: is_payment_process=true, mapped_categories=[], primary_category=null,
  category_confidence_score=Low
  Why: Clearly payment-related. However, exception queue covers all payment
  types - no payer-payee evidence, including segment/channel, allows a
  specific category to be assigned.

Example 8 - Genuinely not payment-related (Step 0 exclusion)
  l3_activity_name: Design New Term Deposit Product Features
  Description: "Product team designs new term deposit features based on
  market research and competitor analysis"
  Result: is_payment_process=false, exclusion_gate_applied=true
  Why: Pure product design with no payment processing, payment system, or
  payment governance reference. No plausible payment signal present.

──────────────────────────────────────────────────────────────────────────────
IDENTIFIER RULE
──────────────────────────────────────────────────────────────────────────────

Each process includes l3_process_UUID. This is the unique join key.
Copy it exactly into every output mapping. Do not alter or truncate.
Do not use l3_activity_id or process_id as the primary key - the same
l3_activity_id may appear in multiple process contexts with different UUIDs.

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
      "is_payment_process": true | false,
      "payment_process_type": "Direct | Enabling | Non-payment",
      "payment_process_confidence": "High | Medium | Low",
      "payment_process_rationale": "1-2 sentences with at least one quoted phrase",
      "exclusion_gate_applied": true | false,
      "exclusion_reason": "explanation if gate applied, else null",
      "mapped_categories": ["valid category strings only - or empty array"],
      "primary_category": "single valid category string — MUST NOT be null if is_payment_process=true",
      "category_confidence_score": "High | Medium | Low",
      "rule_hits": ["R1", "R2", ...],
      "channel_segment_used": true | false,
      "mapping_rationale": "quoted text; why chosen category applies; why closest alternative does not; cite channel/segment if used"
    }
  ],
  "notes": ["batch-level observations"]
}

──────────────────────────────────────────────────────────────────────────────
HARD RULES
──────────────────────────────────────────────────────────────────────────────

- Emit exactly one mapping entry per input process.
- Use only l3_process_UUID values present in the input batch.
- Use only the five valid payment category strings exactly as listed.
- payment_process_type must be exactly one of: "Direct", "Enabling", "Non-payment".
- Every rationale must include at least one quoted phrase from the process payload.
- Payroll, staff payment, vendor payment, and contractor payment processes are
  NEVER excluded — they map to Supplier / Contractor / Employee Payments.
- IT security, ITGC, BCP, change management, access management, governance,
  and resilience processes are NEVER classified as Non-payment at Step 0 —
  they are Enabling at minimum.
- Do NOT stop evaluation because a process lacks explicit payment keywords.
- If is_payment_process=false: mapped_categories=[], primary_category=null.
- If is_payment_process=true (Direct or Enabling): primary_category MUST
  be assigned. Empty mapped_categories is NOT permitted for payment-related
  processes. Use closest alternative with Low confidence if needed.
- Do not return empty categories if a defensible closest alternative exists.
- primary_category must not be null for any payment-related process.
- If multiple categories are plausible, always assign exactly one primary_category.
- Multiple categories require distinct payer-payee evidence for each.
- If primary_category is set, it must appear in mapped_categories.
- Channel and customer segment fields must corroborate, not solely drive,
  a mapping — always combine with process description or L2 context.
- Do not invent facts. All inference must be grounded in supplied fields.
"""


# ──────────────────────────────────────────────────────────────────────────────
#  AUGMENTATION FUNCTIONS  (post-processing — no additional LLM calls)
# ──────────────────────────────────────────────────────────────────────────────

def derive_objective_confidence(rule_hits: list) -> str:
    """
    Confidence derived from which rules fired — independent of model self-assessment.
    R1, R2, R3 are the core evidence rules. R4, R5, R6 are structural/naming.
    Two or more core rules = High. One = Medium. None = Low.
    """
    strong = {"R1", "R2", "R3"}
    count  = len(set(rule_hits or []) & strong)
    if count >= 2:
        return "High"
    elif count == 1:
        return "Medium"
    return "Low"


def verify_rationale_citation(rationale: str, process_payload: dict) -> bool:
    """
    Verify that at least one quoted phrase in the rationale actually exists
    in the source process payload. Returns False if rationale has no quotes
    or if none of the quoted phrases match.
    Quoted phrases shorter than 4 characters are ignored (trivial matches).
    """
    if not rationale:
        return False

    # Flatten all string values from the payload (recursive)
    source_parts = []

    def extract(obj):
        if isinstance(obj, str):
            source_parts.append(obj)
        elif isinstance(obj, list):
            for item in obj:
                extract(item)
        elif isinstance(obj, dict):
            for v in obj.values():
                extract(v)

    extract(process_payload)
    source_text = " ".join(source_parts).lower()

    quoted = re.findall(r'"([^"]{4,})"', rationale)
    if not quoted:
        return False
    return any(phrase.lower() in source_text for phrase in quoted)


def calculate_field_coverage(process_payload: dict, fields: list) -> dict:
    """
    Score how many of the relevant source fields are non-empty.
    Low coverage + High stated confidence is a flag worth reviewing.
    """
    present = sum(
        1 for f in fields
        if process_payload.get(f) and str(process_payload[f]).strip()
        and str(process_payload[f]).strip().lower() not in ("none", "null", "nan")
    )
    total = len(fields)
    pct   = round(present / total * 100) if total > 0 else 0
    return {
        "field_coverage_pct":  pct,
        "field_coverage_tier": "High" if pct >= 70 else "Medium" if pct >= 40 else "Low",
    }


def check_batch_distribution(mappings: list) -> list:
    """
    Warn if one category dominates the batch — potential model drift.
    Threshold: >70% same primary category in a single batch.
    """
    warnings = []
    cats = [m.get("primary_category") for m in mappings if m.get("primary_category")]
    if len(cats) < 2:
        return warnings
    for cat, count in Counter(cats).items():
        pct = count / len(mappings) * 100
        if pct > 70:
            warnings.append(
                f"DISTRIBUTION WARNING: '{cat}' = {pct:.0f}% of batch "
                f"({count}/{len(mappings)}). Review for model drift."
            )
    return warnings


def verify_channel_segment_claim(mapping: dict, process_payload: dict) -> dict:
    """
    Guard against hallucination risk from "possible processes" wording.
    If the model claims channel_segment_used=true but l3_activity_channels
    and l3_activity_customer_segments are both empty/absent in the source
    payload, the model fabricated evidence that does not exist. Flag this
    explicitly — it is a stronger signal than generic citation_verified
    because it targets the exact new fields most likely to be hallucinated
    against now that the model has been told to consider channel/segment.
    """
    claimed_used = mapping.get("channel_segment_used", False)
    if not claimed_used:
        return {"channel_segment_claim_verified": None}  # not applicable

    channel = str(process_payload.get("l3_activity_channels", "")).strip()
    segment = str(process_payload.get("l3_activity_customer_segments", "")).strip()
    has_real_evidence = bool(channel) or bool(segment)

    return {"channel_segment_claim_verified": has_real_evidence}


def augment_mapping(mapping: dict, process_payload: dict) -> dict:
    """
    Add objective confidence, citation verification, field coverage,
    confidence conflict detection, channel/segment claim verification,
    and SME review flag to a mapping dict.
    """
    m = dict(mapping)

    # 1. Objective confidence from rule hits
    m["objective_confidence"] = derive_objective_confidence(m.get("rule_hits", []))

    # 2. Rationale citation verification
    #    Check both rationale fields — use whichever is populated
    rationale = m.get("mapping_rationale") or m.get("payment_process_rationale") or ""
    m["citation_verified"] = verify_rationale_citation(rationale, process_payload)

    # 3. Confidence conflict — stated High but objective Low
    stated    = m.get("category_confidence_score", "Low")
    objective = m["objective_confidence"]
    m["confidence_conflict"] = (stated == "High" and objective == "Low")

    # 4. Field coverage (now includes channel + customer segment)
    coverage = calculate_field_coverage(process_payload, RELEVANT_SOURCE_FIELDS)
    m.update(coverage)

    # 5. Channel/segment claim verification — hallucination guard for new fields
    claim_check = verify_channel_segment_claim(m, process_payload)
    m.update(claim_check)
    channel_segment_hallucinated = (claim_check["channel_segment_claim_verified"] is False)

    # 6. SME review flag — any of the following triggers review
    m["sme_review_flag"] = (
        not m["citation_verified"]
        or m["confidence_conflict"]
        or channel_segment_hallucinated
        or coverage["field_coverage_tier"] == "Low"
        or m.get("category_confidence_score") == "Low"
        or m.get("payment_process_confidence") == "Low"
    )

    return m


# ──────────────────────────────────────────────────────────────────────────────
#  CORE UTILITIES
# ──────────────────────────────────────────────────────────────────────────────

def preflight_check():
    print("\n  Pre-flight environment check")
    print("  " + "─" * 60)
    for var in ("REQUESTS_CA_BUNDLE", "SSL_CERT_FILE", "HTTP_PROXY",
                "HTTPS_PROXY", "AZURE_CA_BUNDLE"):
        print(f"  {var:<22} = {os.getenv(var) or '(not set)'}")
    if os.getenv("AZURE_CA_BUNDLE"):
        print("  WARNING: AZURE_CA_BUNDLE is set — recommend removing for Azure CLI auth.")


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


def load_processes(input_file: str) -> list:
    with open(input_file, "r", encoding="utf-8") as f:
        data = json.load(f)
    if isinstance(data, dict) and "processes" in data:
        return data["processes"]
    if isinstance(data, list):
        return data
    raise ValueError("Input must be a JSON list or an object with a 'processes' key.")


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


def validate_output(parsed: dict, batch: list) -> list:
    """Structural validation — UUID completeness, valid categories, consistency."""
    warnings = []
    input_ids  = {str(p.get("l3_process_UUID")) for p in batch}
    mappings   = parsed.get("mappings", [])
    output_ids = {str(m.get("l3_process_UUID")) for m in mappings}

    missing = sorted(input_ids - output_ids)
    extra   = sorted(output_ids - input_ids)
    if missing:
        warnings.append(f"Missing UUIDs in output: {missing}")
    if extra:
        warnings.append(f"Unknown UUIDs returned: {extra}")
    if len(mappings) != len(batch):
        warnings.append(
            f"Expected {len(batch)} mappings, received {len(mappings)}"
        )

    for m in mappings:
        uid = m.get("l3_process_UUID")
        ptype = m.get("payment_process_type")
        if ptype and ptype not in VALID_PROCESS_TYPES:
            warnings.append(f"Invalid payment_process_type {ptype!r} for UUID {uid}")
        for cat in (m.get("mapped_categories") or []):
            if cat not in VALID_CATEGORIES:
                warnings.append(f"Invalid category {cat!r} for UUID {uid}")
        primary = m.get("primary_category")
        if primary and primary not in VALID_CATEGORIES:
            warnings.append(f"Invalid primary_category {primary!r} for UUID {uid}")
        # For Enabling/Direct processes, primary_category must be assigned
        if m.get("is_payment_process") and not primary:
            warnings.append(
                f"is_payment_process=true but primary_category is null for UUID {uid} "
                f"(type={ptype}). Check closest alternative."
            )
        if not m.get("is_payment_process") and (m.get("mapped_categories") or []):
            warnings.append(
                f"is_payment_process=false but categories returned for UUID {uid}"
            )
        cats = m.get("mapped_categories") or []
        if primary and cats and primary not in cats:
            warnings.append(
                f"primary_category '{primary}' not in mapped_categories for UUID {uid}"
            )

    return warnings


# ──────────────────────────────────────────────────────────────────────────────
#  LLM CALL
# ──────────────────────────────────────────────────────────────────────────────

def call_llm(client: AzureOpenAI, batch: list, batch_no: int) -> tuple:
    """
    Call the LLM for one batch. Returns (parsed_dict, meta_dict).
    Uses exponential backoff with jitter for retries.
    """
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
    # reasoning_effort is supported for GPT-5.x reasoning models
    # temperature is NOT supported — intentionally omitted
    if MODEL.startswith("gpt-5"):
        kwargs["reasoning_effort"] = REASONING_EFFORT

    for attempt in range(1, RETRY_COUNT + 1):
        try:
            t0       = time.time()
            response = client.chat.completions.create(**kwargs)
            latency  = int((time.time() - t0) * 1000)

            text   = response.choices[0].message.content or ""
            parsed = parse_json_response(text)

            usage_obj = getattr(response, "usage", None)
            usage = {}
            if usage_obj:
                usage = {
                    "input_tokens":  getattr(usage_obj, "prompt_tokens",     None),
                    "output_tokens": getattr(usage_obj, "completion_tokens", None),
                    "total_tokens":  getattr(usage_obj, "total_tokens",      None),
                }
                details = getattr(usage_obj, "completion_tokens_details", None)
                if details:
                    usage["reasoning_tokens"] = getattr(details, "reasoning_tokens", None)

            return parsed, {
                "batch_no":     batch_no,
                "latency_ms":   latency,
                "usage":        usage,
                "raw_response": text,
            }

        except Exception as exc:
            if attempt < RETRY_COUNT:
                sleep_time = min(
                    RETRY_BASE_SLEEP * (2 ** (attempt - 1)),
                    RETRY_MAX_SLEEP
                )
                jitter     = random.uniform(0, 2)
                total_sleep = sleep_time + jitter
                print(
                    f"    Attempt {attempt} failed: {exc}. "
                    f"Retrying in {total_sleep:.1f}s…"
                )
                time.sleep(total_sleep)
            else:
                raise


# ──────────────────────────────────────────────────────────────────────────────
#  OUTPUT
# ──────────────────────────────────────────────────────────────────────────────

def flatten_results(batch_results: list, process_lookup_by_batch: dict) -> pd.DataFrame:
    rows = []
    for result in batch_results:
        usage = result.get("usage") or {}
        batch_no = result.get("batch_no")
        lookup = process_lookup_by_batch.get(batch_no, {})

        for m in result.get("mappings", []):
            uuid = str(m.get("l3_process_UUID", ""))
            source = lookup.get(uuid, {})

            # Pull review fields directly from the SOURCE payload, not the model's
            # echoed output — guarantees accuracy and is independent of any
            # model error in copying these fields back.
            rows.append({
                # Identifiers
                "batch_no":                  batch_no,
                "l3_process_UUID":           m.get("l3_process_UUID"),
                "process_id":                m.get("process_id"),
                # Review context — sourced directly from input payload
                "l2_process_name":           source.get("l2_process_name", m.get("l2_process_name")),
                "l2_process_description":    source.get("l2_process_description", ""),
                "l3_activity_name":          source.get("l3_activity_name", m.get("l3_activity_name")),
                "l3_activity_description":   source.get("l3_activity_description", source.get("description", "")),
                # Step 0 — Exclusion gate
                "exclusion_gate_applied":    m.get("exclusion_gate_applied"),
                "exclusion_reason":          m.get("exclusion_reason"),
                # Step 1 — Three-tier classification
                "is_payment_process":        m.get("is_payment_process"),
                "payment_process_type":      m.get("payment_process_type"),
                "payment_process_confidence":m.get("payment_process_confidence"),
                "payment_process_rationale": m.get("payment_process_rationale"),
                # Step 2 — Category mapping
                "mapped_categories":         " | ".join(m.get("mapped_categories", []) or []),
                "primary_category":          m.get("primary_category"),
                "category_confidence_score": m.get("category_confidence_score"),
                "rule_hits":                 " | ".join(m.get("rule_hits", []) or []),
                "mapping_rationale":         m.get("mapping_rationale"),
                # Augmentation — objective cross-checks
                "objective_confidence":      m.get("objective_confidence"),
                "citation_verified":         m.get("citation_verified"),
                "confidence_conflict":       m.get("confidence_conflict"),
                "channel_segment_used":      m.get("channel_segment_used"),
                "channel_segment_claim_verified": m.get("channel_segment_claim_verified"),
                "field_coverage_pct":        m.get("field_coverage_pct"),
                "field_coverage_tier":       m.get("field_coverage_tier"),
                "sme_review_flag":           m.get("sme_review_flag"),
                # Token usage
                "input_tokens":              usage.get("input_tokens"),
                "output_tokens":             usage.get("output_tokens"),
                "total_tokens":              usage.get("total_tokens"),
                "reasoning_tokens":          usage.get("reasoning_tokens"),
            })
    return pd.DataFrame(rows)


def build_summary(df: pd.DataFrame) -> pd.DataFrame:
    """High-level run summary for the summary sheet."""
    total = len(df)
    if total == 0:
        return pd.DataFrame()
    payment = df["is_payment_process"].sum() if "is_payment_process" in df else 0
    excluded = df["exclusion_gate_applied"].sum() if "exclusion_gate_applied" in df else 0
    sme_flags = df["sme_review_flag"].sum() if "sme_review_flag" in df else 0
    no_citation = (~df["citation_verified"]).sum() if "citation_verified" in df else 0
    conf_conflict = df["confidence_conflict"].sum() if "confidence_conflict" in df else 0
    channel_seg_hallucinated = (
        (df["channel_segment_claim_verified"] == False).sum()
        if "channel_segment_claim_verified" in df else 0
    )
    channel_seg_used_count = (
        (df["channel_segment_used"] == True).sum()
        if "channel_segment_used" in df else 0
    )

    cat_dist = {}
    if "primary_category" in df:
        cat_dist = df[df["primary_category"].notna()]["primary_category"].value_counts().to_dict()

    # Process type breakdown
    direct_count  = (df["payment_process_type"] == "Direct").sum()  if "payment_process_type" in df else 0
    enabling_count= (df["payment_process_type"] == "Enabling").sum() if "payment_process_type" in df else 0
    null_primary  = (df["is_payment_process"] == True).sum() - (df["primary_category"].notna() & (df["is_payment_process"] == True)).sum() if "primary_category" in df else 0

    rows = [
        ("Total processes",             total),
        ("Excluded (gate applied)",      int(excluded)),
        ("Payment-related (all)",        int(payment)),
        ("  ↳ Direct payment",           int(direct_count)),
        ("  ↳ Enabling (ITGC/BCP/Gov)",  int(enabling_count)),
        ("Not payment-related",          total - int(payment)),
        ("SME review flagged",           int(sme_flags)),
        ("Citation unverified",          int(no_citation)),
        ("Confidence conflict (Hi↔Lo)",  int(conf_conflict)),
        ("Channel/segment used",         int(channel_seg_used_count)),
        ("Channel/segment hallucinated", int(channel_seg_hallucinated)),
        ("Missing primary_category",     int(null_primary)),
        ("──────────────────────────", "──────────────"),
    ] + [(f"Category: {k}", v) for k, v in cat_dist.items()]

    return pd.DataFrame(rows, columns=["Metric", "Count"])


def write_outputs(input_file: str, batch_results: list,
                  raw_records: list, all_warnings: list,
                  process_lookup_by_batch: dict) -> None:
    input_path  = Path(input_file)
    output_dir  = input_path.parent
    stem        = input_path.stem.replace("_aggregated_for_llm_uuid", "")
    suffix      = f"_llm_payment_relevance_category_uuid_{REASONING_EFFORT}"

    json_path   = output_dir / f"{stem}{suffix}.json"
    jsonl_path  = output_dir / f"{stem}{suffix}_raw.jsonl"
    xlsx_path   = output_dir / f"{stem}{suffix}.xlsx"

    combined = {
        "run_ts": datetime.now().isoformat(timespec="seconds"),
        "config": {
            "endpoint":              ENDPOINT,
            "api_version":           API_VERSION,
            "model":                 MODEL,
            "reasoning_effort":      REASONING_EFFORT,
            "batch_size":            BATCH_SIZE,
            "max_completion_tokens": MAX_COMPLETION_TOKENS,
        },
        "warnings":        all_warnings,
        "mappings":        [],
        "batch_metadata":  [],
    }
    for result in batch_results:
        combined["mappings"].extend(result.get("mappings", []))
        combined["batch_metadata"].append({
            "batch_no":   result.get("batch_no"),
            "latency_ms": result.get("latency_ms"),
            "usage":      result.get("usage"),
            "warnings":   result.get("warnings", []),
            "notes":      result.get("notes", []),
        })

    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(combined, f, ensure_ascii=False, indent=2)

    with open(jsonl_path, "w", encoding="utf-8") as f:
        for record in raw_records:
            f.write(json.dumps(record, ensure_ascii=False) + "\n")

    df_all = flatten_results(batch_results, process_lookup_by_batch)

    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        df_all.to_excel(writer, index=False, sheet_name="mappings")

        # SME review queue — all flagged rows in one place
        if "sme_review_flag" in df_all.columns:
            df_sme = df_all[df_all["sme_review_flag"] == True].copy()
            df_sme.to_excel(writer, index=False, sheet_name="sme_review_queue")

        # Run summary
        build_summary(df_all).to_excel(writer, index=False, sheet_name="summary")

        # Batch metadata
        pd.DataFrame(combined["batch_metadata"]).to_excel(
            writer, index=False, sheet_name="batch_metadata"
        )

        # Warnings
        pd.DataFrame({"warning": all_warnings}).to_excel(
            writer, index=False, sheet_name="warnings"
        )

    print("\n  Outputs saved:")
    print(f"  JSON   → {json_path}")
    print(f"  JSONL  → {jsonl_path}")
    print(f"  Excel  → {xlsx_path}")
    if "sme_review_flag" in df_all.columns:
        sme_count = (df_all["sme_review_flag"] == True).sum()
        print(f"\n  SME review queue: {sme_count} / {len(df_all)} processes flagged")


# ──────────────────────────────────────────────────────────────────────────────
#  MAIN
# ──────────────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Payment relevance + category mapping using l3_process_UUID."
    )
    parser.add_argument("input_file", nargs="?", default=DEFAULT_INPUT_FILE)
    parser.add_argument("--batch-size", type=int, default=BATCH_SIZE)
    args = parser.parse_args()

    if not os.path.exists(args.input_file):
        raise FileNotFoundError(f"Input file not found: {args.input_file}")

    print("=" * 72)
    print("  Payment Relevance + Category LLM Runner — l3_process_UUID")
    print(f"  Model                 : {MODEL}")
    print(f"  Reasoning effort      : {REASONING_EFFORT}")
    print(f"  Max completion tokens : {MAX_COMPLETION_TOKENS}")
    print(f"  Batch size            : {args.batch_size}")
    print(f"  Input file            : {args.input_file}")
    print("=" * 72)

    processes = load_processes(args.input_file)

    missing_uuid = [
        p.get("process_id") for p in processes if not p.get("l3_process_UUID")
    ]
    if missing_uuid:
        raise ValueError(
            f"Processes missing l3_process_UUID: {missing_uuid}. "
            "Re-run the aggregation script to add UUIDs."
        )

    print(f"\n  Loaded {len(processes):,} process payloads")
    preflight_check()

    print("\n  Initialising Azure OpenAI client…")
    client = init_client()
    print("  Client ready.")

    batch_results = []
    raw_records   = []
    all_warnings  = []
    process_lookup_by_batch = {}  # batch_no -> {uuid: source_payload} for output review fields
    total_batches = (len(processes) + args.batch_size - 1) // args.batch_size

    for batch_no, batch in chunk_list(processes, args.batch_size):
        print(f"\n  Batch {batch_no}/{total_batches}: {len(batch)} process(es)")

        parsed, meta = call_llm(client, batch, batch_no)

        # Structural validation
        struct_warnings = validate_output(parsed, batch)

        # Post-processing augmentation
        process_lookup = {str(p.get("l3_process_UUID")): p for p in batch}
        process_lookup_by_batch[batch_no] = process_lookup  # retained for output review fields
        augmented_mappings = []
        for m in parsed.get("mappings", []):
            uuid    = str(m.get("l3_process_UUID", ""))
            payload = process_lookup.get(uuid, {})
            augmented_mappings.append(augment_mapping(m, payload))
        parsed["mappings"] = augmented_mappings

        # Batch distribution check
        dist_warnings = check_batch_distribution(augmented_mappings)

        batch_warnings = struct_warnings + dist_warnings
        result = {
            "batch_no":   batch_no,
            "mappings":   augmented_mappings,
            "notes":      parsed.get("notes", []),
            "warnings":   batch_warnings,
            "latency_ms": meta.get("latency_ms"),
            "usage":      meta.get("usage"),
        }
        batch_results.append(result)
        raw_records.append({
            "batch_no":            batch_no,
            "input_l3_process_UUIDs": [p.get("l3_process_UUID") for p in batch],
            "parsed_response":     parsed,
            "raw_response":        meta.get("raw_response"),
            "usage":               meta.get("usage"),
            "latency_ms":          meta.get("latency_ms"),
            "warnings":            batch_warnings,
        })
        all_warnings.extend([f"Batch {batch_no}: {w}" for w in batch_warnings])

        sme_count = sum(1 for m in augmented_mappings if m.get("sme_review_flag"))
        print(f"  Returned  : {len(augmented_mappings)} mappings")
        print(f"  Latency   : {result['latency_ms']}ms")
        usage = result["usage"] or {}
        rt = f" | {usage.get('reasoning_tokens')} thinking" if usage.get("reasoning_tokens") else ""
        print(f"  Tokens    : {usage.get('input_tokens')} in | {usage.get('output_tokens')} out{rt}")
        print(f"  SME flags : {sme_count}/{len(augmented_mappings)}")
        if batch_warnings:
            for w in batch_warnings:
                print(f"  ⚠  {w}")

    write_outputs(args.input_file, batch_results, raw_records, all_warnings, process_lookup_by_batch)
    print("\n  Done.")


if __name__ == "__main__":
    main()
