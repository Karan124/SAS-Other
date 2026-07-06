r"""
run_payment_relevance_and_category_llm_uuid.py
───────────────────────────────────────────────────────────────────────────────
Payments Controls PoC — Process Payment Relevance & Category Classifier
Full production run: 3,885 processes. Uses l3_process_UUID as the join key.

Classification:
  Direct      — directly executes payment transactions → gets category mapping
  Enabling    — supports payment infrastructure (ITGC, BCP, governance, change)
                → payment-related but NO category assigned (always null)
  Non-payment — no payment linkage → excluded

Key features:
  - Checkpoint / resume : safe to interrupt and restart mid-run
  - Incremental JSONL   : each batch saved to disk immediately
  - Progress display    : %, ETA, running cost in AUD
  - --force flag        : ignore checkpoint and reprocess from scratch
  - reasoning_effort    : "medium" — required for nuanced classification quality
  - Windows-safe JSON   : ensure_ascii=True avoids OSError 22 on Windows
  - Validation          : flags Enabling processes that incorrectly receive
                          a category as PROMPT VIOLATION warnings

Before running (PowerShell):
  az account set --subscription 6c72e6c5-ed48-4030-b29c-34e2849c9288
  $env:REQUESTS_CA_BUNDLE = "C:\Users\m061400\ai-test\cacert.pem"
  $env:SSL_CERT_FILE      = "C:\Users\m061400\ai-test\cacert.pem"
  Remove-Item Env:AZURE_CA_BUNDLE -ErrorAction SilentlyContinue
  python run_payment_relevance_and_category_llm_uuid.py           (fresh or resume)
  python run_payment_relevance_and_category_llm_uuid.py --force   (ignore checkpoint)
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
    r"\payment_process_candidates_aggregated_for_llm_uuid.json"
)

BATCH_SIZE             = 3        # was 5 — smaller = more reasoning per process
MAX_COMPLETION_TOKENS  = 16000    # was 12000 — reasoning tokens need headroom
REASONING_EFFORT       = "medium" # structured classification with nuanced cases

RETRY_COUNT      = 3
RETRY_BASE_SLEEP = 5    # seconds — base for exponential backoff
RETRY_MAX_SLEEP  = 60   # seconds — cap
INTER_BATCH_SLEEP = 0.5  # seconds between batches — light throttle for long runs

# Cost tracking (for running AUD estimate printed per batch)
INPUT_PRICE_USD_PER_M  = 1.75
OUTPUT_PRICE_USD_PER_M = 14.00
AUD_USD_RATE           = 0.65

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
    "l3_activity_channels",            # where the process originates/takes place
    "l3_activity_customer_segments",   # parties involved in the process
    "l3_activity_product_service",     # product/service context (loan, FX, trade finance etc.)
]

# Fields specifically used as payer-payee corroborating evidence.
# Used to populate channel_segment_used cross-check (augmentation).
PAYER_PAYEE_EVIDENCE_FIELDS = [
    "l3_activity_channels",
    "l3_activity_customer_segments",
    "l3_activity_product_service",
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
    "Direct processes receive payment category mapping. "
    "Enabling processes are payment-related but receive NO category — "
    "mapped_categories=[] and primary_category=null always.\n"
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

For each process in the batch, classify it as Direct, Enabling, or
Non-payment, then follow the routing rule below exactly.

CRITICAL ROUTING RULE — read this before anything else:

  Direct      → is_payment_process=true  → proceed to Step 2 (category mapping)
  Enabling    → is_payment_process=true  → STOP after Step 1. No category.
                mapped_categories=[] and primary_category=null. Always.
  Non-payment → is_payment_process=false → STOP after Step 0. No category.

Enabling processes never receive a payment category — not under any
circumstance, not even where strong payer-payee evidence or named payment
systems exist in the description. Category mapping is for Direct only.

Do not invent facts. All reasoning must be grounded in supplied fields only.

──────────────────────────────────────────────────────────────────────────────
STEP 0 — NON-PAYMENT EXCLUSION  (narrow gate — clear evidence required)
──────────────────────────────────────────────────────────────────────────────

Mark is_payment_process=false, payment_process_type="Non-payment", and
exclusion_gate_applied=true ONLY where there is CLEAR EVIDENCE of no
linkage to payments AND the process is unambiguously one of:

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

CORE RULE 1 — PAYMENT LIFECYCLE VS PRODUCT / ACCOUNT / FACILITY / CONTRACT LIFECYCLE

An activity is payment-related ONLY if it directly participates in the
payment lifecycle. Payment lifecycle activities include:
  receiving, validating, authorising, executing, routing, clearing,
  settling, posting, reconciling, monitoring, notifying, investigating,
  recalling, recovering payments, or controlling/governing payment execution.

The following activity types are NOT payment-related unless the supplied
evidence clearly shows direct payment execution, payment instruction
handling, payment settlement, payment reconciliation, payment exception
management or payment control:
  - Account creation, account setup, account activation
  - Customer onboarding, customer verification
  - Product implementation, product configuration
  - Facility setup, facility documentation
  - Contract preparation, terms and conditions preparation
  - Financial table preparation, customer communications
  - Account, facility, contract or customer lifecycle management

Do NOT classify an activity as payment-related merely because it is
associated with a payment-enabled product, account, card, loan, deposit,
facility or customer.

Do NOT use this gate for: IT security, ITGC, change management, access
management, monitoring, governance, business continuity, resilience, or
supplier oversight processes. These are Enabling by nature at a bank.
Do not exclude solely because a process lacks an explicit payment keyword.

──────────────────────────────────────────────────────────────────────────────
STEP 1 — THREE-TIER CLASSIFICATION
──────────────────────────────────────────────────────────────────────────────

Classify each process. Set payment_process_type to exactly one of:
"Direct", "Enabling", or "Non-payment".

DIRECT — is_payment_process=true, payment_process_type="Direct"
  Process directly executes, processes, settles, reconciles, posts, routes,
  authorises, amends, cancels, or reports on payment transactions.
  Evidence: explicit payment action verbs, named payment systems, payer-payee
  relationships, payment instruments (Direct Debit, BPAY, RTGS, SWIFT, etc.).

ENABLING — is_payment_process=true, payment_process_type="Enabling"
  Process supports the infrastructure, security, governance, continuity, or
  third-party relationships that enable payment operations to function.
  Classify as Enabling — even with no named payment system — for:
    - IT security / ITGC: vulnerability management, pen testing, access
      management, patch management, secure design
    - Business continuity / resilience / disaster recovery
    - Change management for technology or operations
    - Third-party and supplier governance for technology/operational providers
    - Incident and crisis management
    - Risk oversight and assurance for operational or technology risk
  AFTER CLASSIFYING AS ENABLING: stop. Set mapped_categories=[] and
  primary_category=null. Do not read Step 2. Move to output.

NON-PAYMENT — is_payment_process=false, payment_process_type="Non-payment"
  Only where there is clear evidence of no linkage to payments.
  See Step 0. Do not use for IT, BCP, governance, or change management.

Confidence for payment_process_confidence:
  High   — Named payment system AND explicit payment action or disbursement.
  Medium — Payment relevance strongly implied by context, segment, channel,
            or L2 process name, but no explicit keyword.
  Low    — Inferred from process type (e.g. IT security at a bank = Enabling).
            Acceptable — flagged for SME review.

──────────────────────────────────────────────────────────────────────────────
STEP 1 OUTPUT FOR ENABLING AND NON-PAYMENT PROCESSES
──────────────────────────────────────────────────────────────────────────────

If payment_process_type = "Enabling" or "Non-payment":
  mapped_categories          = []
  primary_category           = null
  category_confidence_score  = null
  rule_hits                  = []
  channel_segment_used       = false
  mapping_rationale          = null
Do not attempt to assign a category. Do not read Step 2.

──────────────────────────────────────────────────────────────────────────────
STEP 2 — PAYMENT CATEGORY MAPPING  (DIRECT PROCESSES ONLY)
──────────────────────────────────────────────────────────────────────────────

THIS STEP APPLIES TO DIRECT PROCESSES ONLY.
If payment_process_type is "Enabling" or "Non-payment" — stop. Return to
Step 1 output. Do not read any further.

MANDATORY FOR DIRECT PROCESSES: Always attempt to assign primary_category.
Where any defensible payer-payee signal exists — even weak — assign the
closest alternative category with category_confidence_score=Low and explain.
Only return null primary_category for a Direct process when Core Rule 5
applies: no payer-payee evidence whatsoever remains after exhausting all
signals (instrument, direction, product/service, segment, channel, L2 context).
In that case state the closest alternative considered in mapping_rationale.

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
  Visa/Mastercard settlement, POS/merchant settlements, RITS/RTGS.

  Supplier / Contractor / Employee Payments
  The bank pays its own suppliers, contractors, or employees.
  Signals: Vendor payments, internal staff payroll, contractor invoices,
  supplier expense reimbursements.

Note: Signals are illustrative. They do not override process evidence.

──────────────────────────────────────────────────────────────────────────────
CORE RULES FOR CATEGORY MAPPING (Direct processes only)
──────────────────────────────────────────────────────────────────────────────

CORE RULE 2 — ULTIMATE PAYER-PAYEE PRINCIPLE
Always determine category from the ultimate payer and ultimate payee.
Payment rails, channels and settlement mechanisms are execution context.
They are NOT payment categories. Do NOT determine category from:
  SWIFT, RTGS, RITS, NPP, BPAY, Direct Entry, Visa, Mastercard, PEXA,
  OFI, clearing, settlement, inbound, outbound, lifecycle stage, or
  activity title alone.
Determine category from: payment purpose, product/service, borrower,
beneficiary, merchant, biller, customer, financial institution, settlement
counterparty, account owner, scheme participant or clearing participant.

CORE RULE 3 — INBOUND / OUTBOUND WORDING
Inbound and outbound indicate payment direction only. They do not by
themselves identify the payer-payee category. Always combine direction
with payment purpose, product context, or counterparty evidence.

CORE RULE 4 — CUSTOMER REPAYMENT RULE
Where a customer initiates, authorises or submits a payment to repay a
loan, mortgage, credit card, overdraft, lease or other obligation owed to
a financial institution, classify as Customer to Institution — even if
interbank clearing, RTGS, SWIFT, BPAY, Direct Entry, OFI processing or
settlement occurs behind the scenes.

──────────────────────────────────────────────────────────────────────────────
USING CHANNEL, CUSTOMER SEGMENT AND PRODUCT/SERVICE AS PAYER-PAYEE EVIDENCE (Direct only)
──────────────────────────────────────────────────────────────────────────────

Each process may include three supporting context fields:

  l3_activity_channels          — where the process takes place or originates
  l3_activity_customer_segments — the parties/customer types involved
  l3_activity_product_service   — the product or service associated with the
                                   activity (e.g. Home Loans, Transaction
                                   Accounts, Trade Finance, FX Derivatives,
                                   Merchant Acquiring, Bonds/REPOs)

PRODUCT / SERVICE CONTEXT RULE (strong supporting signal):
  - Home Loans, Loan Products, Debt Products, Asset Finance, Credit Products
    → repayment context implies Customer to Institution;
      disbursement/drawdown context implies Institution to Customer
  - Transaction Accounts → customer-originated payments, deposits or transfers
  - Merchant Acquiring → Customer to Customer (cardholder pays merchant)
  - Trade Finance → Customer to Customer, Customer to Institution or
    Institution to Institution depending on payment purpose and counterparties
  - Financial Markets (Bonds/REPOs, FX Derivatives, IR Derivatives,
    Structured Notes, Loans and Deposits) → Institution to Institution where
    evidence shows interbank/counterparty settlement
  Product/service context must not override explicit payer-payee evidence.

CHANNEL AND SEGMENT INTERPRETATION RULE:
  - ATM, Branch, Digital Online/Mobile → often customer-originated; category
    still depends on who pays whom
  - Corporate Online, GTS Direct Connectivity → commercial or institutional
    payment initiation; do NOT assume Institution to Institution unless there
    is evidence of interbank or counterparty settlement
  - Relationship Manager → assisted customer activity; not category evidence
  - Consumer, Commercial, Small Business, Institutional → party context;
    do not prove payer/payee direction by themselves
  - Institutional segment + clearing/settlement/ESA/RITS/interbank obligation
    language → supports Institution to Institution
  - These fields STRENGTHEN a mapping already suggested by R1-R3 evidence.
    They must not be the SOLE basis for a mapping.
  - If absent or generic, fall back to R1-R6 evidence only.

EVIDENCE HIERARCHY FOR CATEGORY MAPPING (apply in order — higher overrides lower):
  1. Explicit payer-payee relationship
  2. Explicit payment purpose
  3. Product / service context
  4. Beneficiary, borrower, merchant, biller, customer, institution or
     counterparty references
  5. Customer segment context
  6. Business context
  7. Channel context
  8. Payment rail / settlement mechanism
  9. Lifecycle stage
  10. Activity title

Minimum Mapping Requirement (Direct processes only)

Category mapping requires at least one of:
  (a) Explicit payer-payee relationship
  (b) Explicit payment action — settle, disburse, process payment, etc.
  (c) Explicit payment system/process reference — RTGS, BPAY, clearing, etc.
  (d) Explicit system, third-party, or governance reference in the context
      of payment processing
  (e) Customer segment + channel + process action that together unambiguously
      imply a payer-payee relationship
  (f) Payment instrument type + transaction direction + L2/L3 context that
      together unambiguously imply a category

If none of (a)-(f): mapped_categories=[], primary_category=null for Direct.
Explain why no category is defensible and state the closest alternative
considered.

CORE RULE 5 — PAYMENT-RELATED BUT CATEGORY-UNCERTAIN IS ALLOWED
If the process clearly handles payment activity but does not contain enough
evidence to determine the ultimate payer-payee relationship, return:
  is_payment_process=true, mapped_categories=[], primary_category=null,
  category_confidence_score=Low. Do not force a category.

──────────────────────────────────────────────────────────────────────────────
INSTRUMENT AND DIRECTION INFERENCE (Direct processes only — generic L3 activities)
──────────────────────────────────────────────────────────────────────────────

Do NOT rely only on explicit payer-payee wording in the L3 activity name.
If the L3 activity is a generic payment processing step (e.g. receive,
validate, process inbound payment), apply the following:

  Step A — Identify payment instrument / type (strongest single signal)
  e.g. direct debit, BPAY, EFT, cheque, NPP, RTGS, SWIFT, PEXA, Visa

  Step B — Identify transaction direction
  Inbound (funds coming in) vs outbound (funds going out)

  Step C — Identify L2 process context
  e.g. lending, supplier payments, customer transfers, payroll, clearing

  Step D — Determine economic purpose
  e.g. repayment, disbursement, transfer, settlement, payroll

Supporting signal guidance (apply with higher-ranked evidence — see hierarchy):
  - Direct Debit / BPAY as instrument → supports Customer to Institution
    (customer submitting payment to institution), but confirm with payment
    purpose — BPAY is a rail, not a category determinant on its own
  - Inbound + lending / loan context → supports Customer to Institution
  - Outbound + lending / loan context → supports Institution to Customer
  - Consumer EFT / transfer between non-bank parties → supports Customer to Customer
  - Interbank / counterparty / clearing / settlement + institutional context
    → supports Institution to Institution, but only where evidence shows
    interbank obligation — not solely because RTGS, SWIFT or clearing is mentioned
  - Payroll / vendor / supplier outbound → supports Supplier / Contractor / Employee

Remember Core Rule 2: payment rails alone (SWIFT, RTGS, BPAY, clearing,
Direct Entry) are NOT category determinants. Always confirm with payer-payee
evidence ranked higher in the evidence hierarchy.

Generic Direct processes MUST inherit classification from the dominant
payment use case. Only apply contextual inference when:
  - The process clearly sits within a defined payment use case, AND
  - There is no conflicting strong signal

Do NOT return empty mapped_categories for a Direct process if:
  - Direction is known, AND
  - A strong supporting signal (instrument, product context or L2 context) exists

──────────────────────────────────────────────────────────────────────────────
MAPPING RULES (Direct processes only)
──────────────────────────────────────────────────────────────────────────────

R1 — Direct Payer-Payee Relevance
Map only where a specific payer-to-payee relationship is clearly identifiable
or unambiguously inferable from a named payment system, use case, or the
combination of customer segment and process action.
  Include: payer and payee identifiable; relationship matches a category;
           inference directly from named payment system or segment + action.
  Exclude: no identifiable payer-payee; general activities with no payment
           relationship; cannot determine who pays whom.

R2 — Category-Specific Payment Context
Process must operate in a specific payment scenario, not generic activity.
  Include: description indicates who pays whom; references payment product
           processing; aligns with incomplete/inaccurate/duplicate/unauthorised
           payment scenarios.
  Exclude: strategy, planning, product design, marketing, pure HR (excluding
           payroll), general reporting without direct payment execution context;
           no funds movement; upstream activity before any payment exists.

R3 — Payment Instruction or Execution Link
Process must create, control, act upon a payment instruction, or enable,
secure, control, or recover systems/processes that execute payments.
  Include: creates payment instruction; approves/releases payments; processes
           payments between defined parties; disburses funds to employees/
           contractors/suppliers; activities supporting payment systems or
           lifecycle stages even where no direct payment action exists.
  Exclude: only where there is no identifiable linkage to any payment system,
           platform, or payment process.

R4 — Multi-Category Applicability
Multiple categories only where explicitly used across different payer-payee
relationships in a single activity.

Priority logic (apply in order):
  1. Prefer the category most explicitly supported by the process text.
  2. Prefer the category with the clearest payer-payee in l3 activity name,
     description, tasks, systems, third parties, channel, or segment.
  3. L3 activity name is the primary naming signal.
  4. Interbank settlement, clearing, external banks, RBA, RITS, OFIs,
     Institutional segment in settlement context → Institution to Institution.
  5. Customer repayments, deposits, ATM, BPAY, card repayments, Consumer/
     Commercial in repayment/deposit context → Customer to Institution.
  6. Disbursements, refunds, interest, loan settlements, welfare, benefit/tax,
     Consumer in disbursement context → Institution to Customer.
  7. Transfers between customers/merchants, P2P, direct debit, card payments,
     customer-originated movement → Customer to Customer.
  8. Vendor payments, internal staff payroll, supplier/contractor/employee
     payments → Supplier / Contractor / Employee.
  9. If genuinely across multiple flows with no dominant category, return all
     defensible categories and explain.
  Always set primary_category when any category is mapped.

R5 — Category-Aligned Third-Party or Scheme Participation
Category always determined by ultimate payer-payee, not intermediary.
  Include: systems/third parties in payment between defined payer and payee.
  Exclude: third-party governance only with no direct execution/control link.

R6 — Functional Naming Test
Naming supports but must never solely drive mapping (supports R1-R3 only).
  Positive signals: Process, Disburse, Settle, Pay, Transfer (with payer/payee).
  Negative signals (alone): Manage, Develop, Design, Oversee.

Confidence for category_confidence_score (Direct processes only):
  High   — Explicit payer-payee AND explicit payment action or system.
  Medium — Defensible but relies on contextual inference or combined fields.
  Low    — Tentative, limited text, or closest alternative assigned.

If Direct but no category is defensible:
  mapped_categories=[], primary_category=null, category_confidence_score=Low.
  State the closest alternative considered in mapping_rationale.

Contrastive reasoning requirement (Direct processes only):
  For every mapping_rationale, state why the chosen category applies AND
  why the closest alternative does not.

Workflow (per process):
  1. Review all supplied fields — l3_process_UUID, process_id, l2_process_name,
     l3_activity_name, description, tasks, systems, third_parties,
     governance_context, l3_activity_channels, l3_activity_customer_segments,
     l3_activity_product_service.
  2. Apply Step 0 — if Non-payment: stop, return empty mapping.
  3. Apply Step 1 — classify as Direct, Enabling, or Non-payment.
  4. IF Enabling: set mapped_categories=[], primary_category=null,
     category_confidence_score=null, mapping_rationale=null. Stop.
     Do not proceed to steps 5-13.
  5. IF Direct: test Minimum Mapping Requirement (conditions a-f).
  6. Apply R1 — payer-payee relationship present (including via segment/channel)?
  7. Apply R2 — category-specific payment context, not generic enterprise?
  8. Apply R3 — directly creates, controls, or acts on payment instruction?
  9. Apply R5 where systems, schemes, or third parties are involved.
  10. Apply R6 as naming support only.
  11. Emit category/categories. Apply R4 if multiple are defensible.
  12. Assign category_confidence_score.
  13. Include mapping_rationale with quoted text and contrastive reasoning.
      Cite channel/segment if used.
  14. If key fields are missing, state this and be conservative about confidence.
  15. If no direct payment action exists, identify the payment system or
      lifecycle stage supported and map the category based on that linkage.

──────────────────────────────────────────────────────────────────────────────
WORKED EXAMPLES
──────────────────────────────────────────────────────────────────────────────

Example 1 — Customer to Institution (Direct, High)
  l3_activity_name: Receive Incoming Funds
  Description: "Customer sends funds into bank account"
  Segment: "BM 03.01.01 - Consumer"
  payment_process_type: Direct
  mapped_categories: ["Customer to Institution"], primary_category: same, confidence: High
  Why: "Customer sends funds into bank account" explicitly identifies customer
  as payer and institution as payee, corroborated by Consumer segment. R1 and
  R2 satisfied. Customer to Customer less appropriate — payee is the bank.
  Institution to Customer less appropriate — bank is receiving, not paying.

Example 2 — Institution to Customer (Direct, High)
  l3_activity_name: Disburse Funds
  Description: "Institution disburses funds to the customer"
  payment_process_type: Direct
  mapped_categories: ["Institution to Customer"], primary_category: same, confidence: High
  Why: "Institution disburses funds to the customer" explicitly identifies
  institution as payer and customer as payee. R1 and R3 satisfied.
  Customer to Institution less appropriate — customer is not paying the bank.

Example 3 — Supplier / Contractor / Employee Payments (Direct, High)
  l3_activity_name: Create Payroll Payment Batch
  Description: "Periodic staff payroll runs"
  payment_process_type: Direct
  mapped_categories: ["Supplier / Contractor / Employee Payments"], confidence: High
  Why: "Payroll Payment Batch" and "staff payroll runs" directly support
  employee payments. NOT excluded at Step 0 — payroll is a valid category.

Example 4 — Institution to Institution (Direct, High)
  l3_activity_name: Settle Interbank Obligation
  Description: "Settlement between banks through clearing network"
  Segment: "BM 03.03.01 - Institutional"
  payment_process_type: Direct
  mapped_categories: ["Institution to Institution"], confidence: High
  Why: "Interbank Obligation" and "clearing network", corroborated by
  Institutional segment. R1, R3, and R5 satisfied.

Example 5 — Multiple categories (Direct, Medium)
  l3_activity_name: Generate Payment Advice
  Description: "Advice applies to both customers and counterparties"
  payment_process_type: Direct
  mapped_categories: ["Institution to Customer", "Institution to Institution"],
  primary_category: Institution to Customer, confidence: Medium
  Why: Text supports both audiences. R4 applied. Primary assigned to the most
  explicitly stated audience.

Example 6 — Enabling, no category (ITGC)
  l3_activity_name: Manage Business Continuity Testing
  Description: "Periodic testing and review of Business Continuity Plans to
  ensure resilience of operations including payment processing"
  payment_process_type: Enabling
  is_payment_process: true
  mapped_categories: []
  primary_category: null
  category_confidence_score: null
  mapping_rationale: null
  Why: BCP and resilience processes at a bank are Enabling by nature. The
  process supports payment operations continuity but does not directly execute
  a payment transaction. Under the routing rule, Enabling processes do not
  proceed to category mapping. No category assigned regardless of any
  segment, channel, or L2 context present.

Example 7 — Enabling, no category (Privileged Access)
  l3_activity_name: Manage Privileged Access to Payment Systems
  Description: "Review and approve privileged access requests to SWIFT
  messaging infrastructure used for international payment settlement"
  payment_process_type: Enabling
  is_payment_process: true
  payment_process_confidence: High
  mapped_categories: []
  primary_category: null
  mapping_rationale: null
  Why: Access governance directly tied to a named payment system (SWIFT).
  Classified as Enabling — it controls access to payment infrastructure but
  does not directly execute, settle, or reconcile payments. Routing rule
  stops evaluation here. No category assigned.

Example 8 — Non-payment (Step 0 exclusion)
  l3_activity_name: Design New Term Deposit Product Features
  Description: "Product team designs new term deposit features based on
  market research and competitor analysis"
  payment_process_type: Non-payment
  is_payment_process: false, exclusion_gate_applied: true
  Why: Pure product design. No payment processing, payment system, or
  payment governance reference. No plausible payment signal.

Example 9 — Direct, no defensible category (closest alternative stated)
  l3_activity_name: Monitor Payment Exception Queue
  Description: "Team monitors exception queue for flagged payment transactions"
  payment_process_type: Direct
  is_payment_process: true
  mapped_categories: [], primary_category: null, confidence: Low
  Why: Payment-related (Direct — monitors payment transactions). Exception
  queue covers all payment types — no payer-payee evidence allows a specific
  category. Closest alternative considered was Customer to Institution (most
  common flow at a retail bank) but insufficient evidence to assign it.

──────────────────────────────────────────────────────────────────────────────
ADDITIONAL WORKED EXAMPLES (SME-validated)
──────────────────────────────────────────────────────────────────────────────

Example 10 — Not payment-related: contract documentation
  L2 Process: Prepare Product and Service Arrangement
  L3 Activity: Prepare Contract Documentation
  payment_process_type: Non-payment, is_payment_process: false
  Why: Prepares terms, conditions, financial tables and customer documentation.
  Does not process, authorise, execute, settle, post or reconcile payments.
  This is product/contract lifecycle, not payment lifecycle.

Example 11 — Not payment-related: account/facility setup
  L2 Process: Implement Product Arrangement
  L3 Activity: Set-up Product Account / Facility
  payment_process_type: Non-payment, is_payment_process: false
  Why: Creates and configures accounts or facilities. Product/account lifecycle
  management. Not payment lifecycle management.

Example 12 — SWIFT does not determine Institution to Institution
  L3 Activity: Authorise Outbound Payments - SWIFT
  payment_process_type: Direct, is_payment_process: true
  Category: Do NOT automatically assign Institution to Institution.
  Why: SWIFT is a payment rail. Determine category from ultimate payer-payee.
  Could be Institution to Customer (SWIFT international wire to customer),
  Institution to Institution (interbank settlement), or other depending on
  payment purpose and counterparty evidence.

Example 13 — Loan repayment always Customer to Institution
  Context: Payment purpose is loan repayment, mortgage repayment, credit card
  repayment, lease repayment or overdraft repayment.
  Category: Customer to Institution
  Why: Customer is paying an obligation owed to the institution. This applies
  even if the transaction uses RTGS, SWIFT, BPAY, Direct Entry, clearing or
  OFI settlement behind the scenes (Core Rule 4).

Example 14 — Loan disbursement always Institution to Customer
  Context: Payment purpose is loan drawdown, loan settlement, loan
  disbursement or redraw to the customer or customer-nominated account.
  Category: Institution to Customer
  Why: Institution is paying or disbursing funds to the customer or their
  nominated beneficiary.

──────────────────────────────────────────────────────────────────────────────
IDENTIFIER RULE
──────────────────────────────────────────────────────────────────────────────

Each process includes l3_process_UUID. This is the unique join key.
Copy it exactly into every output mapping. Do not alter or truncate.
Do not use l3_activity_id or process_id as the primary key — the same
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
      "mapped_categories": ["valid category strings — or empty array"],
      "primary_category": "valid category string if Direct, else null",
      "category_confidence_score": "High | Medium | Low if Direct, else null",
      "rule_hits": ["R1", "R2", ... if Direct, else empty array],
      "channel_segment_used": true | false,
      "product_service_used": true | false,
      "mapping_rationale": "quoted text + contrastive reasoning if Direct, else null"
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
- Every payment_process_rationale must include at least one quoted phrase.
- Payroll, staff payment, vendor payment, and contractor payment processes are
  NEVER excluded — they are Direct and map to Supplier / Contractor / Employee.
- IT security, ITGC, BCP, change management, access management, governance,
  and resilience processes are NEVER Non-payment — they are Enabling at minimum.
- If payment_process_type = "Enabling":
    mapped_categories=[], primary_category=null, mapping_rationale=null.
    This is correct. Do not assign a category under any circumstances.
- If payment_process_type = "Direct":
    Always attempt to assign primary_category. Use closest alternative at Low
    confidence when evidence is thin. Only return null when Core Rule 5 applies
    and no defensible category signal exists after exhausting all evidence —
    in that case fully explain and state the closest alternative considered.
- If payment_process_type = "Non-payment":
    mapped_categories=[], primary_category=null.
- Multiple categories require distinct payer-payee evidence for each.
- If primary_category is set, it must appear in mapped_categories.
- Channel and customer segment must corroborate, not solely drive, a mapping.
- Do not invent facts. All inference must be grounded in supplied fields.
"""



# ──────────────────────────────────────────────────────────────────────────────
#  AUGMENTATION FUNCTIONS


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

    channel  = str(process_payload.get("l3_activity_channels", "")).strip()
    segment  = str(process_payload.get("l3_activity_customer_segments", "")).strip()
    product  = str(process_payload.get("l3_activity_product_service", "")).strip()
    has_real_evidence = bool(channel) or bool(segment) or bool(product)

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
    # Note: Enabling processes with null category are correct — do not flag
    # them solely for missing category. Only flag genuine quality issues.
    is_enabling = m.get("payment_process_type") == "Enabling"
    m["sme_review_flag"] = (
        not m["citation_verified"]
        or m["confidence_conflict"]
        or channel_segment_hallucinated
        or coverage["field_coverage_tier"] == "Low"
        or (not is_enabling and m.get("category_confidence_score") == "Low")
        or m.get("payment_process_confidence") == "Low"
    )

    return m



# ──────────────────────────────────────────────────────────────────────────────
#  CHECKPOINT / RESUME
# ──────────────────────────────────────────────────────────────────────────────

def get_output_paths(input_file: str) -> dict:
    """Derive all output file paths from the input file path."""
    p      = Path(input_file)
    stem   = p.stem.replace("_aggregated_for_llm_uuid", "")
    suffix = f"_llm_payment_relevance_category_uuid_{REASONING_EFFORT}"
    return {
        "dir":   p.parent,
        "stem":  stem,
        "jsonl": p.parent / f"{stem}{suffix}_raw.jsonl",
        "json":  p.parent / f"{stem}{suffix}.json",
        "xlsx":  p.parent / f"{stem}{suffix}.xlsx",
    }


def load_checkpoint(jsonl_path: Path) -> set:
    """
    Read existing JSONL output to find already-processed UUIDs.
    Called at startup to enable resume. Returns set of UUID strings.
    """
    processed = set()
    if not jsonl_path.exists():
        return processed
    with open(jsonl_path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                record = json.loads(line)
                for uuid in record.get("input_l3_process_UUIDs", []):
                    processed.add(str(uuid))
            except json.JSONDecodeError:
                continue
    return processed


def write_batch_checkpoint(record: dict, jsonl_path: Path) -> None:
    """
    Append a single batch record to the JSONL checkpoint file.
    Uses ensure_ascii=True for Windows compatibility (avoids OSError 22).
    Strips raw_response before writing — it can contain problematic characters
    and is not needed for checkpoint reconstruction.
    """
    safe_record = {k: v for k, v in record.items() if k != "raw_response"}
    try:
        line = json.dumps(safe_record, ensure_ascii=True) + "\n"
        with open(jsonl_path, "a", encoding="utf-8") as f:
            f.write(line)
    except (OSError, TypeError, ValueError) as e:
        # Log the error but do not crash — the batch result is still held
        # in memory and will be used for output generation at the end.
        print(f"  WARNING: Could not write batch {record.get('batch_no')} "
              f"to checkpoint: {e}. Run may not be resumable from this point.")


def reconstruct_from_jsonl(jsonl_path: Path) -> tuple:
    """
    Read the complete JSONL checkpoint to reconstruct batch_results and
    process_lookup_by_batch. Used at the end to generate final outputs
    consistently whether the run was fresh or resumed.
    Returns (batch_results, process_lookup_by_batch, all_warnings).
    """
    batch_results           = []
    process_lookup_by_batch = {}
    all_warnings            = []

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
            warnings = rec.get("warnings", [])

            batch_results.append({
                "batch_no":   batch_no,
                "mappings":   mappings,
                "notes":      rec.get("parsed_response", {}).get("notes", []),
                "warnings":   warnings,
                "latency_ms": rec.get("latency_ms"),
                "usage":      rec.get("usage"),
            })

            # Rebuild process_lookup from source_review_fields stored in JSONL
            source_fields = rec.get("source_review_fields", {})
            process_lookup_by_batch[batch_no] = source_fields
            all_warnings.extend([f"Batch {batch_no}: {w}" for w in warnings])

    return batch_results, process_lookup_by_batch, all_warnings

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
        # Only Direct processes require primary_category.
        # Enabling processes must have null — warn if the model assigned one.
        if ptype == "Direct" and not primary:
            warnings.append(
                f"Direct process has null primary_category for UUID {uid}. "
                f"Closest alternative should have been assigned."
            )
        if ptype == "Enabling" and primary:
            warnings.append(
                f"PROMPT VIOLATION: Enabling process has primary_category "
                f"\"{primary}\" for UUID {uid}. Must be null for Enabling."
            )
        if ptype == "Enabling" and (m.get("mapped_categories") or []):
            warnings.append(
                f"PROMPT VIOLATION: Enabling process has mapped_categories "
                f"for UUID {uid}. Must be empty for Enabling."
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

            # Review fields sourced from stored source_review_fields (in JSONL)
            # or directly from the full payload if available. This works for both
            # fresh runs and reconstructed checkpoint runs.
            rows.append({
                # Identifiers
                "batch_no":                  batch_no,
                "l3_process_UUID":           m.get("l3_process_UUID"),
                "process_id":                m.get("process_id"),
                # Review context
                "l2_process_name":           source.get("l2_process_name", m.get("l2_process_name", "")),
                "l2_process_description":    source.get("l2_process_description", ""),
                "l3_activity_name":          source.get("l3_activity_name", m.get("l3_activity_name", "")),
                "l3_activity_description":   source.get("l3_activity_description",
                                                 source.get("description", "")),
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
                "product_service_used":      m.get("product_service_used"),
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
    # Only Direct processes are expected to have primary_category
    direct_with_null = (
        (df["payment_process_type"] == "Direct") & df["primary_category"].isna()
    ).sum() if "payment_process_type" in df and "primary_category" in df else 0
    null_primary = direct_with_null

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
        ("Missing primary_category (Direct only)", int(null_primary)),
        ("──────────────────────────", "──────────────"),
    ] + [(f"Category: {k}", v) for k, v in cat_dist.items()]

    return pd.DataFrame(rows, columns=["Metric", "Count"])


def write_outputs(input_file: str, batch_results: list,
                  raw_records: list, all_warnings: list,
                  process_lookup_by_batch: dict,
                  paths: dict = None) -> None:
    """Generate final JSON and Excel outputs from batch results."""
    if paths is None:
        paths = get_output_paths(input_file)
    json_path  = paths["json"]
    xlsx_path  = paths["xlsx"]

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
        json.dump(combined, f, ensure_ascii=True, indent=2)

    # JSONL is written incrementally per batch — no bulk write here
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
        description="Payment relevance + category mapping. Supports checkpoint/resume."
    )
    parser.add_argument("input_file", nargs="?", default=DEFAULT_INPUT_FILE)
    parser.add_argument("--batch-size", type=int, default=BATCH_SIZE)
    parser.add_argument("--force", action="store_true",
                        help="Ignore existing checkpoint and reprocess from scratch.")
    args = parser.parse_args()

    if not os.path.exists(args.input_file):
        raise FileNotFoundError(f"Input file not found: {args.input_file}")

    paths = get_output_paths(args.input_file)

    print("=" * 72)
    print("  Payment Relevance + Category LLM Runner — l3_process_UUID")
    print(f"  Model                 : {MODEL}")
    print(f"  Reasoning effort      : {REASONING_EFFORT}")
    print(f"  Max completion tokens : {MAX_COMPLETION_TOKENS}")
    print(f"  Batch size            : {args.batch_size}")
    print(f"  Input file            : {args.input_file}")
    print(f"  JSONL checkpoint      : {paths['jsonl']}")
    print("=" * 72)

    all_processes = load_processes(args.input_file)

    missing_uuid = [
        p.get("process_id") for p in all_processes if not p.get("l3_process_UUID")
    ]
    if missing_uuid:
        raise ValueError(
            f"Processes missing l3_process_UUID: {missing_uuid[:5]}... "
            "Re-run the aggregation script to add UUIDs."
        )

    print(f"\n  Loaded {len(all_processes):,} process payloads")
    preflight_check()

    # ── Checkpoint / resume ─────────────────────────────────────────────────
    if args.force and paths["jsonl"].exists():
        paths["jsonl"].unlink()
        print("\n  --force: existing checkpoint deleted. Starting from scratch.")

    processed_uuids = load_checkpoint(paths["jsonl"])
    if processed_uuids:
        print(f"\n  Checkpoint found: {len(processed_uuids):,} processes already done.")
    else:
        print("\n  No checkpoint found. Starting fresh run.")

    processes = [
        p for p in all_processes
        if str(p.get("l3_process_UUID")) not in processed_uuids
    ]

    if not processes:
        print("  All processes already completed. Regenerating output files...")
        batch_results, process_lookup_by_batch, all_warnings = reconstruct_from_jsonl(paths["jsonl"])
        write_outputs(args.input_file, batch_results, [], all_warnings,
                      process_lookup_by_batch, paths)
        print("\n  Done.")
        return

    total_batches_remaining = (len(processes) + args.batch_size - 1) // args.batch_size
    print(f"  Remaining to process : {len(processes):,}")
    print(f"  Already completed    : {len(processed_uuids):,}")
    print(f"  Batches remaining    : {total_batches_remaining:,}\n")

    print("  Initialising Azure OpenAI client...")
    client = init_client()
    print("  Client ready.\n")

    batch_results           = []
    raw_records             = []
    process_lookup_by_batch = {}
    total_in_tok   = 0
    total_out_tok  = 0
    run_start      = time.time()

    for batch_idx, (batch_no, batch) in enumerate(
            chunk_list(processes, args.batch_size), 1):

        # ── Progress display ─────────────────────────────────────────────────
        pct   = batch_idx / total_batches_remaining * 100
        bar   = "#" * int(pct / 5) + "." * (20 - int(pct / 5))
        elapsed = time.time() - run_start
        eta_s   = (elapsed / batch_idx) * (total_batches_remaining - batch_idx) if batch_idx > 1 else 0
        eta_str = (f"{int(eta_s//3600)}h {int((eta_s%3600)//60)}m"
                   if eta_s > 60 else f"{int(eta_s)}s")
        cost_usd = (total_in_tok / 1_000_000 * INPUT_PRICE_USD_PER_M +
                    total_out_tok / 1_000_000 * OUTPUT_PRICE_USD_PER_M)
        cost_aud = cost_usd / AUD_USD_RATE

        print(f"  [{bar}] {pct:5.1f}%  "
              f"Batch {batch_idx}/{total_batches_remaining}  "
              f"ETA {eta_str}  "
              f"Cost A${cost_aud:.2f}")

        parsed, meta = call_llm(client, batch, batch_no)

        # Structural validation
        struct_warnings = validate_output(parsed, batch)

        # Post-processing augmentation
        process_lookup = {str(p.get("l3_process_UUID")): p for p in batch}
        process_lookup_by_batch[batch_no] = process_lookup
        augmented_mappings = []
        for m in parsed.get("mappings", []):
            uuid    = str(m.get("l3_process_UUID", ""))
            payload = process_lookup.get(uuid, {})
            augmented_mappings.append(augment_mapping(m, payload))
        parsed["mappings"] = augmented_mappings

        dist_warnings  = check_batch_distribution(augmented_mappings)
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

        # Store source review fields for JSONL reconstruction
        source_review_fields = {
            str(p.get("l3_process_UUID")): {
                "l2_process_name":        str(p.get("l2_process_name", "") or ""),
                "l2_process_description": str(p.get("l2_process_description", "") or ""),
                "l3_activity_name":       str(p.get("l3_activity_name", "") or ""),
                "l3_activity_description":str(p.get("l3_activity_description",
                                               p.get("description", "")) or ""),
                "l3_activity_product_service":str(p.get("l3_activity_product_service", "") or ""),
            }
            for p in batch
        }
        raw_record = {
            "batch_no":               batch_no,
            "input_l3_process_UUIDs": [p.get("l3_process_UUID") for p in batch],
            "source_review_fields":   source_review_fields,
            "parsed_response":        parsed,
            "raw_response":           meta.get("raw_response"),
            "usage":                  meta.get("usage"),
            "latency_ms":             meta.get("latency_ms"),
            "warnings":               batch_warnings,
        }
        raw_records.append(raw_record)
        write_batch_checkpoint(raw_record, paths["jsonl"])

        # Per-batch summary
        usage      = result["usage"] or {}
        in_tok     = usage.get("input_tokens") or 0
        out_tok    = usage.get("output_tokens") or 0
        total_in_tok  += in_tok
        total_out_tok += out_tok
        sme_count  = sum(1 for m in augmented_mappings if m.get("sme_review_flag"))
        rt = (f" | {usage.get('reasoning_tokens')} thinking"
              if usage.get("reasoning_tokens") else "")
        print(f"         {len(augmented_mappings)} mappings  "
              f"{result['latency_ms']}ms  "
              f"{in_tok} in / {out_tok} out{rt}  "
              f"SME: {sme_count}")
        if batch_warnings:
            for w in batch_warnings:
                print(f"         WARNING: {w}")

        if INTER_BATCH_SLEEP > 0:
            time.sleep(INTER_BATCH_SLEEP)

    # ── Final output generation ─────────────────────────────────────────────
    print("\n  Generating final outputs from complete JSONL...")
    all_batch_results, all_lookup, all_warnings_full = reconstruct_from_jsonl(paths["jsonl"])

    final_cost_usd = (total_in_tok / 1_000_000 * INPUT_PRICE_USD_PER_M +
                      total_out_tok / 1_000_000 * OUTPUT_PRICE_USD_PER_M)
    final_cost_aud = final_cost_usd / AUD_USD_RATE
    print(f"  This run:  {len(processes):,} processes  "
          f"{total_in_tok:,} input tokens  "
          f"{total_out_tok:,} output tokens  "
          f"A${final_cost_aud:.2f}")

    write_outputs(args.input_file, all_batch_results, [], all_warnings_full,
                  all_lookup, paths)
    print("\n  Done.")

if __name__ == "__main__":
    main()
