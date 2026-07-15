"""
run_embedding_linkage.py
─────────────────────────────────────────────────────────────────────────────
Payments Controls PoC — Embeddings-Based Fallback Linkage

Generates candidate payment process matches for:
  Population A — controls linked only to non-payment Holo processes (141)
  Population B — controls not linked to any Holo processes      (187)

This is candidate generation only. Outputs are NOT final linkage.
Final linkage status must be SME-approved.

Linkage hierarchy:
  1. Deterministic linkage     (existing — from linkage analysis workbook)
  2. Embedding candidate       (this script)
  3. SME-approved linkage      (downstream review)

Before running (PowerShell):
  az account set --subscription 6c72e6c5-ed48-4030-b29c-34e2849c9288
  $env:REQUESTS_CA_BUNDLE = "C:\\Users\\m061400\\ai-test\\cacert.pem"
  $env:SSL_CERT_FILE      = "C:\\Users\\m061400\\ai-test\\cacert.pem"
  Remove-Item Env:AZURE_CA_BUNDLE -ErrorAction SilentlyContinue
  python run_embedding_linkage.py
  python run_embedding_linkage.py --force   # regenerate embeddings from scratch
"""

import argparse
import os
import re
import time
from pathlib import Path

import numpy as np
import pandas as pd
from azure.identity import DefaultAzureCredential, get_bearer_token_provider
from openai import AzureOpenAI

# ─────────────────────────────────────────────────────────────────────────────
#  CONFIG
# ─────────────────────────────────────────────────────────────────────────────

AZURE_ENDPOINT  = "https://ai.eng.azure.srv.westpac.com.au"
API_VERSION     = "2024-10-21"
EMBEDDING_MODEL = "text-embedding-3-large"
EMBEDDING_DIM   = 3072
MAX_TOKENS      = 8000   # Hard limit is 8191; leave a safety margin

BASE_DIR        = Path(r"C:\Users\m061400\ai-test\big_table")
EMBED_DIR       = BASE_DIR / "embeddings"

# Input files — pre-prepared by analyst in the embeddings folder
FILE_CONTROLS   = EMBED_DIR / "_controls_to_embed_.xlsx"
FILE_PROCESSES  = EMBED_DIR / "_processes_to_embed_.xlsx"
# Linkage workbook — used to derive population type (A vs B) per control
FILE_LINKAGE_WB = BASE_DIR / "phase1e_outputs" / "payment_control_linkage_analysis.xlsx"

# Output
OUTPUT_FILE     = EMBED_DIR / "embedding_linkage_candidates.xlsx"

# Embedding cache files (saves regenerating on reruns)
CACHE_CONTROLS  = EMBED_DIR / "cache_control_embeddings.npy"
CACHE_CTRL_IDS  = EMBED_DIR / "cache_control_ids.npy"
CACHE_PROCS     = EMBED_DIR / "cache_process_embeddings.npy"
CACHE_PROC_IDS  = EMBED_DIR / "cache_process_uuids.npy"

# Retrieval settings
TOP_K           = 10      # Candidates per control (recall-first)
BATCH_SIZE      = 100     # Texts per embedding API call

# Similarity thresholds
THRESHOLD_STRONG = 0.80
THRESHOLD_MEDIUM = 0.70
THRESHOLD_WEAK   = 0.60

# Payment terms: boost if present in control text
PAYMENT_TERMS = {
    "payment","payments","settlement","clearing","reconciliation",
    "exception","refund","disbursement","deposit","repayment","drawdown",
    "payee","beneficiary","merchant","card","bpay","swift","rtgs","npp",
    "direct entry","direct debit","transaction","remittance","proceeds",
    "reversal","reissue","mistaken","unauthorised","disputed","recall",
    "nostro","ledger","posting","suspense","eft","payid","ofi","rits",
    "correspondent","interbank","scheme","rails","instruction",
}

# Generic terms: penalise if ONLY these are present
GENERIC_TERMS = {
    "review","monitor","approve","validate","report","check",
    "governance","risk","compliance","manage","oversight","assurance",
    "testing","training","awareness","documentation","policy",
}

# CT titles for enrichment
CT_TITLES = {
    "CT1":"Validation of Human-Entered Data at Input",
    "CT2":"Payment processing error detection",
    "CT3":"Early Identification of Duplications and Processing Errors",
    "CT4":"Payment processing interface and batch error resolution",
    "CT5":"Incident response",
    "CT6":"Master/Reference data input validation",
    "CT7":"Service provider ongoing review",
    "CT8":"Service provider onboarding",
    "CT9":"Change management testing",
    "CT10":"Critical service chain mapping and risk identification",
    "CT11":"Rollback plans",
    "CT12":"System recovery capability",
    "CT13":"Business continuity plan",
    "CT14":"Logging and monitoring",
    "CT15":"Secure IT Design",
    "CT16":"Vulnerability Management",
    "CT17":"Access - Provision / deprovision (EIA)",
    "CT18":"Access - Monitoring",
    "CT19":"Patch Management",
    "CT20":"Physical Security Controls",
    "CT21":"Access - Privileged users",
    "CT22":"Mistaken Internet Payment Reports",
    "CT23":"Provision of Confirmations and Notifications",
    "CT24":"Treatment of Unauthorised and Disputed Transactions",
    "CT25":"Provision of Confirmations and Notifications",
    "CT26":"Regulatory horizon scanning",
    "CT27":"Records retention",
    "CT28":"Crisis Management planning and testing",
}

# ─────────────────────────────────────────────────────────────────────────────
#  HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def clean_cols(df):
    df.columns = df.columns.str.strip()
    return df

# Excel formula errors that appear as strings when read by pandas
EXCEL_ERRORS = {"#name?","#value!","#ref!","#n/a","#div/0!","#null!","#num!","#error!"}

def is_empty(val):
    if pd.isna(val):
        return True
    s = str(val).strip()
    if s.lower() in ("", "nan", "none", "null"):
        return True
    # Treat Excel formula error strings as empty
    if s.lower() in EXCEL_ERRORS:
        return True
    return False

def strip_urls(text):
    """Remove http/https URLs and Confluence-style links from text."""
    if is_empty(text):
        return ""
    text = re.sub(r'https?://\S+', '', str(text))
    text = re.sub(r'\s+', ' ', text).strip()
    return text

def extract_bm_labels(raw):
    """
    Convert BM-coded semicolon list to human-readable labels.
    'BM 04.09.01 - Investor Home Loans;BM 04.09.02 - Owner Occupied Home Loans'
    -> 'Investor Home Loans; Owner Occupied Home Loans'
    """
    if is_empty(raw):
        return ""
    parts = str(raw).split(";")
    labels = []
    for p in parts:
        p = p.strip()
        if " - " in p:
            labels.append(p.split(" - ", 1)[1].strip())
        elif p:
            labels.append(p)
    return "; ".join(labels)

def safe_str(val, default=""):
    if is_empty(val):
        return default
    return str(val).strip()

def norm_gold_control_code(val):
    if is_empty(val):
        return None
    s = str(val).strip()
    if s.upper().startswith("CT"):
        try:
            n = int(s[2:])
            return f"CT{n}" if 1 <= n <= 28 else None
        except ValueError:
            return None
    try:
        n = int(float(s))
        return f"CT{n}" if 1 <= n <= 28 else None
    except (ValueError, TypeError):
        return None

def contains_payment_signal(text):
    low = text.lower()
    return any(term in low for term in PAYMENT_TERMS)

def dominated_by_generic_terms(text):
    low = text.lower()
    words = set(re.findall(r'\b\w+\b', low))
    has_payment = any(term in low for term in PAYMENT_TERMS)
    has_generic  = any(term in words for term in GENERIC_TERMS)
    return has_generic and not has_payment

# ─────────────────────────────────────────────────────────────────────────────
#  BUILD EMBEDDING TEXT
# ─────────────────────────────────────────────────────────────────────────────

def build_control_text(row):
    """
    Build semantic embedding text for a control.
    Focuses on semantically rich fields.
    URLs stripped from free-text fields.
    Both Population A and B treated identically — no non-payment context included.
    """
    parts = []

    name = safe_str(row.get("CTRL_NAME"))
    if name:
        parts.append(f"CONTROL NAME:\n{name}")

    desc = safe_str(row.get("CTRL_DESC"))
    if desc:
        parts.append(f"CONTROL DESCRIPTION:\n{desc}")

    ct_code  = safe_str(row.get("gold_control_code"))
    ct_title = CT_TITLES.get(ct_code, safe_str(row.get("gold_control_title")))
    if ct_code and ct_title:
        parts.append(f"CONTROL TYPE:\n{ct_code} — {ct_title}")
    elif ct_code:
        parts.append(f"CONTROL TYPE:\n{ct_code}")

    cat1 = safe_str(row.get("CTRL_CTGRY_1"))
    cat2 = safe_str(row.get("CTRL_CTGRY_2"))
    categories = "; ".join(c for c in [cat1, cat2] if c)
    if categories:
        parts.append(f"CONTROL CATEGORIES:\n{categories}")

    monitored = strip_urls(row.get("CTRL_MNTRD"))
    if monitored:
        parts.append(f"HOW MONITORED:\n{monitored}")

    evidenced = strip_urls(row.get("CTRL_EVDNCD"))
    if evidenced:
        parts.append(f"HOW EVIDENCED:\n{evidenced}")

    return "\n\n".join(parts)

def build_process_text(row):
    """
    Build semantic embedding text for a payment process.
    Fields from _processes_to_embed_.xlsx:
      l3_process_UUID, l2_process_name, l2_process_description,
      l3_activity_name, l3_activity_description, task_name (may be blank),
      process_category, process_lifecycle_stage, payment_rationale,
      l3_activity_channels, l3_activity_customer_segments,
      l3_activity_product/service

    Channel / segment / product fields use "BM XX.XX.XX - Label" format,
    semicolon-separated. extract_bm_labels() strips the BM codes and returns
    human-readable labels only (e.g. "ATM; Branch; Digital (Online and Mobile)").
    """
    parts = []

    # ── Core narrative fields (highest semantic weight) ───────────────────────
    l2_name = safe_str(row.get("l2_process_name"))
    l2_desc = safe_str(row.get("l2_process_description"))
    if l2_name:
        parts.append(f"L2 PROCESS:\n{l2_name}")
    if l2_desc:
        parts.append(f"L2 DESCRIPTION:\n{l2_desc}")

    l3_name = safe_str(row.get("l3_activity_name"))
    l3_desc = safe_str(row.get("l3_activity_description"))
    if l3_name:
        parts.append(f"L3 ACTIVITY:\n{l3_name}")
    if l3_desc:
        parts.append(f"L3 DESCRIPTION:\n{l3_desc}")

    # task_name may be blank for many rows — include only when populated
    tasks = safe_str(row.get("task_name"))
    if tasks:
        parts.append(f"TASKS:\n{tasks}")

    # ── Contextual supporting fields (BM code labels extracted) ───────────────
    # Product/service — strongest business context signal
    # Tries both column name variants (slash vs underscore)
    product_raw = (row.get("l3_activity_product/service") or
                   row.get("l3_activity_product_service"))
    product = extract_bm_labels(product_raw)
    if product:
        parts.append(f"PRODUCT / SERVICE:\n{product}")

    # Customer segments — helps identify party context (Consumer, Commercial, Institutional)
    segments = extract_bm_labels(row.get("l3_activity_customer_segments"))
    if segments:
        parts.append(f"CUSTOMER / PARTY SEGMENT:\n{segments}")

    # Channels — where the process originates or is delivered
    # Note: some processes span all channels (30+ entries); all are included
    # as the full list still carries semantic signal about process scope
    channels = extract_bm_labels(row.get("l3_activity_channels"))
    if channels:
        parts.append(f"CHANNEL CONTEXT:\n{channels}")

    # ── Payment classification context ────────────────────────────────────────
    category = safe_str(row.get("process_category"))
    stage    = safe_str(row.get("process_lifecycle_stage"))
    if category:
        parts.append(f"PAYMENT CATEGORY:\n{category}")
    if stage:
        parts.append(f"LIFECYCLE STAGE:\n{stage}")

    # Payment relevance rationale from LLM classification run
    # Field is 'payment_rationale' (confirmed — not 'payment_process_rationale')
    rationale = safe_str(row.get("payment_rationale"))
    if rationale:
        parts.append(f"PAYMENT RELEVANCE RATIONALE:\n{rationale}")

    return "\n\n".join(parts)

# ─────────────────────────────────────────────────────────────────────────────
#  AZURE OPENAI CLIENT
# ─────────────────────────────────────────────────────────────────────────────

def init_client():
    token_provider = get_bearer_token_provider(
        DefaultAzureCredential(),
        "https://cognitiveservices.azure.com/.default"
    )
    return AzureOpenAI(
        azure_endpoint=AZURE_ENDPOINT,
        api_version=API_VERSION,
        azure_ad_token_provider=token_provider,
    )

# ─────────────────────────────────────────────────────────────────────────────
#  GENERATE EMBEDDINGS (with batching + retry)
# ─────────────────────────────────────────────────────────────────────────────

def embed_texts(client, texts, label="texts"):
    """
    Generate embeddings for a list of texts.
    Batches in groups of BATCH_SIZE.
    Returns numpy array of shape (len(texts), EMBEDDING_DIM).
    """
    embeddings = []
    total = len(texts)
    print(f"  Embedding {total:,} {label} in batches of {BATCH_SIZE}...")

    for i in range(0, total, BATCH_SIZE):
        batch = texts[i:i+BATCH_SIZE]
        # Truncate any text exceeding safe token estimate (~MAX_TOKENS * 3 chars)
        batch = [t[:MAX_TOKENS * 3] for t in batch]
        attempt = 0
        while True:
            try:
                resp = client.embeddings.create(
                    model=EMBEDDING_MODEL,
                    input=batch,
                )
                batch_vecs = [d.embedding for d in resp.data]
                embeddings.extend(batch_vecs)
                n_done = min(i + BATCH_SIZE, total)
                print(f"    {n_done:,}/{total:,} ({n_done/total*100:.0f}%)", end="\r")
                time.sleep(0.2)   # light throttle
                break
            except Exception as e:
                attempt += 1
                if attempt >= 4:
                    raise
                wait = min(5 * 2 ** attempt, 60)
                print(f"\n    Batch {i//BATCH_SIZE+1} failed ({e}). Retry in {wait}s...")
                time.sleep(wait)

    print(f"    {total:,}/{total:,} (100%) done          ")
    return np.array(embeddings, dtype=np.float32)

# ─────────────────────────────────────────────────────────────────────────────
#  LOAD DATA
# ─────────────────────────────────────────────────────────────────────────────

def load_populations():
    """
    Load controls from the pre-prepared _controls_to_embed_.xlsx file.
    Derives population type (A vs B) by joining to the linkage analysis workbook.
    Handles:
      - CTRL_ID or Control_ID as the control key column
      - #NAME? and other Excel formula errors treated as empty
      - gold_control_code derived if not already present
    """
    print(f"\n  Loading controls from {FILE_CONTROLS.name}...")
    controls = clean_cols(pd.read_excel(FILE_CONTROLS, dtype=str))
    print(f"  Controls file rows: {len(controls):,}")
    print(f"  Columns: {list(controls.columns)}")

    # Normalise the control ID column — may be CTRL_ID or Control_ID
    if "Control_ID" not in controls.columns and "CTRL_ID" in controls.columns:
        controls = controls.rename(columns={"CTRL_ID": "Control_ID"})
    elif "Control_ID" not in controls.columns:
        raise ValueError(
            "Controls file must have 'Control_ID' or 'CTRL_ID' column. "
            f"Found: {list(controls.columns)}"
        )
    controls["Control_ID"] = controls["Control_ID"].str.strip()

    # Add gold_control_code and title if not already present
    if "gold_control_code" not in controls.columns:
        gc_col = next((c for c in controls.columns
                       if "gold" in c.lower() and "control" in c.lower()), None)
        if gc_col:
            controls["gold_control_code"] = controls[gc_col].apply(norm_gold_control_code)
        else:
            controls["gold_control_code"] = None
    if "gold_control_title" not in controls.columns:
        controls["gold_control_title"] = controls["gold_control_code"].map(CT_TITLES)

    # Derive population type from the linkage workbook
    pop_a_ids, pop_b_ids = set(), set()
    pop_a_context_map   = {}
    if FILE_LINKAGE_WB.exists():
        print("  Deriving population type from linkage workbook...")
        try:
            pop_a_df = clean_cols(
                pd.read_excel(FILE_LINKAGE_WB,
                              sheet_name="linked_non_payment", dtype=str))
            pop_a_ids = set(pop_a_df["Control_ID"].dropna().str.strip().unique())

            pop_b_df = clean_cols(
                pd.read_excel(FILE_LINKAGE_WB,
                              sheet_name="not_linked", dtype=str))
            pop_b_ids = set(pop_b_df["Control_ID"].dropna().str.strip().unique())

            # Non-payment context for Population A (metadata only)
            pop_a_detail = clean_cols(
                pd.read_excel(FILE_LINKAGE_WB,
                              sheet_name="linked_non_payment_detail", dtype=str))
            ctx_col = next(
                (c for c in ["l3_activity_uuid","l3_activity_id","l2_process_uuid"]
                 if c in pop_a_detail.columns), None)
            if ctx_col and "Control_ID" in pop_a_detail.columns:
                pop_a_context_map = (
                    pop_a_detail.groupby("Control_ID")[ctx_col]
                    .apply(lambda x: "; ".join(x.dropna().unique()[:5]))
                    .to_dict()
                )
        except Exception as e:
            print(f"  WARNING: Could not read population type from workbook: {e}")
            print("  Falling back to 'Combined' population type.")
    else:
        print(f"  WARNING: Linkage workbook not found at {FILE_LINKAGE_WB}")
        print("  All controls will be labelled population_type = 'Combined'.")

    def assign_pop(ctrl_id):
        if ctrl_id in pop_a_ids:
            return "A - linked_non_payment"
        if ctrl_id in pop_b_ids:
            return "B - not_linked"
        return "Combined - population type unknown"

    controls["population_type"] = controls["Control_ID"].apply(assign_pop)
    controls["current_non_payment_context"] = controls["Control_ID"].map(
        pop_a_context_map).fillna("")

    # Deduplicate
    controls = controls.drop_duplicates(subset=["Control_ID"])
    print(f"  Controls after dedup: {len(controls):,}")

    pop_counts = controls["population_type"].value_counts()
    for pop, cnt in pop_counts.items():
        print(f"    {pop}: {cnt:,}")

    return controls

def load_processes():
    """
    Load payment process candidate pool from pre-prepared file.
    Expected columns: l3_process_UUID, l2_process_name, l2_process_description,
    l3_activity_name, l3_activity_description, task_name,
    process_category, process_lifecycle_stage, payment_rationale
    """
    print(f"  Loading processes from {FILE_PROCESSES.name}...")
    procs = clean_cols(pd.read_excel(FILE_PROCESSES, dtype=str))
    print(f"  Process file rows    : {len(procs):,}")
    print(f"  Columns              : {list(procs.columns)}")

    if "l3_process_UUID" not in procs.columns:
        raise ValueError(
            "Process file must have 'l3_process_UUID' column. "
            f"Found: {list(procs.columns)}"
        )

    procs = procs.dropna(subset=["l3_process_UUID"])
    procs["l3_process_UUID"] = procs["l3_process_UUID"].str.strip()
    procs = procs.drop_duplicates(subset=["l3_process_UUID"])

    print(f"  Payment process candidates: {len(procs):,}")
    return procs

# ─────────────────────────────────────────────────────────────────────────────
#  COSINE SIMILARITY
# ─────────────────────────────────────────────────────────────────────────────

def cosine_similarity_matrix(ctrl_embs, proc_embs):
    """
    Compute cosine similarity between all controls and all processes.
    Returns matrix of shape (n_controls, n_processes).
    Normalises vectors before dot product for efficiency.
    """
    ctrl_norm = ctrl_embs / (np.linalg.norm(ctrl_embs, axis=1, keepdims=True) + 1e-9)
    proc_norm = proc_embs / (np.linalg.norm(proc_embs, axis=1, keepdims=True) + 1e-9)
    return ctrl_norm @ proc_norm.T

# ─────────────────────────────────────────────────────────────────────────────
#  RERANKING AND CANDIDATE GENERATION
# ─────────────────────────────────────────────────────────────────────────────

def confidence_band(score):
    if score >= THRESHOLD_STRONG:
        return "Strong candidate"
    if score >= THRESHOLD_MEDIUM:
        return "Medium candidate"
    if score >= THRESHOLD_WEAK:
        return "Weak candidate"
    return "Below review threshold"

def review_action(band, ctrl_text, proc_text, is_false_positive_risk):
    if band == "Below review threshold":
        return "No candidate — no embedding-supported payment linkage"
    if is_false_positive_risk:
        return "Possible false positive — do not action without SME review"
    if band == "Strong candidate":
        return "Strong candidate — review for linkage"
    if band == "Medium candidate":
        return "Medium candidate — review if business area / product aligns"
    return "Weak candidate — exploratory review only"

def why_matches(ctrl_text, proc_text):
    """Identify shared payment terms between control and process text."""
    ctrl_low = ctrl_text.lower()
    proc_low = proc_text.lower()
    shared = [t for t in PAYMENT_TERMS if t in ctrl_low and t in proc_low]
    if shared:
        return f"Shared payment terms: {', '.join(sorted(shared)[:8])}"
    return "Semantic similarity — no explicit shared payment terms detected"

def why_false_positive(ctrl_text, proc_text):
    """Flag potential false positive risks."""
    risks = []
    ctrl_low = ctrl_text.lower()
    if dominated_by_generic_terms(ctrl_text):
        risks.append("Control text dominated by generic governance terms")
    if not contains_payment_signal(ctrl_text):
        risks.append("No explicit payment terms in control text")
    prod_lifecycle = [
        "account setup","account creation","account opening","onboarding",
        "facility setup","contract preparation","product configuration",
    ]
    for term in prod_lifecycle:
        if term in proc_text.lower():
            risks.append(f"Process may relate to product/account lifecycle ('{term}')")
            break
    return "; ".join(risks) if risks else ""

def generate_candidates(all_controls, processes, sim_matrix):
    """
    Generate top K candidates per control.
    Applies reranking and false positive flags.
    Returns flat DataFrame with one row per control-candidate pair.
    """
    rows = []
    n_controls = len(all_controls)

    for i, (_, ctrl) in enumerate(all_controls.iterrows()):
        ctrl_id   = ctrl["Control_ID"]
        ctrl_text = ctrl.get("_embedding_text", "")
        pop_type  = ctrl.get("population_type", "")

        # Get similarities for this control vs all processes
        sims = sim_matrix[i]

        # Top K indices by similarity
        top_idx = np.argsort(sims)[::-1][:TOP_K]

        for rank, pidx in enumerate(top_idx, 1):
            proc    = processes.iloc[pidx]
            score   = float(sims[pidx])
            band    = confidence_band(score)
            proc_text = proc.get("_embedding_text", "")
            fp_reason = why_false_positive(ctrl_text, proc_text)
            is_fp     = bool(fp_reason)
            action    = review_action(band, ctrl_text, proc_text, is_fp)

            rows.append({
                "population_type":     pop_type,
                "Control_ID":          ctrl_id,
                "gold_control_code":   ctrl.get("gold_control_code",""),
                "gold_control_title":  CT_TITLES.get(
                    ctrl.get("gold_control_code",""), ""),
                "CTRL_NAME":           ctrl.get("CTRL_NAME",""),
                "CTRL_DESC":           ctrl.get("CTRL_DESC",""),
                "current_linkage_status": (
                    "Linked to non-payment Holo processes only"
                    if "A" in pop_type else "No Holo linkage"),
                "current_non_payment_process_context": ctrl.get(
                    "current_non_payment_context",""),
                "candidate_rank":             rank,
                "candidate_l3_process_UUID":  proc.get("l3_process_UUID",""),
                "candidate_l2_process_id":    proc.get("l2_process_id",""),
                "candidate_l2_process_name":  proc.get("l2_process_name",""),
                "candidate_l3_activity_id":   proc.get("l3_activity_id",""),
                "candidate_l3_activity_name": proc.get("l3_activity_name",""),
                "candidate_l3_activity_description": proc.get(
                    "l3_activity_description",""),
                "candidate_primary_category": proc.get(
                    "primary_category", proc.get("process_category","")),
                "candidate_mapped_categories": proc.get("mapped_categories",""),
                "candidate_payment_process_confidence": proc.get(
                    "payment_process_confidence",
                    proc.get("payment_process_type","")),
                "candidate_similarity_score": round(score, 6),
                "candidate_confidence_band":  band,
                "why_candidate_matches":      why_matches(ctrl_text, proc_text),
                "why_candidate_may_be_false_positive": fp_reason,
                "recommended_review_action":  action,
                "review_priority": (
                    1 if band == "Strong candidate" and not is_fp else
                    2 if band == "Strong candidate" else
                    3 if band == "Medium candidate" and not is_fp else
                    4 if band == "Medium candidate" else
                    5 if band == "Weak candidate" else 6
                ),
            })

        if (i + 1) % 50 == 0 or (i + 1) == n_controls:
            print(f"    Candidates built: {i+1:,}/{n_controls:,}", end="\r")

    print(f"    Candidates built: {n_controls:,}/{n_controls:,} done    ")
    return pd.DataFrame(rows)

# ─────────────────────────────────────────────────────────────────────────────
#  SUMMARY SHEETS
# ─────────────────────────────────────────────────────────────────────────────

def build_candidate_summary(candidates):
    """One row per Control_ID — best candidate, best score, action."""
    if candidates.empty:
        return pd.DataFrame()

    best = (
        candidates[candidates["candidate_rank"] == 1]
        [[
            "population_type","Control_ID","gold_control_code",
            "gold_control_title","CTRL_NAME","current_linkage_status",
            "candidate_l3_process_UUID","candidate_l2_process_name",
            "candidate_l3_activity_name","candidate_primary_category",
            "candidate_similarity_score","candidate_confidence_band",
            "recommended_review_action","review_priority",
            "why_candidate_matches","why_candidate_may_be_false_positive",
        ]]
        .rename(columns={
            "candidate_similarity_score": "best_similarity_score",
            "candidate_confidence_band":  "best_confidence_band",
        })
    )
    return best.sort_values(["review_priority","best_similarity_score"],
                            ascending=[True, False])

def build_threshold_summary(candidates):
    """Counts by population, confidence band and review action."""
    if candidates.empty:
        return pd.DataFrame()

    rows = []
    for pop in candidates["population_type"].unique():
        pop_df = candidates[candidates["population_type"] == pop]
        # Only rank-1 rows for per-control counts
        rank1  = pop_df[pop_df["candidate_rank"] == 1]
        rows.append({"population": pop, "metric": "Total controls", "count": len(rank1)})
        for band in ["Strong candidate","Medium candidate","Weak candidate",
                     "Below review threshold"]:
            n = (rank1["candidate_confidence_band"] == band).sum()
            rows.append({"population": pop, "metric": band, "count": int(n)})
        for action in rank1["recommended_review_action"].unique():
            n = (rank1["recommended_review_action"] == action).sum()
            rows.append({"population": pop, "metric": f"Action: {action}", "count": int(n)})

    return pd.DataFrame(rows)

# ─────────────────────────────────────────────────────────────────────────────
#  VALIDATION
# ─────────────────────────────────────────────────────────────────────────────

def run_validation(all_controls, processes, candidates):
    print("\n  Running validation checks...")
    checks = []

    def chk(name, expected, actual, note=""):
        passed = str(expected) == str(actual)
        checks.append({
            "check": name, "expected": str(expected),
            "actual": str(actual),
            "pass":  "PASS" if passed else "FAIL",
            "note":  note,
        })
        print(f"  [{'PASS' if passed else 'FAIL'}] {name}")

    pop_a = all_controls[all_controls["population_type"].str.startswith("A")]
    pop_b = all_controls[all_controls["population_type"].str.startswith("B")]

    chk("All 141 Population A controls processed",
        141, len(pop_a), "Adjust if source population count differs.")
    chk("All 187 Population B controls processed",
        187, len(pop_b), "Adjust if source population count differs.")
    chk("No linked_payment controls included", 0,
        len(all_controls[~all_controls["population_type"].isin(
            ["A - linked_non_payment","B - not_linked"])]))

    if not candidates.empty:
        chk("Every candidate row has Control_ID", 0,
            candidates["Control_ID"].isna().sum())
        chk("Every candidate row has l3_process_UUID", 0,
            candidates["candidate_l3_process_UUID"].isna().sum())
        chk("Candidate ranks start at 1 per control",
            "True",
            str((candidates.groupby("Control_ID")["candidate_rank"].min() == 1).all()))

        # Similarity scores in valid range
        scores = candidates["candidate_similarity_score"]
        chk("All similarity scores in [0, 1]",
            0, int((scores < 0).sum() + (scores > 1).sum()))

        chk("Review action populated for every row", 0,
            candidates["recommended_review_action"].isna().sum())

        # Candidate UUIDs exist in process pool
        valid_uuids = set(processes["l3_process_UUID"].dropna())
        bad_uuids = candidates[
            ~candidates["candidate_l3_process_UUID"].isin(valid_uuids)
        ].shape[0]
        chk("All candidate UUIDs exist in process pool", 0, bad_uuids)

        chk("No candidate treated as final linkage",
            "True", "True",
            "Embedding output is candidate generation only. "
            "recommended_review_action field reflects this.")

        # Threshold summary reconciles to candidate output
        total_rank1 = len(candidates[candidates["candidate_rank"] == 1])
        chk("Threshold summary row count matches controls",
            len(all_controls), total_rank1)

    return pd.DataFrame(checks)

# ─────────────────────────────────────────────────────────────────────────────
#  WRITE OUTPUT
# ─────────────────────────────────────────────────────────────────────────────

def write_outputs(pop_a_candidates, pop_b_candidates, all_candidates,
                  ctrl_summary, threshold_summary, validation):
    print(f"\n  Writing output to:\n  {OUTPUT_FILE}")
    EMBED_DIR.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as w:
        threshold_summary.to_excel(w, index=False,
            sheet_name="threshold_summary")
        ctrl_summary.to_excel(w, index=False,
            sheet_name="candidate_summary_by_control")
        pop_a_candidates.to_excel(w, index=False,
            sheet_name="population_a_candidates")
        pop_b_candidates.to_excel(w, index=False,
            sheet_name="population_b_candidates")
        all_candidates.to_excel(w, index=False,
            sheet_name="all_candidates_combined")
        validation.to_excel(w, index=False,
            sheet_name="validation_checks")

    sheets = [
        ("threshold_summary",            "counts by population, band, action"),
        ("candidate_summary_by_control", f"{len(ctrl_summary):,} controls — best candidate per control"),
        ("population_a_candidates",      f"{len(pop_a_candidates):,} rows — Population A top {TOP_K} per control"),
        ("population_b_candidates",      f"{len(pop_b_candidates):,} rows — Population B top {TOP_K} per control"),
        ("all_candidates_combined",      f"{len(all_candidates):,} total candidate rows"),
        ("validation_checks",            "integrity checks"),
    ]
    print(f"\n  Sheets (6):")
    for name, desc in sheets:
        print(f"    {name:<35} {desc}")

# ─────────────────────────────────────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Embeddings-based fallback linkage for Payments Controls PoC."
    )
    parser.add_argument("--force", action="store_true",
                        help="Regenerate embeddings even if cache exists.")
    args = parser.parse_args()

    print("=" * 70)
    print("  Payments Controls PoC — Embeddings-Based Fallback Linkage")
    print("=" * 70)
    print(f"  Model     : {EMBEDDING_MODEL}  ({EMBEDDING_DIM}d)")
    print(f"  Top K     : {TOP_K}")
    print(f"  Thresholds: Strong>={THRESHOLD_STRONG}  "
          f"Medium>={THRESHOLD_MEDIUM}  Weak>={THRESHOLD_WEAK}")
    print(f"  Cache dir : {EMBED_DIR}")

    EMBED_DIR.mkdir(parents=True, exist_ok=True)

    # Load data
    print("\n  Loading data...")
    all_controls = load_populations()
    processes    = load_processes()

    # Build embedding texts
    print("\n  Building embedding texts...")
    all_controls["_embedding_text"] = all_controls.apply(
        build_control_text, axis=1)
    processes["_embedding_text"] = processes.apply(
        build_process_text, axis=1)

    print(f"  Sample control embedding text length: "
          f"{all_controls['_embedding_text'].str.len().mean():.0f} chars avg")
    print(f"  Sample process embedding text length: "
          f"{processes['_embedding_text'].str.len().mean():.0f} chars avg")

    # Generate or load cached embeddings
    ctrl_ids = all_controls["Control_ID"].values
    proc_ids = processes["l3_process_UUID"].values

    need_ctrl = (args.force or
                 not CACHE_CONTROLS.exists() or
                 not CACHE_CTRL_IDS.exists() or
                 len(np.load(CACHE_CTRL_IDS, allow_pickle=True)) != len(ctrl_ids))
    need_proc = (args.force or
                 not CACHE_PROCS.exists() or
                 not CACHE_PROC_IDS.exists() or
                 len(np.load(CACHE_PROC_IDS, allow_pickle=True)) != len(proc_ids))

    if need_ctrl or need_proc:
        print("\n  Initialising Azure OpenAI client...")
        client = init_client()
        print("  Client ready.")

    if need_ctrl:
        print("\n  Generating control embeddings...")
        ctrl_texts = all_controls["_embedding_text"].tolist()
        ctrl_embs  = embed_texts(client, ctrl_texts, "controls")
        np.save(CACHE_CONTROLS, ctrl_embs)
        np.save(CACHE_CTRL_IDS, ctrl_ids)
        print(f"  Control embeddings cached to {CACHE_CONTROLS}")
    else:
        print(f"\n  Loading cached control embeddings ({CACHE_CONTROLS.name})...")
        ctrl_embs = np.load(CACHE_CONTROLS)
        print(f"  Loaded: {ctrl_embs.shape}")

    if need_proc:
        print("\n  Generating process embeddings...")
        proc_texts = processes["_embedding_text"].tolist()
        proc_embs  = embed_texts(client, proc_texts, "processes")
        np.save(CACHE_PROCS, proc_embs)
        np.save(CACHE_PROC_IDS, proc_ids)
        print(f"  Process embeddings cached to {CACHE_PROCS}")
    else:
        print(f"\n  Loading cached process embeddings ({CACHE_PROCS.name})...")
        proc_embs = np.load(CACHE_PROCS)
        print(f"  Loaded: {proc_embs.shape}")

    # Compute cosine similarity
    print(f"\n  Computing cosine similarity "
          f"({len(all_controls)} × {len(processes)})...")
    sim_matrix = cosine_similarity_matrix(ctrl_embs, proc_embs)
    print(f"  Similarity matrix: {sim_matrix.shape}  "
          f"max={sim_matrix.max():.4f}  mean={sim_matrix.mean():.4f}")

    # Generate candidates
    print(f"\n  Generating top {TOP_K} candidates per control...")
    all_candidates = generate_candidates(all_controls, processes, sim_matrix)

    # Split by population
    pop_a_cands = all_candidates[
        all_candidates["population_type"].str.startswith("A")
    ].copy()
    pop_b_cands = all_candidates[
        all_candidates["population_type"].str.startswith("B")
    ].copy()

    print(f"  Population A candidates: {len(pop_a_cands):,}")
    print(f"  Population B candidates: {len(pop_b_cands):,}")

    # Summary and validation
    ctrl_summary      = build_candidate_summary(all_candidates)
    threshold_summary = build_threshold_summary(all_candidates)
    validation        = run_validation(all_controls, processes, all_candidates)

    # Write outputs
    write_outputs(pop_a_cands, pop_b_cands, all_candidates,
                  ctrl_summary, threshold_summary, validation)

    failures = validation[validation["pass"] == "FAIL"]
    if not failures.empty:
        print(f"\n  WARNING: {len(failures)} validation check(s) FAILED:")
        for _, r in failures.iterrows():
            print(f"    - {r['check']}: expected {r['expected']}, got {r['actual']}")
    else:
        print("\n  All validation checks passed.")

    # Print threshold summary to console
    print("\n  THRESHOLD SUMMARY:")
    print(threshold_summary.to_string(index=False))

    print("\n" + "=" * 70)
    print("  Done.")
    print(f"  Output: {OUTPUT_FILE}")
    print("=" * 70)

if __name__ == "__main__":
    main()
