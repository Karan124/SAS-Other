"""
build_linkage_candidates.py
─────────────────────────────────────────────────────────────────────────────
Payments Controls PoC — Multi-Signal Linkage Candidate Generation
Stage 1 of 2

Generates candidate control-to-process linkages for:
  Population A: 141 controls linked to non-payment Holo processes only
  Population B: 187 controls with no Holo linkage

Three signals combined into a composite score per control-process pair:
  1. Application matching  — Alfabet apps in control text vs process Alfabet apps
  2. Product matching      — Product terms in control text vs process products
  3. Semantic similarity   — Text embedding cosine similarity (enriched with apps/products)
  4. CT alignment          — Rule-based CT outcome vs process lifecycle stage

Outputs top 15 candidates per control for LLM reranking in Stage 2.

Before running (PowerShell):
  az account set --subscription 6c72e6c5-ed48-4030-b29c-34e2849c9288
  $env:REQUESTS_CA_BUNDLE = "C:\\path\\to\\cacert.pem"
  $env:SSL_CERT_FILE      = "C:\\path\\to\\cacert.pem"
  Remove-Item Env:AZURE_CA_BUNDLE -ErrorAction SilentlyContinue
  python build_linkage_candidates.py
  python build_linkage_candidates.py --force   # regenerate embeddings
"""

import argparse
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

AZURE_ENDPOINT   = "https://ai.eng.azure.srv.westpac.com.au"
API_VERSION      = "2024-10-21"
EMBEDDING_MODEL  = "text-embedding-3-large"
EMBEDDING_DIM    = 3072
MAX_CHAR         = 24000   # ~8000 tokens safety margin

BASE_DIR         = Path(r"C:\Users\m061400\ai-test\big_table")
OUTPUT_DIR       = BASE_DIR / "multi_signal"

FILE_CONTROLS    = BASE_DIR / "juno_payment_controls_gold.xlsx"
FILE_LINKAGE     = BASE_DIR / "juno_holo_deterministic_linkage.xlsx"
FILE_PROCESSES   = BASE_DIR / "holocentric_payment_processes.xlsx"
FILE_PRODUCTS    = BASE_DIR / "products" / "holocentric_products_complete.xlsx"
FILE_ALFABET     = BASE_DIR / "Application_Extract.xlsx"
FILE_LINKAGE_WB  = BASE_DIR / "phase1e_outputs" / "payment_control_linkage_analysis.xlsx"

CACHE_CTRL_EMBS  = OUTPUT_DIR / "cache_ctrl_embs.npy"
CACHE_CTRL_IDS   = OUTPUT_DIR / "cache_ctrl_ids.npy"
CACHE_PROC_EMBS  = OUTPUT_DIR / "cache_proc_embs.npy"
CACHE_PROC_IDS   = OUTPUT_DIR / "cache_proc_ids.npy"

OUTPUT_FILE      = OUTPUT_DIR / "candidates_top15.xlsx"

TOP_K            = 15    # candidates per control for LLM reranking
EMBED_BATCH      = 100   # texts per embedding API call

# Composite score weights — adjusted dynamically when signals are absent
W_APP     = 0.45
W_PRODUCT = 0.25
W_SEMANTIC= 0.20
W_CT      = 0.10

# CT-to-lifecycle alignment map (specific CTs only; broad CTs score 0.5 always)
CT_LIFECYCLE_EXPECTED = {
    "CT1":  ["Initiation & Validation & Authorisation"],
    "CT6":  ["Initiation & Validation & Authorisation"],
    "CT3":  ["Execution & Early Processing Assurance"],
    "CT4":  ["Execution & Early Processing Assurance"],
    "CT14": ["Clearing / Settlement"],
    "CT2":  ["Posting & Accounting, Detection"],
    "CT27": ["Posting & Accounting, Detection"],
    "CT22": ["Notification & Reporting"],
    "CT23": ["Notification & Reporting"],
    "CT5":  ["Incident response, disputes, recovery followups"],
    "CT24": ["Incident response, disputes, recovery followups"],
    "CT28": ["Incident response, disputes, recovery followups"],
}
CT_BROAD = {f"CT{i}" for i in [7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,25,26]}

# Control text fields to concatenate for app/product/embedding search
CTRL_TEXT_FIELDS = [
    "CTRL_NAME", "CTRL_DESC", "CTRL_DESC_OF_CTRL",
    "CTRL_MNTRD", "CTRL_EVDNCD", "CTRL_CTGRY_1", "CTRL_CTGRY_2",
    "CTRL_CATEGORY", "CTRL_LOC_LIST",
]

# ─────────────────────────────────────────────────────────────────────────────
#  HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def clean_cols(df):
    df.columns = df.columns.str.strip()
    return df

def safe_str(val):
    if pd.isna(val):
        return ""
    s = str(val).replace("\t"," ").replace("\r"," ").replace("\n"," ")
    return re.sub(r" {2,}", " ", s).strip()

def is_empty(val):
    return not safe_str(val)

def norm_gold_ctrl(val):
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

def jaccard(set_a, set_b):
    """Jaccard similarity between two sets. Returns 0 if both empty."""
    if not set_a and not set_b:
        return 0.0
    if not set_a or not set_b:
        return 0.0
    inter = len(set_a & set_b)
    union = len(set_a | set_b)
    return inter / union if union else 0.0

def ct_alignment(ct_code, lifecycle_stage):
    """
    1.0 if CT is expected at this lifecycle stage,
    0.5 if CT is a broad system/governance control,
    0.0 if CT is specific but unexpected at this stage.
    """
    if not ct_code:
        return 0.0
    if ct_code in CT_BROAD:
        return 0.5
    expected = CT_LIFECYCLE_EXPECTED.get(ct_code, [])
    if not expected:
        return 0.0
    if lifecycle_stage in expected:
        return 1.0
    # CT14 partial credit at non-clearing stages (dual role)
    if ct_code == "CT14":
        return 0.5
    return 0.0

def compute_composite(app_s, prod_s, sem_s, ct_s,
                      has_ctrl_apps, has_proc_apps, has_proc_prods):
    """
    Composite score with dynamic weight redistribution when signals are absent.
    """
    app_avail  = has_ctrl_apps and has_proc_apps
    prod_avail = has_proc_prods

    if app_avail and prod_avail:
        w = (W_APP, W_PRODUCT, W_SEMANTIC, W_CT)
    elif app_avail and not prod_avail:
        # No product signal — boost app and semantic
        w = (W_APP + 0.12, 0, W_SEMANTIC + 0.08, W_CT)
    elif not app_avail and prod_avail:
        # No app signal — boost product and semantic
        w = (0, W_PRODUCT + 0.22, W_SEMANTIC + 0.18, W_CT)
    else:
        # Neither app nor product — rely on semantic + CT
        w = (0, 0, W_SEMANTIC + W_APP + W_PRODUCT, W_CT)

    return w[0]*app_s + w[1]*prod_s + w[2]*sem_s + w[3]*ct_s

# ─────────────────────────────────────────────────────────────────────────────
#  APP VOCABULARY BUILDER
# ─────────────────────────────────────────────────────────────────────────────

def build_app_vocabulary(alfabet_df):
    """
    Build a search vocabulary from the Alfabet master app list.
    For each app name, extract:
      - Full name (lowercase, stripped)
      - Parenthetical content if present
      - Acronyms (2+ uppercase letter sequences)
      - Short tokens (words ≥4 chars that aren't common English words)
    Returns: list of (canonical_name, set_of_search_terms)
    """
    # Find the application name column (flexible)
    name_col = next(
        (c for c in alfabet_df.columns
         if "application" in c.lower() and "name" in c.lower()),
        alfabet_df.columns[0]
    )

    vocab = []  # list of (canonical, {term1, term2, ...})
    SKIP_WORDS = {
        "system","platform","service","and","for","the","on","of","in",
        "with","data","base","online","group","bank","west","pac","wbc",
    }

    for raw in alfabet_df[name_col].dropna().unique():
        raw = str(raw).strip()
        if not raw:
            continue

        terms = set()
        canonical = raw

        # Full name (lowercase, no special chars)
        full_lower = raw.lower()
        terms.add(full_lower)

        # Parenthetical content
        parens = re.findall(r'\(([^)]+)\)', raw)
        for p in parens:
            for tok in p.split("/"):
                tok = tok.strip()
                if len(tok) >= 2:
                    terms.add(tok.lower())

        # Acronyms: sequences of 2+ uppercase letters
        acronyms = re.findall(r'\b[A-Z]{2,}\b', raw)
        terms.update(a.lower() for a in acronyms)

        # Significant words from the name (≥4 chars, not in skip list)
        words = re.sub(r'[^a-zA-Z0-9\s]', ' ', raw).split()
        for w in words:
            wl = w.lower()
            if len(wl) >= 4 and wl not in SKIP_WORDS:
                terms.add(wl)

        vocab.append((canonical, terms))

    print(f"  App vocabulary: {len(vocab):,} applications")
    return vocab


def find_app_mentions(ctrl_text, vocab):
    """
    Search a control's concatenated text for Alfabet application mentions.
    Returns set of canonical app names that appear in the text.
    Uses whole-word matching to avoid false positives.
    """
    text_lower = ctrl_text.lower()
    found = set()
    for canonical, terms in vocab:
        for term in terms:
            if len(term) < 3:
                continue
            # Whole-word search (term surrounded by non-alphanumeric chars or boundaries)
            pattern = r'(?<![a-z0-9])' + re.escape(term) + r'(?![a-z0-9])'
            if re.search(pattern, text_lower):
                found.add(canonical)
                break
    return found


def parse_process_apps(raw_app_str, vocab):
    """
    Parse semicolon-delimited Alfabet apps from a process record.
    Returns set of canonical app names (matched against vocab).
    """
    if is_empty(raw_app_str):
        return set()
    found = set()
    for part in str(raw_app_str).split(";"):
        part = part.strip()
        if not part:
            continue
        # Try to match against canonical names in vocab
        part_lower = part.lower()
        matched = False
        for canonical, terms in vocab:
            if part_lower == canonical.lower() or part_lower in terms:
                found.add(canonical)
                matched = True
                break
        if not matched:
            # No match in vocab — use the raw name as-is
            found.add(part)
    return found


def find_product_mentions(ctrl_text, product_catalogue):
    """
    Search control text for product names from the catalogue.
    Returns set of matched product names.
    """
    text_lower = ctrl_text.lower()
    found = set()
    for prod in product_catalogue:
        prod_lower = prod.lower()
        if len(prod_lower) < 4:
            continue
        pattern = r'(?<![a-z0-9])' + re.escape(prod_lower) + r'(?![a-z0-9])'
        if re.search(pattern, text_lower):
            found.add(prod)
    return found

# ─────────────────────────────────────────────────────────────────────────────
#  EMBEDDING HELPERS
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


def embed_texts(client, texts, label="texts"):
    all_embs = []
    total = len(texts)
    print(f"  Embedding {total:,} {label}...")
    for i in range(0, total, EMBED_BATCH):
        batch = [t[:MAX_CHAR] for t in texts[i:i+EMBED_BATCH]]
        for attempt in range(4):
            try:
                resp = client.embeddings.create(
                    model=EMBEDDING_MODEL, input=batch)
                all_embs.extend([d.embedding for d in resp.data])
                print(f"    {min(i+EMBED_BATCH, total):,}/{total:,}", end="\r")
                time.sleep(0.2)
                break
            except Exception as e:
                if attempt == 3:
                    raise
                time.sleep(5 * 2**attempt)
    print(f"    {total:,}/{total:,} done        ")
    return np.array(all_embs, dtype=np.float32)


def cosine_sim_matrix(a, b):
    a_n = a / (np.linalg.norm(a, axis=1, keepdims=True) + 1e-9)
    b_n = b / (np.linalg.norm(b, axis=1, keepdims=True) + 1e-9)
    return (a_n @ b_n.T).astype(np.float32)


def build_ctrl_embedding_text(row, detected_apps, detected_products):
    """Enriched control embedding text with detected apps and products."""
    parts = []
    name = safe_str(row.get("CTRL_NAME"))
    if name:
        parts.append(f"CONTROL NAME:\n{name}")
    desc = safe_str(row.get("CTRL_DESC"))
    if desc:
        parts.append(f"CONTROL DESCRIPTION:\n{desc}")
    desc2 = safe_str(row.get("CTRL_DESC_OF_CTRL"))
    if desc2 and desc2 != desc:
        parts.append(f"CONTROL ACTIVITY:\n{desc2}")
    gc = norm_gold_ctrl(row.get("gold_control"))
    if gc:
        from CT_TITLES_LOCAL import CT_TITLES
        title = CT_TITLES.get(gc, "")
        parts.append(f"CONTROL TYPE:\n{gc} — {title}" if title else f"CONTROL TYPE:\n{gc}")
    if detected_apps:
        parts.append(f"REFERENCED APPLICATIONS:\n{'; '.join(sorted(detected_apps))}")
    if detected_products:
        parts.append(f"REFERENCED PRODUCTS:\n{'; '.join(sorted(detected_products))}")
    monitored = safe_str(row.get("CTRL_MNTRD"))
    evidenced = safe_str(row.get("CTRL_EVDNCD"))
    if monitored:
        parts.append(f"HOW MONITORED:\n{monitored}")
    if evidenced:
        parts.append(f"HOW EVIDENCED:\n{evidenced}")
    return "\n\n".join(parts)


def build_proc_embedding_text(row):
    """Enriched process embedding text with apps and products."""
    parts = []
    l2_name = safe_str(row.get("l2_process_name"))
    l2_desc = safe_str(row.get("l2_process_description"))
    l3_name = safe_str(row.get("l3_activity_name"))
    l3_desc = safe_str(row.get("l3_activity_description"))
    if l2_name: parts.append(f"L2 PROCESS:\n{l2_name}")
    if l2_desc: parts.append(f"L2 DESCRIPTION:\n{l2_desc}")
    if l3_name: parts.append(f"L3 ACTIVITY:\n{l3_name}")
    if l3_desc: parts.append(f"L3 DESCRIPTION:\n{l3_desc}")
    tasks = safe_str(row.get("task_name"))
    if tasks:  parts.append(f"TASKS:\n{tasks}")
    apps = safe_str(row.get("alphabet_app"))
    if apps:
        # Strip BM codes from app names if present, else use as-is
        clean_apps = "; ".join(
            re.sub(r'^BM\s+[\d.]+\s*[-–—]\s*', '', a.strip())
            for a in apps.split(";") if a.strip()
        )
        parts.append(f"APPLICATIONS:\n{clean_apps}")
    prod = safe_str(row.get("final_product") or row.get("clean_product"))
    if prod: parts.append(f"PRODUCTS:\n{prod}")
    cat   = safe_str(row.get("process_category"))
    stage = safe_str(row.get("process_lifecycle_stage"))
    rat   = safe_str(row.get("payment_rationale"))
    if cat:   parts.append(f"PAYMENT CATEGORY:\n{cat}")
    if stage: parts.append(f"LIFECYCLE STAGE:\n{stage}")
    if rat:   parts.append(f"PAYMENT RELEVANCE:\n{rat}")
    return "\n\n".join(parts)

# ─────────────────────────────────────────────────────────────────────────────
#  LOAD DATA
# ─────────────────────────────────────────────────────────────────────────────

def load_data():
    print("\n  Loading input files...")

    controls = clean_cols(pd.read_excel(FILE_CONTROLS, dtype=str))
    controls["Control_ID"]       = controls["Control_ID"].str.strip()
    controls["gold_control_code"]= controls["gold_control"].apply(norm_gold_ctrl)
    print(f"  Controls            : {len(controls):,}")

    # Normalise text fields
    for col in CTRL_TEXT_FIELDS:
        if col in controls.columns:
            controls[col] = controls[col].apply(safe_str)

    # Population A/B IDs
    pop_a_ids, pop_b_ids = set(), set()
    if FILE_LINKAGE_WB.exists():
        pop_a_ids = set(clean_cols(
            pd.read_excel(FILE_LINKAGE_WB, sheet_name="linked_non_payment", dtype=str)
        )["Control_ID"].dropna().str.strip())
        pop_b_ids = set(clean_cols(
            pd.read_excel(FILE_LINKAGE_WB, sheet_name="not_linked", dtype=str)
        )["Control_ID"].dropna().str.strip())
    print(f"  Population A        : {len(pop_a_ids):,}")
    print(f"  Population B        : {len(pop_b_ids):,}")

    controls["population_type"] = controls["Control_ID"].apply(
        lambda x: ("A - linked_non_payment" if x in pop_a_ids
                   else "B - not_linked" if x in pop_b_ids
                   else "Other"))
    target_controls = controls[
        controls["population_type"].isin(["A - linked_non_payment","B - not_linked"])
    ].drop_duplicates(subset=["Control_ID"]).copy()
    print(f"  Target controls     : {len(target_controls):,}")

    # Processes
    processes = clean_cols(pd.read_excel(FILE_PROCESSES, dtype=str))
    prod_col = next((c for c in processes.columns
                     if "product" in c.lower() and "service" in c.lower()), None)
    if prod_col and prod_col != "l3_activity_product_service":
        processes = processes.rename(columns={prod_col: "l3_activity_product_service"})
    processes["l3_process_UUID"] = processes["l3_process_UUID"].apply(
        lambda x: safe_str(x))
    processes = processes.dropna(subset=["l3_process_UUID"])
    processes = processes[processes["l3_process_UUID"] != ""]
    processes["process_lifecycle_stage"] = processes["process_lifecycle_stage"].apply(
        safe_str)

    # Merge in products
    if FILE_PRODUCTS.exists():
        prods_df = clean_cols(pd.read_excel(FILE_PRODUCTS, dtype=str))
        if "l3_process_UUID" in prods_df.columns:
            prods_df["l3_process_UUID"] = prods_df["l3_process_UUID"].apply(safe_str)
            prod_cols = [c for c in ["final_product","clean_product"]
                         if c in prods_df.columns]
            if prod_cols:
                processes = processes.merge(
                    prods_df[["l3_process_UUID"] + prod_cols],
                    on="l3_process_UUID", how="left"
                )
    print(f"  Processes           : {len(processes):,}")

    # Alfabet app list
    alfabet = clean_cols(pd.read_excel(FILE_ALFABET, dtype=str))
    print(f"  Alfabet apps        : {len(alfabet):,}")

    return target_controls, processes, alfabet

# ─────────────────────────────────────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--force", action="store_true",
                        help="Regenerate embeddings even if cache exists.")
    args = parser.parse_args()

    print("=" * 70)
    print("  Multi-Signal Linkage Candidate Generation")
    print(f"  Embedding model : {EMBEDDING_MODEL}")
    print(f"  Top K candidates: {TOP_K}")
    print("=" * 70)

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    # ── Load ─────────────────────────────────────────────────────────────────
    target_controls, processes, alfabet = load_data()

    # ── Build app vocabulary ──────────────────────────────────────────────────
    print("\n  Building Alfabet app vocabulary...")
    app_vocab = build_app_vocabulary(alfabet)

    # ── Build product catalogue from process data ─────────────────────────────
    print("\n  Building product catalogue from process data...")
    prod_values = []
    for col in ["final_product","clean_product","l3_activity_product_service"]:
        if col in processes.columns:
            for val in processes[col].dropna():
                for p in str(val).split(","):
                    p = p.strip()
                    if p and len(p) >= 3:
                        prod_values.append(p)
            break
    product_catalogue = sorted(set(p for p in prod_values if p))
    print(f"  Product catalogue   : {len(product_catalogue):,} terms")

    # ── Extract signals from controls ──────────────────────────────────────────
    print("\n  Extracting app and product signals from control text...")
    ctrl_concat_text = {}
    ctrl_apps        = {}
    ctrl_products    = {}

    for _, row in target_controls.iterrows():
        cid = row["Control_ID"]
        text_parts = [safe_str(row.get(f,"")) for f in CTRL_TEXT_FIELDS
                      if f in row.index]
        full_text = " ".join(text_parts)
        ctrl_concat_text[cid] = full_text
        ctrl_apps[cid]        = find_app_mentions(full_text, app_vocab)
        ctrl_products[cid]    = find_product_mentions(full_text, product_catalogue)

    n_with_apps = sum(1 for a in ctrl_apps.values() if a)
    n_with_prods = sum(1 for p in ctrl_products.values() if p)
    print(f"  Controls with app matches     : {n_with_apps:,}/{len(target_controls):,}")
    print(f"  Controls with product matches : {n_with_prods:,}/{len(target_controls):,}")

    # App signal sample
    for cid, apps in list(ctrl_apps.items())[:3]:
        if apps:
            print(f"    {cid}: {list(apps)[:3]}")

    # ── Parse process apps and products ───────────────────────────────────────
    print("\n  Parsing process-side signals...")
    proc_apps     = {}
    proc_products = {}
    for _, row in processes.iterrows():
        uuid = row["l3_process_UUID"]
        proc_apps[uuid] = parse_process_apps(row.get("alphabet_app",""), app_vocab)
        prod_val = safe_str(row.get("final_product","") or row.get("clean_product",""))
        proc_products[uuid] = {p.strip() for p in prod_val.split(",")
                               if len(p.strip()) >= 3} if prod_val else set()

    n_proc_apps  = sum(1 for a in proc_apps.values() if a)
    n_proc_prods = sum(1 for p in proc_products.values() if p)
    print(f"  Processes with apps     : {n_proc_apps:,}/{len(processes):,}")
    print(f"  Processes with products : {n_proc_prods:,}/{len(processes):,}")

    # ── Build embedding texts ─────────────────────────────────────────────────
    print("\n  Building embedding texts...")

    # Temporarily inject CT_TITLES into a pseudo-module for the helper function
    import sys, types
    ct_mod = types.ModuleType("CT_TITLES_LOCAL")
    ct_mod.CT_TITLES = {
        "CT1":"Validation of Human-Entered Data at Input",
        "CT2":"Payment processing error detection",
        "CT3":"Early Identification of Duplications and Processing Errors",
        "CT4":"Payment processing interface and batch error resolution",
        "CT5":"Incident response","CT6":"Master/Reference data input validation",
        "CT7":"Service provider ongoing review","CT8":"Service provider onboarding",
        "CT9":"Change management testing","CT10":"Critical service chain mapping",
        "CT11":"Rollback plans","CT12":"System recovery capability",
        "CT13":"Business continuity plan","CT14":"Logging and monitoring",
        "CT15":"Secure IT Design","CT16":"Vulnerability Management",
        "CT17":"Access - Provision / deprovision","CT18":"Access - Monitoring",
        "CT19":"Patch Management","CT20":"Physical Security Controls",
        "CT21":"Access - Privileged users",
        "CT22":"Mistaken Internet Payment Reports",
        "CT23":"Provision of Confirmations and Notifications",
        "CT24":"Treatment of Unauthorised and Disputed Transactions",
        "CT25":"Provision of Confirmations and Notifications",
        "CT26":"Regulatory horizon scanning","CT27":"Records retention",
        "CT28":"Crisis Management planning and testing",
    }
    sys.modules["CT_TITLES_LOCAL"] = ct_mod

    ctrl_ids   = target_controls["Control_ID"].tolist()
    ctrl_texts = [
        build_ctrl_embedding_text(
            row, ctrl_apps[row["Control_ID"]], ctrl_products[row["Control_ID"]])
        for _, row in target_controls.iterrows()
    ]

    proc_ids   = processes["l3_process_UUID"].tolist()
    proc_texts = [build_proc_embedding_text(row)
                  for _, row in processes.iterrows()]

    # ── Generate or load embeddings ───────────────────────────────────────────
    need_ctrl = (args.force or
                 not CACHE_CTRL_EMBS.exists() or
                 len(np.load(CACHE_CTRL_IDS, allow_pickle=True)) != len(ctrl_ids))
    need_proc = (args.force or
                 not CACHE_PROC_EMBS.exists() or
                 len(np.load(CACHE_PROC_IDS, allow_pickle=True)) != len(proc_ids))

    if need_ctrl or need_proc:
        print("\n  Initialising Azure OpenAI client...")
        client = init_client()
        print("  Client ready.")

    if need_ctrl:
        print("\n  Generating control embeddings...")
        ctrl_embs = embed_texts(client, ctrl_texts, "controls")
        np.save(CACHE_CTRL_EMBS, ctrl_embs)
        np.save(CACHE_CTRL_IDS, np.array(ctrl_ids))
        print(f"  Cached: {CACHE_CTRL_EMBS.name}")
    else:
        print(f"\n  Loading cached control embeddings...")
        ctrl_embs = np.load(CACHE_CTRL_EMBS)

    if need_proc:
        print("\n  Generating process embeddings...")
        proc_embs = embed_texts(client, proc_texts, "processes")
        np.save(CACHE_PROC_EMBS, proc_embs)
        np.save(CACHE_PROC_IDS, np.array(proc_ids))
        print(f"  Cached: {CACHE_PROC_EMBS.name}")
    else:
        print(f"\n  Loading cached process embeddings...")
        proc_embs = np.load(CACHE_PROC_EMBS)

    print(f"  Control embs : {ctrl_embs.shape}")
    print(f"  Process embs : {proc_embs.shape}")

    # ── Compute cosine similarity ─────────────────────────────────────────────
    print(f"\n  Computing cosine similarity "
          f"({len(ctrl_ids)} × {len(proc_ids)})...")
    sim_matrix = cosine_sim_matrix(ctrl_embs, proc_embs)
    print(f"  Similarity matrix: {sim_matrix.shape}  "
          f"max={sim_matrix.max():.4f}  mean={sim_matrix.mean():.4f}")

    # ── Build candidate table ─────────────────────────────────────────────────
    print(f"\n  Building top-{TOP_K} candidate pairs per control...")

    # Process metadata lookup
    proc_meta = processes.set_index("l3_process_UUID")[[
        "l3_activity_name","l2_process_name","l2_process_id",
        "l3_activity_description","l3_activity_id",
        "process_category","process_lifecycle_stage",
        "alphabet_app","value_stream_name","payment_rationale",
    ] + (["final_product"] if "final_product" in processes.columns else
         ["clean_product"] if "clean_product" in processes.columns else [])
    ].to_dict("index")

    ctrl_meta = target_controls.set_index("Control_ID")[[
        "CTRL_NAME","gold_control","gold_control_code","population_type",
        "CTRL_TYP","CTRL_ASSESS_RTNG",
    ]].to_dict("index")

    rows = []
    for ci, cid in enumerate(ctrl_ids):
        sims = sim_matrix[ci]
        top_idx = np.argsort(sims)[::-1][:TOP_K]

        c_apps    = ctrl_apps.get(cid, set())
        c_prods   = ctrl_products.get(cid, set())
        ct_code   = ctrl_meta.get(cid, {}).get("gold_control_code","")
        pop_type  = ctrl_meta.get(cid, {}).get("population_type","")

        for rank, pidx in enumerate(top_idx, 1):
            uuid     = proc_ids[pidx]
            sem_s    = float(sims[pidx])
            p_meta   = proc_meta.get(uuid, {})
            p_apps   = proc_apps.get(uuid, set())
            p_prods  = proc_products.get(uuid, set())
            stage    = p_meta.get("process_lifecycle_stage","")

            app_s  = jaccard(c_apps, p_apps)
            prod_s = jaccard(c_prods, p_prods)
            ct_s   = ct_alignment(ct_code, stage)
            comp_s = compute_composite(
                app_s, prod_s, sem_s, ct_s,
                bool(c_apps), bool(p_apps), bool(p_prods)
            )

            matched_apps  = sorted(c_apps & p_apps)
            matched_prods = sorted(c_prods & p_prods)

            prod_val = safe_str(p_meta.get("final_product","") or
                                p_meta.get("clean_product",""))

            rows.append({
                "population_type":           pop_type,
                "Control_ID":                cid,
                "CTRL_NAME":                 ctrl_meta.get(cid,{}).get("CTRL_NAME",""),
                "gold_control_code":         ct_code,
                "l3_process_UUID":           uuid,
                "l3_activity_id":            p_meta.get("l3_activity_id",""),
                "l3_activity_name":          p_meta.get("l3_activity_name",""),
                "l3_activity_description":   p_meta.get("l3_activity_description",""),
                "l2_process_id":             p_meta.get("l2_process_id",""),
                "l2_process_name":           p_meta.get("l2_process_name",""),
                "process_category":          p_meta.get("process_category",""),
                "process_lifecycle_stage":   stage,
                "process_products":          prod_val,
                "process_apps":              safe_str(p_meta.get("alphabet_app","")),
                "candidate_rank":            rank,
                "composite_score":           round(comp_s, 6),
                "semantic_score":            round(sem_s, 6),
                "app_score":                 round(app_s, 6),
                "product_score":             round(prod_s, 6),
                "ct_alignment_score":        round(ct_s, 6),
                "matched_apps":              "; ".join(matched_apps),
                "matched_products":          "; ".join(matched_prods),
                "ctrl_detected_apps":        "; ".join(sorted(c_apps)),
                "ctrl_detected_products":    "; ".join(sorted(c_prods)),
                "app_match_fired":           bool(matched_apps),
                "product_match_fired":       bool(matched_prods),
            })

        if (ci + 1) % 50 == 0 or (ci + 1) == len(ctrl_ids):
            print(f"    {ci+1:,}/{len(ctrl_ids):,} controls processed", end="\r")

    print(f"    {len(ctrl_ids):,}/{len(ctrl_ids):,} controls processed    ")
    candidates_df = pd.DataFrame(rows)
    print(f"  Total candidate pairs: {len(candidates_df):,}")

    # ── Build signal summary (one row per control) ────────────────────────────
    summary_rows = []
    for cid in ctrl_ids:
        ctrl_cands = candidates_df[candidates_df["Control_ID"] == cid]
        best = ctrl_cands.iloc[0] if not ctrl_cands.empty else {}
        summary_rows.append({
            "Control_ID":          cid,
            "CTRL_NAME":           ctrl_meta.get(cid,{}).get("CTRL_NAME",""),
            "population_type":     ctrl_meta.get(cid,{}).get("population_type",""),
            "gold_control_code":   ctrl_meta.get(cid,{}).get("gold_control_code",""),
            "ctrl_detected_apps":  "; ".join(sorted(ctrl_apps.get(cid,set()))),
            "ctrl_detected_prods": "; ".join(sorted(ctrl_products.get(cid,set()))),
            "has_app_signal":      bool(ctrl_apps.get(cid)),
            "has_product_signal":  bool(ctrl_products.get(cid)),
            "best_composite_score": round(float(ctrl_cands["composite_score"].max()), 6)
                                    if not ctrl_cands.empty else 0,
            "best_semantic_score": round(float(ctrl_cands["semantic_score"].max()), 6)
                                   if not ctrl_cands.empty else 0,
            "top_app_match_fired": bool(
                ctrl_cands.iloc[0]["app_match_fired"] if not ctrl_cands.empty else False),
            "top_candidate_process": str(best.get("l3_activity_name","")) if best.get("l3_activity_name") else "",
            "top_candidate_category": str(best.get("process_category","")) if best.get("process_category") else "",
        })
    signal_summary = pd.DataFrame(summary_rows)

    # ── Validation ────────────────────────────────────────────────────────────
    checks = []
    def chk(name, expected, actual, note=""):
        passed = str(expected) == str(actual)
        checks.append({"check":name,"expected":str(expected),
                        "actual":str(actual),"pass":"PASS" if passed else "FAIL","note":note})
        print(f"  [{'PASS' if passed else 'FAIL'}] {name}")

    print("\n  Validation checks...")
    chk("Population A + B controls processed", len(target_controls), len(ctrl_ids))
    chk("Candidate rows = controls × top_k", len(target_controls)*TOP_K, len(rows))
    chk("No null l3_process_UUID in candidates", 0,
        candidates_df["l3_process_UUID"].isna().sum())
    chk("Composite scores in [0, 1]", 0,
        int(((candidates_df["composite_score"] < 0) |
             (candidates_df["composite_score"] > 1)).sum()))
    chk("Semantic scores in [0, 1]", 0,
        int(((candidates_df["semantic_score"] < 0) |
             (candidates_df["semantic_score"] > 1)).sum()))

    # ── Write outputs ─────────────────────────────────────────────────────────
    print(f"\n  Writing to {OUTPUT_FILE}...")
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as w:
        candidates_df.to_excel(w, index=False, sheet_name="candidate_pairs")
        signal_summary.to_excel(w, index=False, sheet_name="signal_summary")
        pd.DataFrame(checks).to_excel(w, index=False, sheet_name="validation_checks")

        # Quick summary
        agg_df = pd.DataFrame([
            ("Target controls", len(target_controls)),
            ("Population A", (signal_summary["population_type"]
                              == "A - linked_non_payment").sum()),
            ("Population B", (signal_summary["population_type"]
                              == "B - not_linked").sum()),
            ("Total candidate pairs", len(candidates_df)),
            ("Controls with app signal", int(signal_summary["has_app_signal"].sum())),
            ("Controls with product signal", int(signal_summary["has_product_signal"].sum())),
            ("Candidates with app match fired", int(candidates_df["app_match_fired"].sum())),
            ("Avg best composite score", f"{signal_summary['best_composite_score'].mean():.4f}"),
            ("Avg best semantic score",  f"{signal_summary['best_semantic_score'].mean():.4f}"),
        ], columns=["Metric","Value"])
        agg_df.to_excel(w, index=False, sheet_name="summary")

    print(f"\n  Output : {OUTPUT_FILE}")
    print(f"  Sheets : candidate_pairs | signal_summary | validation_checks | summary")
    print("\n  Stage 1 complete. Run run_linkage_llm_reranker.py for Stage 2.")
    print("=" * 70)


if __name__ == "__main__":
    main()
