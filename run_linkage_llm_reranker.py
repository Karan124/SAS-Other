"""
run_linkage_llm_reranker.py
─────────────────────────────────────────────────────────────────────────────
Payments Controls PoC — Multi-Signal Linkage LLM Reranker
Stage 2 of 2

Reads the top-15 candidates per control from Stage 1 and uses the LLM to
determine which candidates represent defensible control-to-process linkages.

The LLM has full context:
  - Complete control description and all signals
  - Each candidate's process name, description, Alfabet apps, products
  - Which signals fired (app match, product match, composite score)
  - CT taxonomy alignment

Output distinguishes linkage confidence levels and always requires SME review.

Before running (PowerShell):
  az account set --subscription 6c72e6c5-ed48-4030-b29c-34e2849c9288
  $env:REQUESTS_CA_BUNDLE = "C:\\path\\to\\cacert.pem"
  $env:SSL_CERT_FILE      = "C:\\path\\to\\cacert.pem"
  Remove-Item Env:AZURE_CA_BUNDLE -ErrorAction SilentlyContinue
  python run_linkage_llm_reranker.py
  python run_linkage_llm_reranker.py --force
"""

import argparse
import json
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

AZURE_ENDPOINT        = "https://ai.eng.azure.srv.westpac.com.au"
API_VERSION           = "2024-10-21"
MODEL                 = "gpt-5.4"
REASONING_EFFORT      = "medium"
MAX_COMPLETION_TOKENS = 8000

BASE_DIR     = Path(r"C:\Users\m061400\ai-test\big_table")
MULTI_DIR    = BASE_DIR / "multi_signal"

CANDIDATES_FILE  = MULTI_DIR / "candidates_top15.xlsx"
CONTROLS_FILE    = BASE_DIR / "juno_payment_controls_gold.xlsx"
LINKAGE_WB       = BASE_DIR / "phase1e_outputs" / "payment_control_linkage_analysis.xlsx"
CHECKPOINT       = MULTI_DIR / "llm_reranker_checkpoint.jsonl"
OUTPUT_FILE      = MULTI_DIR / "linkage_recommendations.xlsx"

RETRY_COUNT   = 3
RETRY_BASE    = 5
SLEEP_BETWEEN = 0.5

INPUT_PRICE_USD_PER_M  = 1.75
OUTPUT_PRICE_USD_PER_M = 14.00
AUD_USD_RATE           = 0.65

# ─────────────────────────────────────────────────────────────────────────────
#  PROMPTS
# ─────────────────────────────────────────────────────────────────────────────

SYSTEM_PROMPT = """
You are a payment risk control-to-process mapping analyst at an Australian bank.

Your task is to determine which candidate Holocentric payment processes a given
JUNO control most plausibly governs, based on ALL available evidence.

You must be conservative. It is better to return fewer high-confidence
recommendations than many weak ones. Reviewers will use your output to decide
where to invest SME validation time.

Return valid JSON only. No markdown. No text outside the JSON object.
""".strip()

PROMPT_TEMPLATE = """
CONTROL DETAILS
───────────────────────────────────────────────────────────────────────

Control ID      : {ctrl_id}
Control Name    : {ctrl_name}
CT Outcome      : {ct_code} — {ct_title}
Population      : {population_type}

Control Description:
{ctrl_desc}

Control Activity:
{ctrl_activity}

How Monitored:
{ctrl_monitored}

How Evidenced:
{ctrl_evidenced}

Applications detected in control text (from Alfabet master list):
{ctrl_apps}

Products detected in control text:
{ctrl_products}

───────────────────────────────────────────────────────────────────────
CANDIDATE PAYMENT PROCESSES (ranked by composite score)
───────────────────────────────────────────────────────────────────────

{candidates_block}

───────────────────────────────────────────────────────────────────────
DECISION CRITERIA — apply in order
───────────────────────────────────────────────────────────────────────

1. APPLICATION MATCH (strongest signal)
   The control text mentions specific applications from the Alfabet master list.
   If a candidate process runs on any of those same applications, this is strong
   structural evidence of linkage. Prioritise candidates where app_match_fired = true.

2. PRODUCT MATCH (strong signal)
   The control text references specific products (e.g., Home Loans, BPAY, ORMB).
   If a candidate process handles the same products, this strongly supports linkage.

3. ACTIVITY ALIGNMENT (medium signal)
   Does the control's risk domain (what it monitors, validates, or reconciles)
   align with what the candidate process actually does? Look for operational overlap.
   A control about settlement reconciliation aligns with a settlement process.
   A control about billing aligns with billing and invoicing processes.

4. CT TAXONOMY ALIGNMENT (supporting signal)
   The control's CT outcome (e.g., CT1 = data input validation) should align with
   the process lifecycle stage. Use ct_alignment_score as supporting evidence.

───────────────────────────────────────────────────────────────────────
CONFIDENCE LEVELS
───────────────────────────────────────────────────────────────────────

High   — At least one application OR product match confirmed AND
         activity alignment is clear.
Medium — Product or application match confirmed but not both, OR strong
         activity alignment without explicit app/product match.
Low    — Only semantic similarity and general activity alignment. No
         explicit app or product match.

NOT RECOMMENDED — When there is no defensible evidence of linkage.
Do not recommend a process just because it scored highest in composite score
if the underlying evidence is weak.

───────────────────────────────────────────────────────────────────────
OUTPUT FORMAT
───────────────────────────────────────────────────────────────────────

Return a single JSON object. No markdown. No text outside the JSON.

{{
  "control_id": "{ctrl_id}",
  "recommendations": [
    {{
      "l3_process_UUID": "<exact UUID from candidate list>",
      "l3_activity_name": "<exact name from candidate list>",
      "recommended_for_linkage": true,
      "confidence": "High | Medium | Low",
      "primary_signal": "app_match | product_match | activity_alignment | composite_only",
      "rationale": "<2-3 sentences citing specific evidence from the control text
                    and the process details that support this recommendation>",
      "requires_sme_review": true
    }}
  ],
  "no_linkage_recommended": false,
  "no_linkage_reason": "<explain if no_linkage_recommended is true>"
}}

RULES:
  - Only include candidates you are recommending (recommended_for_linkage = true).
  - If no candidates are defensible, return an empty recommendations array and
    set no_linkage_recommended = true with a reason.
  - requires_sme_review must always be true.
  - l3_process_UUID and l3_activity_name must be copied exactly from the candidate list.
  - Prioritise app_match and product_match candidates over semantic-only candidates.
  - A single control can link to multiple processes (multi-process linkage is valid).
""".strip()

# ─────────────────────────────────────────────────────────────────────────────
#  HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def safe_str(val, default=""):
    if pd.isna(val) if hasattr(pd, 'isna') else val is None:
        return default
    return str(val).strip()

def norm_gold_ctrl(val):
    if not val or str(val).strip() in ("","nan","None"):
        return ""
    s = str(val).strip()
    if s.upper().startswith("CT"):
        try:
            n = int(s[2:])
            return f"CT{n}" if 1 <= n <= 28 else ""
        except ValueError:
            return ""
    try:
        n = int(float(s))
        return f"CT{n}" if 1 <= n <= 28 else ""
    except (ValueError, TypeError):
        return ""

CT_TITLES = {
    "CT1":"Validation of Human-Entered Data at Input",
    "CT2":"Payment processing error detection",
    "CT3":"Early Identification of Duplications and Processing Errors",
    "CT4":"Payment processing interface and batch error resolution",
    "CT5":"Incident response",
    "CT6":"Master/Reference data input validation",
    "CT14":"Logging and monitoring","CT22":"Mistaken Internet Payment Reports",
    "CT23":"Provision of Confirmations and Notifications",
    "CT24":"Treatment of Unauthorised and Disputed Transactions",
    "CT27":"Records retention","CT28":"Crisis Management planning and testing",
}

def parse_json_response(text):
    text = re.sub(r'^```(?:json)?\s*','',text.strip())
    text = re.sub(r'\s*```$','',text).strip()
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        m = re.search(r'\{.*\}', text, re.DOTALL)
        if m:
            return json.loads(m.group(0))
        raise

def load_checkpoint(path):
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
                done.add(rec.get("control_id",""))
            except json.JSONDecodeError:
                continue
    return done

def write_checkpoint(rec, path):
    with open(path, "a", encoding="utf-8") as f:
        f.write(json.dumps(rec, ensure_ascii=True) + "\n")

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
#  BUILD PROMPT PER CONTROL
# ─────────────────────────────────────────────────────────────────────────────

def build_candidates_block(cand_rows):
    """Format the top-15 candidates as a structured text block."""
    blocks = []
    for _, r in cand_rows.iterrows():
        app_flag  = "YES" if r.get("app_match_fired") else "no"
        prod_flag = "YES" if r.get("product_match_fired") else "no"
        block = (
            f"Rank {int(r['candidate_rank']):02d} | "
            f"Composite: {float(r['composite_score']):.4f} | "
            f"Semantic: {float(r['semantic_score']):.4f} | "
            f"App match: {app_flag} | Product match: {prod_flag} | "
            f"CT align: {float(r['ct_alignment_score']):.1f}\n"
            f"  UUID       : {r['l3_process_UUID']}\n"
            f"  L3 Activity: {safe_str(r.get('l3_activity_name'))}\n"
            f"  L2 Process : {safe_str(r.get('l2_process_name'))}\n"
            f"  Category   : {safe_str(r.get('process_category'))}\n"
            f"  Stage      : {safe_str(r.get('process_lifecycle_stage'))}\n"
            f"  Apps       : {safe_str(r.get('process_apps'))[:300]}\n"
            f"  Products   : {safe_str(r.get('process_products'))[:200]}\n"
            f"  Description: {safe_str(r.get('l3_activity_description'))[:400]}\n"
            f"  Matched apps : {safe_str(r.get('matched_apps'))}\n"
            f"  Matched prods: {safe_str(r.get('matched_products'))}"
        )
        blocks.append(block)
    return "\n\n".join(blocks)


def build_prompt(ctrl_row, ctrl_full, cand_rows):
    cid       = safe_str(ctrl_row.get("Control_ID",""))
    ct_code   = norm_gold_ctrl(ctrl_full.get("gold_control",""))
    ct_title  = CT_TITLES.get(ct_code,"") if ct_code else ""
    pop_type  = safe_str(ctrl_row.get("population_type",""))
    ctrl_apps = safe_str(ctrl_row.get("ctrl_detected_apps","")) or "(none detected)"
    ctrl_prod = safe_str(ctrl_row.get("ctrl_detected_products","")) or "(none detected)"
    candidates_block = build_candidates_block(cand_rows)

    prompt = PROMPT_TEMPLATE.format(
        ctrl_id        = cid,
        ctrl_name      = safe_str(ctrl_full.get("CTRL_NAME","")),
        ct_code        = ct_code or "(unknown)",
        ct_title       = ct_title or "(unknown)",
        population_type= pop_type,
        ctrl_desc      = safe_str(ctrl_full.get("CTRL_DESC",""))[:2000] or "(not provided)",
        ctrl_activity  = safe_str(ctrl_full.get("CTRL_DESC_OF_CTRL",""))[:1000] or "(not provided)",
        ctrl_monitored = safe_str(ctrl_full.get("CTRL_MNTRD",""))[:600] or "(not provided)",
        ctrl_evidenced = safe_str(ctrl_full.get("CTRL_EVDNCD",""))[:600] or "(not provided)",
        ctrl_apps      = ctrl_apps,
        ctrl_products  = ctrl_prod,
        candidates_block= candidates_block,
    )
    return prompt

# ─────────────────────────────────────────────────────────────────────────────
#  LLM CALL
# ─────────────────────────────────────────────────────────────────────────────

def call_llm(client, prompt, ctrl_id):
    kwargs = {
        "model":   MODEL,
        "messages": [
            {"role":"system","content":SYSTEM_PROMPT},
            {"role":"user",  "content":prompt},
        ],
        "max_completion_tokens": MAX_COMPLETION_TOKENS,
    }
    if MODEL.startswith("gpt-5"):
        kwargs["reasoning_effort"] = REASONING_EFFORT

    for attempt in range(1, RETRY_COUNT+1):
        try:
            t0  = time.time()
            rsp = client.chat.completions.create(**kwargs)
            lat = int((time.time()-t0)*1000)
            text= rsp.choices[0].message.content or ""
            parsed = parse_json_response(text)
            usage = {}
            if hasattr(rsp,"usage") and rsp.usage:
                u = rsp.usage
                usage = {
                    "input_tokens":  getattr(u,"prompt_tokens",0),
                    "output_tokens": getattr(u,"completion_tokens",0),
                }
                d = getattr(u,"completion_tokens_details",None)
                if d:
                    usage["reasoning_tokens"] = getattr(d,"reasoning_tokens",0)
            usage["latency_ms"] = lat
            return parsed, usage
        except Exception as e:
            if attempt < RETRY_COUNT:
                wait = min(RETRY_BASE * 2**(attempt-1) + random.uniform(0,2), 60)
                print(f"    Retry {attempt} for {ctrl_id}: {e} — wait {wait:.0f}s")
                time.sleep(wait)
            else:
                raise

# ─────────────────────────────────────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--force", action="store_true",
                        help="Ignore checkpoint and reprocess all controls.")
    args = parser.parse_args()

    print("=" * 70)
    print("  Multi-Signal Linkage LLM Reranker")
    print(f"  Model    : {MODEL}  ({REASONING_EFFORT})")
    print(f"  Input    : {CANDIDATES_FILE.name}")
    print("=" * 70)

    if not CANDIDATES_FILE.exists():
        print(f"\n  ERROR: {CANDIDATES_FILE} not found.")
        print("  Run build_linkage_candidates.py first.")
        return

    # ── Load candidates ───────────────────────────────────────────────────────
    print("\n  Loading candidates...")
    candidates = pd.read_excel(CANDIDATES_FILE, sheet_name="candidate_pairs",
                               dtype=str, engine="openpyxl")
    signal_summary = pd.read_excel(CANDIDATES_FILE, sheet_name="signal_summary",
                                   dtype=str, engine="openpyxl")

    # Fix numeric columns
    for col in ["composite_score","semantic_score","app_score",
                "product_score","ct_alignment_score","candidate_rank"]:
        if col in candidates.columns:
            candidates[col] = pd.to_numeric(candidates[col], errors="coerce").fillna(0)
    for col in ["app_match_fired","product_match_fired"]:
        if col in candidates.columns:
            candidates[col] = candidates[col].map(
                {"True":True,"False":False,True:True,False:False}).fillna(False)

    all_ctrl_ids = candidates["Control_ID"].unique().tolist()
    print(f"  Controls to process: {len(all_ctrl_ids):,}")

    # ── Load full control details ─────────────────────────────────────────────
    print("  Loading control details...")
    controls_full = pd.read_excel(CONTROLS_FILE, dtype=str, engine="openpyxl")
    controls_full.columns = controls_full.columns.str.strip()
    controls_full["Control_ID"] = controls_full["Control_ID"].str.strip()
    ctrl_lookup = controls_full.set_index("Control_ID").to_dict("index")

    sig_lookup = signal_summary.set_index("Control_ID").to_dict("index") \
        if "Control_ID" in signal_summary.columns else {}

    # ── Checkpoint ────────────────────────────────────────────────────────────
    if args.force and CHECKPOINT.exists():
        CHECKPOINT.unlink()
        print("  --force: checkpoint cleared.")
    done_ids = load_checkpoint(CHECKPOINT)
    pending  = [c for c in all_ctrl_ids if c not in done_ids]
    print(f"  Remaining: {len(pending):,}  (done: {len(done_ids):,})")

    if not pending:
        print("  All controls processed. Regenerating outputs...")
    else:
        print("\n  Initialising Azure OpenAI client...")
        client = init_client()
        print("  Client ready.\n")

        total = len(pending)
        t_in, t_out = 0, 0
        run_start = time.time()

        for idx, cid in enumerate(pending, 1):
            pct = idx/total*100
            bar = "#"*int(pct/5) + "."*(20-int(pct/5))
            elapsed = time.time()-run_start
            eta_s   = (elapsed/idx)*(total-idx) if idx>1 else 0
            eta_str = (f"{int(eta_s//3600)}h {int((eta_s%3600)//60)}m"
                       if eta_s>60 else f"{int(eta_s)}s")
            cost_aud = ((t_in/1e6*INPUT_PRICE_USD_PER_M +
                         t_out/1e6*OUTPUT_PRICE_USD_PER_M)/AUD_USD_RATE)
            print(f"  [{bar}] {pct:5.1f}%  {idx}/{total}  "
                  f"ETA {eta_str}  Cost A${cost_aud:.2f}")

            ctrl_cands = candidates[candidates["Control_ID"]==cid].sort_values(
                "candidate_rank")
            ctrl_full  = ctrl_lookup.get(cid, {})
            ctrl_sig   = sig_lookup.get(cid, {})

            try:
                prompt = build_prompt(ctrl_sig, ctrl_full, ctrl_cands)
                result, usage = call_llm(client, prompt, cid)
            except Exception as e:
                print(f"    ERROR for {cid}: {e}")
                write_checkpoint({
                    "control_id": cid,
                    "recommendations": [],
                    "no_linkage_recommended": True,
                    "no_linkage_reason": f"Script error: {e}",
                    "status": "error",
                }, CHECKPOINT)
                continue

            result["control_id"] = cid
            result["status"]     = "success"

            n_recs  = len(result.get("recommendations",[]))
            in_tok  = usage.get("input_tokens",0)
            out_tok = usage.get("output_tokens",0)
            t_in   += in_tok
            t_out  += out_tok
            rt = f" | {usage.get('reasoning_tokens')} thinking" \
                 if usage.get("reasoning_tokens") else ""
            no_link = result.get("no_linkage_recommended", False)
            print(f"         {cid[:22]}  {usage.get('latency_ms')}ms  "
                  f"{in_tok}in/{out_tok}out{rt}  → "
                  f"{'NO LINK' if no_link else f'{n_recs} rec(s)'}")

            write_checkpoint(result, CHECKPOINT)
            if SLEEP_BETWEEN > 0:
                time.sleep(SLEEP_BETWEEN)

        final_cost = ((t_in/1e6*INPUT_PRICE_USD_PER_M +
                       t_out/1e6*OUTPUT_PRICE_USD_PER_M)/AUD_USD_RATE)
        print(f"\n  This run: {len(pending):,} controls  "
              f"{t_in:,} in  {t_out:,} out  A${final_cost:.2f}")

    # ── Build output DataFrames ───────────────────────────────────────────────
    print("\n  Building output DataFrames from checkpoint...")
    rec_rows, summary_rows = [], []

    with open(CHECKPOINT,"r",encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                rec = json.loads(line)
            except json.JSONDecodeError:
                continue

            cid  = rec.get("control_id","")
            recs = rec.get("recommendations",[])
            ctrl = ctrl_lookup.get(cid,{})
            sig  = sig_lookup.get(cid,{})
            pop  = candidates[candidates["Control_ID"]==cid]["population_type"].iloc[0] \
                   if not candidates[candidates["Control_ID"]==cid].empty else ""

            for r in recs:
                rec_rows.append({
                    "population_type":        pop,
                    "Control_ID":             cid,
                    "CTRL_NAME":              safe_str(ctrl.get("CTRL_NAME","")),
                    "gold_control_code":      norm_gold_ctrl(ctrl.get("gold_control","")),
                    "l3_process_UUID":        safe_str(r.get("l3_process_UUID","")),
                    "l3_activity_name":       safe_str(r.get("l3_activity_name","")),
                    "recommended_for_linkage":r.get("recommended_for_linkage",True),
                    "confidence":             safe_str(r.get("confidence","")),
                    "primary_signal":         safe_str(r.get("primary_signal","")),
                    "rationale":              safe_str(r.get("rationale","")),
                    "requires_sme_review":    True,
                    "ctrl_detected_apps":     safe_str(sig.get("ctrl_detected_apps","")),
                    "ctrl_detected_products": safe_str(sig.get("ctrl_detected_prods","")),
                    "linkage_method":         "multi_signal_llm",
                    "status":                 rec.get("status",""),
                })

            summary_rows.append({
                "population_type":      pop,
                "Control_ID":           cid,
                "CTRL_NAME":            safe_str(ctrl.get("CTRL_NAME","")),
                "gold_control_code":    norm_gold_ctrl(ctrl.get("gold_control","")),
                "n_recommendations":    len(recs),
                "has_high_confidence":  any(r.get("confidence")=="High" for r in recs),
                "has_any_recommendation": len(recs) > 0,
                "no_linkage_recommended": rec.get("no_linkage_recommended",False),
                "no_linkage_reason":    safe_str(rec.get("no_linkage_reason","")),
                "primary_signal":       (recs[0].get("primary_signal","") if recs else ""),
                "ctrl_detected_apps":   safe_str(sig.get("ctrl_detected_apps","")),
                "status":               rec.get("status",""),
            })

    rec_df     = pd.DataFrame(rec_rows)
    summary_df = pd.DataFrame(summary_rows)

    print(f"  Recommendations: {len(rec_df):,} control-process pairs")
    print(f"  Controls with ≥1 recommendation: "
          f"{summary_df['has_any_recommendation'].sum():,}")
    print(f"  Controls with no linkage found: "
          f"{summary_df['no_linkage_recommended'].sum():,}")

    if not rec_df.empty:
        conf_dist = rec_df["confidence"].value_counts()
        for conf, n in conf_dist.items():
            print(f"    {conf}: {n:,}")

    # ── Load existing deterministic linkages for combined view ────────────────
    det_df = pd.DataFrame()
    if LINKAGE_WB.exists():
        try:
            det_detail = pd.read_excel(
                LINKAGE_WB, sheet_name="linked_payment_detail",
                dtype=str, engine="openpyxl")
            det_detail.columns = det_detail.columns.str.strip()
            if "Control_ID" in det_detail.columns:
                det_df = det_detail[
                    ["Control_ID","l3_process_UUID","l3_activity_name",
                     "process_category","process_lifecycle_stage"]
                ].copy() if all(c in det_detail.columns for c in
                    ["Control_ID","l3_process_UUID","l3_activity_name",
                     "process_category","process_lifecycle_stage"]) else pd.DataFrame()
                det_df["linkage_method"] = "deterministic"
                det_df["confidence"]     = "Deterministic"
                det_df["requires_sme_review"] = False
        except Exception as e:
            print(f"  WARNING: Could not load deterministic linkages: {e}")

    # ── Validation ────────────────────────────────────────────────────────────
    checks = []
    def chk(name, expected, actual, note=""):
        passed = str(expected) == str(actual)
        checks.append({"check":name,"expected":str(expected),
                        "actual":str(actual),"pass":"PASS" if passed else "FAIL","note":note})
        print(f"  [{'PASS' if passed else 'FAIL'}] {name}")

    print("\n  Validation checks...")
    total_target = len(all_ctrl_ids)
    processed    = len(summary_df)
    chk("All target controls processed", total_target, processed)
    if not rec_df.empty:
        chk("All recommended linkages have l3_process_UUID", 0,
            rec_df["l3_process_UUID"].eq("").sum())
        chk("requires_sme_review always true", 0,
            (~rec_df["requires_sme_review"]).sum())
        chk("All recommendations flagged as multi_signal_llm", 0,
            (rec_df["linkage_method"] != "multi_signal_llm").sum())

    # ── Write outputs ─────────────────────────────────────────────────────────
    print(f"\n  Writing to {OUTPUT_FILE}...")
    MULTI_DIR.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as w:
        # Summary at top
        summary_stats = pd.DataFrame([
            ("Total controls processed",             len(summary_df)),
            ("Controls with ≥1 recommendation",      int(summary_df["has_any_recommendation"].sum())),
            ("Controls with no linkage found",        int(summary_df["no_linkage_recommended"].sum())),
            ("Total control-process recommendations", len(rec_df)),
            ("High confidence recommendations",
             int((rec_df["confidence"]=="High").sum()) if not rec_df.empty else 0),
            ("Medium confidence recommendations",
             int((rec_df["confidence"]=="Medium").sum()) if not rec_df.empty else 0),
            ("Low confidence recommendations",
             int((rec_df["confidence"]=="Low").sum()) if not rec_df.empty else 0),
            ("App match primary signal",
             int((rec_df["primary_signal"]=="app_match").sum()) if not rec_df.empty else 0),
            ("Product match primary signal",
             int((rec_df["primary_signal"]=="product_match").sum()) if not rec_df.empty else 0),
        ], columns=["Metric","Value"])
        summary_stats.to_excel(w, index=False, sheet_name="summary")

        rec_df.to_excel(     w, index=False, sheet_name="linkage_recommendations")
        summary_df.to_excel( w, index=False, sheet_name="summary_by_control")

        if not rec_df.empty:
            high_conf = rec_df[rec_df["confidence"]=="High"]
            if not high_conf.empty:
                high_conf.to_excel(w, index=False, sheet_name="high_confidence_links")

        no_link = summary_df[summary_df["no_linkage_recommended"]==True]
        if not no_link.empty:
            no_link.to_excel(w, index=False, sheet_name="no_linkage_found")

        if not det_df.empty and not rec_df.empty:
            combined_cols = ["Control_ID","l3_process_UUID","l3_activity_name",
                             "process_category","process_lifecycle_stage",
                             "linkage_method","confidence","requires_sme_review"]
            det_part = det_df[[c for c in combined_cols if c in det_df.columns]].copy()
            new_part = rec_df[[c for c in combined_cols if c in rec_df.columns]].copy()
            combined = pd.concat([det_part, new_part], ignore_index=True)
            combined.to_excel(w, index=False, sheet_name="combined_all_linkages")

        pd.DataFrame(checks).to_excel(w, index=False, sheet_name="validation_checks")

    print(f"\n  Sheets:")
    print(f"    summary | linkage_recommendations | summary_by_control")
    print(f"    high_confidence_links | no_linkage_found")
    print(f"    combined_all_linkages | validation_checks")
    print(f"\n  Output: {OUTPUT_FILE}")
    print("\n" + "=" * 70)
    print("  Done.")
    print("=" * 70)


if __name__ == "__main__":
    main()
