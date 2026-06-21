# LRM Context Pack — Smoke Test Results, Fixes Applied, and Recommended Next Tests

## Purpose
This document is designed as **context for a large reasoning model (LRM)**. It summarises:
1. the **test results** from the latest smoke test run,
2. the **environment / script errors that were identified and fixed**, and
3. the **recommended next tests** focused on hallucination control, reproducibility, and model suitability for the Payments Controls PoC.

---

## 1) Canonical environment context

### Working endpoint
- `https://ai.eng.azure.srv.westpac.com.au`

### API version
- `2024-10-21`

### Working chat model used in the current smoke test
- `gpt-5.4`

### Working embedding model used in the current smoke test
- `text-embedding-3-large`

### Required shell / network assumptions
- Shell: **PowerShell**
- Azure authentication: **Azure CLI (`az login`)** with `DefaultAzureCredential()`
- Subscription: `cst-a00c3d-cin03`
- Proxy must be enabled:
  - `HTTP_PROXY=http://127.0.0.1:9000`
  - `HTTPS_PROXY=http://127.0.0.1:9000`
- `NO_PROXY` / `no_proxy` should be cleared for this working path
- Python client should **not** disable proxy inheritance (`trust_env=False` must not be used)

---

## 2) Latest test run metadata

### Run identifier
- `20260620_210320`

### Overall outcome
- **4 / 6 tests passed**
- Environment status: **working**
- Main remaining issues are **parameter / evaluation-design issues**, not endpoint connectivity issues

---

## 3) Summary of smoke test results

### Test 1 — Basic Connectivity
**Status:** PASS

**What happened:**
- Prompt sent to `gpt-5.4`
- Response returned successfully: `Hello!`
- Latency: `26320ms`
- Usage:
  - input tokens: `23`
  - output tokens: `6`
  - total tokens: `29`
  - reasoning tokens: `0`

**Interpretation:**
- Chat completions path is confirmed working end-to-end.
- Authentication, proxy, endpoint and deployment routing are all functioning.
- The long latency on a tiny prompt suggests there may be model / gateway overhead worth benchmarking later.

---

### Test 2 — Structured JSON Output
**Status:** PASS

**What happened:**
- The model was asked to classify a simple control and return JSON.
- Valid JSON was returned with keys:
  - `control_type`
  - `confidence`
  - `rationale`
- Latency: `3199ms`
- Usage:
  - input tokens: `72`
  - output tokens: `48`
  - total tokens: `120`

**Returned answer summary:**
- `control_type`: `Preventive`
- `confidence`: `0.98`
- rationale grounded in the approval-before-release wording

**Interpretation:**
- Structured extraction / classification workflows are viable.
- This is directly relevant for the Payments Controls PoC because it supports machine-readable outputs, auditability, and downstream scoring / linkage.

---

### Test 3 — Reasoning Effort Comparison (`low` / `medium` / `high`)
**Status:** FAIL

**What happened:**
- All three runs failed with the same 400 error.
- Error message:
  - `Unsupported value: 'temperature' does not support 0 with this model. Only the default (1) value is supported.`

**Interpretation:**
- This is **not** an endpoint or auth failure.
- It is a **parameter compatibility issue** in the test design.
- For this path/model combination, the reasoning-effort test should not force `temperature=0`.

**Implication for future tests:**
- Retest reasoning effort with:
  - temperature omitted entirely, or
  - model default temperature, or
  - a supported temperature setting

---

### Test 4 — Reproducibility (`temperature=0`, 3 repeated runs)
**Status:** WARN

**What happened:**
- All 3 runs completed successfully.
- All 3 runs classified the first control as:
  - `Detective`
  - confidence `0.95`
- The exact JSON text differed slightly across the 3 runs.

**Usage by run:**
- Run 1: total tokens `584`
- Run 2: total tokens `572`
- Run 3: total tokens `596`

**Interpretation:**
- The outputs were **semantically stable** but **not textually identical**.
- This is the correct way to think about LLM reproducibility in practice:
  - classification remained stable,
  - confidence remained stable,
  - rationale wording varied.

**Implication for future tests:**
- Reproducibility should be evaluated using:
  - consistency of label,
  - consistency of confidence range,
  - consistency of evidence / rationale meaning,
  rather than exact byte-for-byte text matching.

---

### Test 5 — Embedding Connectivity
**Status:** PASS

**What happened:**
- A control description was sent to `text-embedding-3-large`.
- Embedding returned successfully.
- Latency: `165ms`
- Vector dimensions: `3072`
- Dimensions matched expectation.
- Usage:
  - input tokens: `329`
  - total tokens: `329`

**Interpretation:**
- Embeddings path is working correctly.
- Embedding dimensionality is correct.
- This is directly usable for:
  - similarity matching,
  - duplicate / overlap detection,
  - obligation / incident / issue linkage,
  - clustering / retrieval workflows.

---

### Test 6 — Token Usage Logging
**Status:** PASS

**What happened:**
- Classification prompt completed successfully.
- Usage object was returned.
- Latency: `2716ms`
- Usage:
  - input tokens: `511`
  - output tokens: `145`
  - total tokens: `656`
  - reasoning tokens: `0`

**Returned classification summary:**
- `classification`: `Detective`
- `lifecycle_stage`: `Reporting`
- `confidence`: `0.84`
- rationale explicitly noted that the source metadata said Preventative, but the actual described activity is more consistent with post-event reconciliation / verification.

**Interpretation:**
- Token logging works.
- Usage metadata is available for later cost monitoring.
- The model demonstrated useful reasoning by distinguishing between source labeling and actual control behaviour.

**Nuance:**
- `reasoning_tokens` appeared as `0`, so no separately surfaced thinking-token consumption was observed in this run.
- This does not necessarily mean reasoning is absent; it means the API response did not report non-zero reasoning-token usage in this call.

---

## 4) Most important findings from the current test run

### A. The environment is now proven working
The successful chat, JSON, embedding, and token-usage tests confirm that:
- the endpoint works,
- Azure auth works,
- proxy routing works,
- the gateway can successfully invoke deployed models,
- and Python SDK calls are operational.

### B. Structured JSON output is already strong
For PoC purposes, this is one of the most valuable outcomes. It supports:
- machine-readable extraction,
- defensible classification artefacts,
- and easier auditability.

### C. Reproducibility concerns should be framed correctly
The run shows:
- stable classification,
- stable confidence,
- non-identical rationale wording.

This should be interpreted as:
- **semantic reproducibility = acceptable**
- **textual exact-repeat reproducibility = not guaranteed**

### D. Reasoning-effort testing still needs redesign
The failure in Test 3 is a test-parameter issue, not a model-access issue.

### E. Embeddings are validated and ready for real PoC use cases
Embedding connectivity and expected vector size were both confirmed.

---

## 5) Errors and issues that were previously fixed (important troubleshooting history)

### 5.1 Wrong endpoint path
**Problem:** Earlier tests used the raw Azure endpoint and/or incorrect suffixes.

**Fix:** Use the Westpac internal gateway endpoint only:
- `https://ai.eng.azure.srv.westpac.com.au`

### 5.2 Proxy bypass caused failures
**Problem:** Earlier scripts bypassed proxy behaviour using one or both of:
- `NO_PROXY` including `.openai.azure.com`
- `httpx.Client(trust_env=False, ...)`

**Fix:**
- clear `NO_PROXY` / `no_proxy`
- do not use `trust_env=False`
- allow normal environment proxy settings

### 5.3 Wrong subscription selected
**Problem:** Azure login initially defaulted to the wrong subscription.

**Fix:** explicitly set:
- `cst-a00c3d-cin03`
- subscription id: `6c72e6c5-ed48-4030-b29c-34e2849c9288`

### 5.4 Certificate / CA trust issues during `az login`
**Problem:** Azure CLI login originally failed with certificate verification errors.

**Fix:**
- resolve corporate CA / trust setup,
- use the correct PowerShell flow,
- get `az login` working,
- then rely on `DefaultAzureCredential()`.

### 5.5 Python script parameter mismatch for GPT‑5.4
**Problem:** The original smoke test used:
- `max_tokens`
which is unsupported in this path for `gpt-5.4`.

**Fix:**
- replace with `max_completion_tokens`

### 5.6 Python path / docstring parsing issue
**Problem:** Windows-path text in the top docstring caused:
- unicode escape / `\U` parsing error

**Fix:**
- make the docstring raw (`r""" ... """`) or escape backslashes

### 5.7 User accidentally tried to run an output folder instead of the script
**Problem:** ran:
- `python .\smoke_results\`
which failed because the folder had no `__main__.py`

**Fix:** run the actual file:
- `python .\smoke_test.py`

### 5.8 UNC path string warning
**Problem:** unescaped UNC path in a normal string caused `SyntaxWarning`.

**Fix:** use a raw string for the UNC path field.

### 5.9 Pydantic deprecation warnings in usage introspection
**Problem:** Test 6 produced deprecation warnings while inspecting usage object attributes.

**Fix recommendation (not yet required for function):**
- if polishing the script, avoid iterating over all instance attributes that trigger Pydantic deprecation warnings.
- These warnings do not invalidate the test result.

---

## 6) Recommended next testing program

The next testing program should explicitly address the two stakeholder concerns:
1. **hallucination / unsupported inference**, and
2. **reproducibility / consistency**.

### 6.1 Hallucination control tests (recommended)

#### Test A — Grounded classification only
Prompt pattern:
- instruct the model to classify **based only on provided text**
- explicitly forbid assumptions
- require the model to say `INSUFFICIENT_INFORMATION` if evidence is missing

**Goal:**
- test whether the model invents facts or cleanly signals insufficiency.

#### Test B — Evidence citation test
Prompt pattern:
- require output JSON to include:
  - classification,
  - confidence,
  - exact quoted evidence phrases from the source control text.

**Goal:**
- test whether the model grounds its conclusions in actual text.

#### Test C — Adversarial / contradiction test
Prompt pattern:
- create examples where source metadata says one thing (e.g. `Preventative`) but description clearly implies another (e.g. detective / post-event review).

**Goal:**
- test whether the model follows the evidence rather than blindly trusting labels.

#### Test D — Sparse / ambiguous input test
Prompt pattern:
- provide incomplete control text and ask for classification.

**Goal:**
- test whether the model says `INSUFFICIENT_INFORMATION` instead of guessing.

---

### 6.2 Reproducibility tests (recommended)

#### Test E — Multi-run semantic consistency
Run the same prompt 10 times and compare:
- classification label
- confidence range
- decision rationale meaning
- cited evidence phrases

**Pass criteria should be semantically defined, not textually exact.**

Suggested acceptance criteria:
- same classification in at least 9/10 runs,
- confidence within a small band (e.g. ±0.05 to 0.10),
- evidence phrases materially consistent.

#### Test F — Cross-model consistency
Run the same grounded prompt across:
- `gpt-4o`
- `gpt-5.4`
- `gpt-5.4-mini`

Compare:
- classification consistency,
- confidence,
- evidence grounding,
- latency,
- token usage.

**Goal:**
- identify the most reliable model for the PoC, not just the most powerful one.

---

### 6.3 Reasoning-effort retest (recommended)
Retest reasoning effort for GPT‑5 family, but redesign the test as follows:
- do **not** set `temperature=0`
- either omit temperature entirely, or let model default apply
- test `low`, `medium`, `high`
- capture:
  - latency,
  - token usage,
  - classification outcome,
  - quality / specificity of rationale,
  - whether reasoning tokens are surfaced.

**Goal:**
- assess whether higher reasoning effort actually improves control-sufficiency / gap-analysis quality enough to justify additional latency and cost.

---

### 6.4 Embedding quality tests (recommended)
Embeddings connectivity is already proven. Next recommended tests:

#### Test G — Pairwise similarity sanity check
Use a small set of controls and compare:
- similar controls vs unrelated controls
- expected ranking order

#### Test H — Obligation/control retrieval prototype
Use a few obligation / issue / incident sentences as queries and check whether the top embedding matches are intuitively correct.

#### Optional benchmark
Compare:
- `text-embedding-3-small`
- `text-embedding-3-large`

Measure:
- ranking quality,
- latency,
- token usage.

---

## 7) Recommended stakeholder framing

### On hallucination
The right control is **not** to assume the model will never hallucinate. Instead:
- require evidence from source text,
- explicitly forbid assumptions,
- flag insufficient information,
- and route ambiguous / low-confidence outcomes for review.

### On reproducibility
The right expectation is:
- **consistent decisions**, not necessarily **identical wording**.

The latest test run already supports this framing because:
- repeated outputs converged on the same classification (`Detective`) and same confidence (`0.95`),
- but wording differed slightly.

This is the right mental model for production use of LLMs in classification workflows.

---

## 8) Practical recommendations for the Payments Controls PoC

### Best immediate uses supported by current evidence
The current results support moving into:
- structured control classification,
- control metadata extraction,
- token/cost logging,
- embedding-based similarity / linkage experiments.

### Best next model benchmarks
Recommended next side-by-side comparisons:
- `gpt-4o`
- `gpt-5.4`
- `gpt-5.4-mini`

### Best settings bias for PoC
For auditable control workflows:
- prefer structured JSON outputs,
- prefer evidence-based grounding,
- prefer low-variability settings,
- evaluate reproducibility semantically rather than literally.

---

## 9) Final concise conclusion

### What has been proven
- The endpoint works.
- Chat completions work.
- Structured JSON output works.
- Embeddings work.
- Token-usage logging works.
- The model can reason usefully about a control even when source metadata may be misleading.

### What still needs work
- Reasoning-effort testing must be redesigned.
- Hallucination controls should be explicitly tested via grounding / evidence / insufficient-information prompts.
- Reproducibility should be measured statistically / semantically, not by exact string match.

### Overall conclusion
The environment is **production-viable for controlled PoC experimentation**, and the next phase should shift from connectivity testing to **assurance testing**:
- grounding,
- evidence citation,
- reproducibility bands,
- cross-model comparison,
- and embedding-relevance evaluation.
