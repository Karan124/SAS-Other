# Duplicate Payment Control Detection — Approach Overview

## What We Are Trying to Do

Identify which of our 708 payment controls are duplicates or near-duplicates of each other, and where consolidation or harmonisation opportunities exist across the payment control population.

---

## Understanding the Existing Data Model

JUNO already has fields that partially answer this question. Before running any analysis, we use three fields to classify every control:

| CTRL_CATEGORY | COMMON_CTRL_TYP | COMMON_CTRL_REFERENCE | Meaning |
|---|---|---|---|
| COMMON | MASTER | blank | The canonical, authoritative version of a shared control |
| COMMON | INSTANCE | [Control ID] | A local copy of a master common control — intentional by design |
| UNIQUE | blank | blank | A control specific to one business area — primary target for analysis |
| CENTRALISED | blank | blank | A control owned and operated by a central function (e.g. Group Risk, Group Technology) — applied broadly but managed from the centre |

**Key insight:** INSTANCE controls are already identified by JUNO as intentional copies of a MASTER. They are not duplicates — they are a deliberate design decision. Our analysis does not need to rediscover them. What we are looking for is everything that falls *outside* this structure.

---

## The Four Findings We Are Looking For

**Finding 1 — Duplicate Control**
Two UNIQUE controls, same objective, same business area. One is a redundant copy of the other with no intentional rationale. Retire the lower-quality one.

**Finding 2 — Common Control Redundancy**
A UNIQUE control is doing the same thing as an existing MASTER or CENTRALISED control — but has not been set up as an INSTANCE. It is a shadow copy that bypasses the established control structure. Retire the UNIQUE and ensure the business area is using the MASTER or CENTRALISED control instead.

**Finding 3 — Common Control Elevation Opportunity**
Two or more UNIQUE controls across different business areas are doing the same thing, but no MASTER exists for it yet. One should become the MASTER; the others become INSTANCES. Any future division wanting the same control uses an INSTANCE rather than creating a new UNIQUE.

**Finding 4 — Redundant Instance**
Two INSTANCE controls pointing to the same MASTER exist within the same business area. Since they are both copies of the same MASTER in the same division, one is unnecessary. This finding is detectable with a simple database query — no AI required.

---

## Step-by-Step Approach

**Step 1 — Pre-classify the population**

Using CTRL_CATEGORY and COMMON_CTRL_TYP, divide the 708 controls into four groups:
- MASTER controls — reference population
- INSTANCE controls — checked for Finding 4 (simple group-by on COMMON_CTRL_REFERENCE + business area)
- CENTRALISED controls — reference population
- UNIQUE controls — primary target for similarity analysis

**Step 2 — Score similarity for UNIQUE controls**

For every UNIQUE control, compute a composite similarity score against every other UNIQUE control, every MASTER, and every CENTRALISED control. This produces a ranked list of candidate pairs.

The composite score uses six signals:

| Signal | Weight | How |
|---|---|---|
| Control description similarity | 35% | Text embeddings (AI reads CTRL_NAME + CTRL_DESC + CTRL_DESC_OF_CTRL and finds controls that mean the same thing even if worded differently) |
| L3 process overlap | 30% | Jaccard similarity on linked L3 process UUIDs — two controls governing the exact same processes are near-certainly duplicates |
| Product overlap | 15% | Jaccard similarity on inferred product sets |
| Structural match | 10% | Same CT type, same control nature (preventative/detective), same category |
| Alfabet application overlap | 5% | Same platforms referenced in control descriptions or process linkage |
| Lifecycle stage overlap | 5% | Jaccard similarity on lifecycle stages of linked processes |

**Step 3 — Determine finding type**

For every candidate pair above the similarity threshold:

```
UNIQUE vs UNIQUE, same division         → Finding 1 (Duplicate Control)
UNIQUE vs UNIQUE, different divisions   → Finding 3 (Elevation Opportunity)
UNIQUE vs MASTER or CENTRALISED         → Finding 2 (Common Control Redundancy)
```

Business area is determined from CTRL_FLDR (the hierarchical folder path such as
/Westpac Group/Customer and Corporate Services/Customer Solutions/Fraud and Scam Operations)
and supplemented by CTRL_OWNER mapped to business unit via the HR organisational table.

**Step 4 — Group into clusters**

Rather than analysing isolated pairs, we build a similarity graph (controls as nodes, candidate pairs as edges) and apply Louvain community detection to find groups of mutually similar controls. A cluster of five controls all doing the same thing is a far stronger finding than five separate pairs.

**Step 5 — LLM validation per cluster**

GPT-5.4 reviews each cluster in full — reading complete control descriptions, CT outcomes, linked processes, products, lifecycle stages, business areas, and effectiveness ratings — and determines:
1. Are these genuine duplicates, partially overlapping, or complementary controls?
2. Which is best of breed and why?
3. What is the recommended action?

All LLM recommendations go to SME review for the final call.

**Step 6 — Best of breed selection**

Within each confirmed duplicate cluster, controls are ranked on:
- Effectiveness ratings (CTRL_ASSESS_RTNG, CTRL_OE_RTNG, CTRL_DE_RTNG): Effective scores higher than Requires Improvement
- Key control flag (CTRL_KEY_CONTRL = Yes → strong preference)
- Second line of defence rating (CTRL_RATING_2LOD if populated)
- Quality of description, evidence (CTRL_EVDNCD), and monitoring (CTRL_MNTRD) — assessed by the LLM

The highest-scoring control is flagged best of breed. Others in the cluster are retirement or harmonisation candidates.

---

## Calibration

Since no known duplicate pairs currently exist, we will run the algorithm first and present the top 20 pairs by composite score to a senior SME for manual review. Their confirmed or rejected labels will calibrate the final thresholds before LLM validation runs at full scale.

---

## Dependencies

| Item | Status |
|---|---|
| 708 payment controls with all text fields | Available |
| Combined corrected linkage file (processes + products) | Pending |
| Alfabet application list and process-level linkages | Available |
| CTRL_FLDR for business area hierarchy | Available |
| HR organisational table (CTRL_OWNER → business unit) | Pending |
