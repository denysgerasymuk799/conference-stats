#!/usr/bin/env python3
"""
Classify ICLR abstracts into contribution archetypes using a comprehensive
multi-signal scoring classifier (no external API required).

Each paper is scored independently on five dimensions; the highest score wins.
Results are saved to iclr_2026/data/archetypes_classified.json.

Usage:
    python iclr_2026/src/classify_archetypes.py
"""

import json
import re
from pathlib import Path

import openpyxl

SCRIPT_DIR = Path(__file__).parent
DATA_DIR   = SCRIPT_DIR.parent / "data"
OUT_FILE   = DATA_DIR / "archetypes_classified.json"

YEARS      = [2022, 2023, 2024, 2025, 2026]
ARCHETYPES = ["New Method", "Empirical Study", "Theory", "Benchmark / Dataset", "Other"]

AI_SAFETY_TERMS = [
    "privacy", "jailbreak", "explainab",
    "ai safety", "alignment", "adversarial", "fairness", "trustworthy",
    "interpretab", "hallucination", "red team", "value alignment",
    "reward hacking", "decepti", "safety", "toxic", "harmful",
    "societal", "ethic",
    "unlearn", "backdoor", "bias", "watermark", "poison", "attack", "vulnerab",
]

# ---------------------------------------------------------------------------
# Compiled patterns
# ---------------------------------------------------------------------------

# --- Theory ---
_T_STRONG = re.compile(
    r"\b(theorem|lemma|corollary|we prove|formal proof|provably optimal|"
    r"convergence (rate|guarantee)|regret bound|sample complexity|"
    r"generalization bound|minimax optimal|pac.learning|"
    r"information.theoretic(ally)?|tight bound|we establish a bound|"
    r"we derive (a|an|tight|optimal|sharp)|lower bound on|upper bound on|"
    r"we show (a|an) .{0,20}(bound|rate|complexity)|"
    r"statistical (guarantee|optimality))\b",
    re.I,
)
_T_MEDIUM = re.compile(
    r"\b(provably|formally (show|prove|characterize)|we establish|"
    r"we derive|finite.sample|excess risk|concentration inequality|"
    r"we analyze the convergence|we characterize the (optimal|minimax)|"
    r"hardness result|np.hard|computationally intractable)\b",
    re.I,
)
_T_PA = re.compile(
    r"\b(theory|theoretic|learning theory|optimization|statistical)\b", re.I
)

# --- Benchmark / Dataset ---
_B_VERB = re.compile(
    r"\b(we introduce|we release|we create|we collect|we curate|"
    r"we build|we construct|we develop|we present) (a |an |our |the )?new\b",
    re.I,
)
_B_VERB2 = re.compile(
    r"\b(we introduce|we release|we create|we collect|we curate|"
    r"we build|we construct)\b",
    re.I,
)
_B_NOUN = re.compile(
    r"\b(benchmark|dataset|testbed|evaluation suite|test suite|corpus|"
    r"leaderboard|data collection|annotated dataset|crowdsourced)\b",
    re.I,
)
_B_PA = re.compile(r"datasets and benchmarks", re.I)

# --- Empirical Study ---
_E_STRONG = re.compile(
    r"\b(we study|we investigate|we analyze|we examine|we characterize|"
    r"we audit|we survey|we assess|we explore (how|why|whether|the|what)|"
    r"empirical (study|analysis|investigation|evaluation|assessment)|"
    r"systematic (study|analysis|review|evaluation)|"
    r"we conduct (a|an) (study|survey|analysis|evaluation|investigation)|"
    r"we perform (a|an) (study|analysis|investigation|evaluation)|"
    r"large.scale (study|analysis|evaluation)|"
    r"in.depth (study|analysis|investigation)|"
    r"we provide (a|an) (study|analysis|evaluation|investigation)|"
    r"we measure (how|whether|the extent|the (effect|impact)))\b",
    re.I,
)
_E_MEDIUM = re.compile(
    r"\b(we find that|we observe that|we show that [a-z ,]{0,40}(fails|"
    r"does not|cannot|struggles|exhibits|has a|is (biased|sensitive|prone)|"
    r"correlates|affects|leads to)|"
    r"our (findings|results) (suggest|indicate|show|reveal|demonstrate)|"
    r"our analysis (reveals|shows|indicates|suggests)|"
    r"we quantify|we measure the (effect|impact|influence|degree|extent)|"
    r"we evaluate (whether|how|the extent|the (effect|impact|trade))|"
    r"to (understand|investigate|study) (why|how|whether|the|what)|"
    r"we (reveal|uncover|discover) (that|how|why|a|an)|"
    r"our (study|investigation|analysis) (shows|finds|reveals|suggests))\b",
    re.I,
)
_E_PA = re.compile(
    r"\b(empirical|analysis|study|understanding|interpretability|"
    r"evaluation|robustness|security|safety|privacy|fairness)\b",
    re.I,
)
# Penalise: empirical phrases that are actually just ablation / performance claims
_E_PENALTY = re.compile(
    r"\b(our (method|model|algorithm|system|framework|architecture|approach|network) "
    r".{0,40}(achieve|outperform|surpass|exceed|improve|state.of.the.art|sota)|"
    r"we show (superior|better|improved|significant improvement)|"
    r"we demonstrate (that our|superior|state.of.the.art))\b",
    re.I,
)

# --- New Method ---
_M_DIRECT = re.compile(
    r"\b(our (method|algorithm|model|architecture|network|system|technique|"
    r"loss( function)?|objective( function)?|regularizer|optimizer|"
    r"estimator|generator|discriminator|encoder|decoder|classifier))\b",
    re.I,
)
_M_VERB = re.compile(
    r"\b(we propose|we introduce|we develop|we design|we present|"
    r"we build|we formulate)\b",
    re.I,
)
_M_NOUN = re.compile(
    r"\b(method|algorithm|model|architecture|network|system|technique|"
    r"framework|module|mechanism|policy|agent|regularizer|loss function|"
    r"objective|estimator|generator|discriminator|encoder|decoder|"
    r"classifier|solver|operator|kernel|sampler)\b",
    re.I,
)
_M_EMPIRICAL_BLOCKER = re.compile(
    r"\b(study|analysis|analyses|evaluat\w+|protocol|survey|"
    r"investigation|examination|audit|measurement|review)\b",
    re.I,
)
_M_FINETUNE = re.compile(
    r"\b(we fine.tune|we finetune|we pretrain|we pre.train|"
    r"we distill|we prune|we quantize)\b",
    re.I,
)
_M_PA = re.compile(
    r"\b(machine learning|deep learning|generative|optimization|"
    r"reinforcement|graph|natural language|vision|audio|speech)\b",
    re.I,
)


# ---------------------------------------------------------------------------
# Scorer
# ---------------------------------------------------------------------------

def score(abstract: str, primary_area: str) -> dict[str, float]:
    a  = abstract or ""
    pa = (primary_area or "").lower()
    s  = {k: 0.0 for k in ARCHETYPES}

    # ---- Theory ----
    s["Theory"] += len(_T_STRONG.findall(a)) * 4.0
    s["Theory"] += len(_T_MEDIUM.findall(a)) * 2.0
    if _T_PA.search(pa):
        s["Theory"] += 1.5

    # ---- Benchmark / Dataset ----
    if _B_PA.search(pa):
        s["Benchmark / Dataset"] += 8.0   # very strong primary-area signal
    # verb + "new" immediately → strong benchmark phrase
    s["Benchmark / Dataset"] += len(_B_VERB.findall(a)) * 3.0
    # verb within 120 chars of benchmark noun
    for m in _B_VERB2.finditer(a):
        snippet = a[m.end(): min(len(a), m.end() + 120)]
        if _B_NOUN.search(snippet):
            s["Benchmark / Dataset"] += 2.5

    # ---- Empirical Study ----
    s["Empirical Study"] += len(_E_STRONG.findall(a)) * 3.5
    s["Empirical Study"] += len(_E_MEDIUM.findall(a)) * 1.5
    if _E_PA.search(pa):
        s["Empirical Study"] += 1.5
    s["Empirical Study"] -= len(_E_PENALTY.findall(a)) * 1.5

    # ---- New Method ----
    s["New Method"] += len(_M_DIRECT.findall(a)) * 2.5
    if _M_FINETUNE.search(a):
        s["New Method"] += 3.0
    # verb + method noun proximity (blocked by empirical noun appearing first)
    for m in _M_VERB.finditer(a):
        snippet = a[m.end(): min(len(a), m.end() + 130)]
        mn = _M_NOUN.search(snippet)
        if mn:
            blocker = _M_EMPIRICAL_BLOCKER.search(snippet[: mn.start()])
            if not blocker:
                s["New Method"] += 2.5
    if _M_PA.search(pa):
        s["New Method"] += 0.5

    return s


def classify(abstract: str, primary_area: str) -> str:
    # Hard rule: primary area explicitly says datasets/benchmarks
    if re.search(r"datasets and benchmarks", (primary_area or ""), re.I):
        return "Benchmark / Dataset"

    sc = score(abstract, primary_area)
    best_arch  = max(sc, key=lambda k: sc[k])
    best_score = sc[best_arch]

    if best_score <= 0.0:
        return "Other"

    # Theory vs Method tie-break: prefer Theory when scores are close
    if (sc["Theory"] >= 3.0
            and sc["Theory"] >= sc["New Method"] * 0.75
            and sc["Theory"] >= sc["Empirical Study"] * 0.75):
        return "Theory"

    return best_arch


# ---------------------------------------------------------------------------
# Per-year processing
# ---------------------------------------------------------------------------

def process_year(year: int) -> dict:
    path = DATA_DIR / f"ICLR_{year}.xlsx"
    wb   = openpyxl.load_workbook(path)
    ws   = wb.active
    h    = {c.value: i for i, c in enumerate(ws[1])}

    def col(row, name):
        idx = h.get(name)
        return str(row[idx] or "") if idx is not None else ""

    all_counts    = {a: 0 for a in ARCHETYPES}
    safety_counts = {a: 0 for a in ARCHETYPES}
    total = safety_total = 0

    for row in ws.iter_rows(min_row=2, values_only=True):
        abstract = col(row, "Abstract")
        pa       = col(row, "Primary Area")
        search   = f"{col(row,'Title')} {col(row,'Keywords')} {pa}".lower()
        is_safety = any(t in search for t in AI_SAFETY_TERMS)

        arch = classify(abstract, pa)
        total += 1
        all_counts[arch] += 1
        if is_safety:
            safety_total += 1
            safety_counts[arch] += 1

    return {
        "total":                total,
        "ai_safety_total":      safety_total,
        "archetypes":           all_counts,
        "ai_safety_archetypes": safety_counts,
    }


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    results = {}
    for year in YEARS:
        print(f"ICLR {year} …", end=" ", flush=True)
        results[str(year)] = process_year(year)
        d = results[str(year)]
        parts = "  ".join(
            f"{a[:3]}={d['archetypes'][a]}" for a in ARCHETYPES
        )
        safe_parts = "  ".join(
            f"{a[:3]}={d['ai_safety_archetypes'][a]}" for a in ARCHETYPES
        )
        print(f"total={d['total']}  [{parts}]")
        print(f"         safety={d['ai_safety_total']}  [{safe_parts}]")

    OUT_FILE.write_text(json.dumps(results, indent=2))
    print(f"\nSaved → {OUT_FILE}")


if __name__ == "__main__":
    main()
