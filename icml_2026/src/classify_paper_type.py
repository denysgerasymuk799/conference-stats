#!/usr/bin/env python3
"""
Classify ICML papers into one of three *paper types* using a multi-signal
scoring classifier over the title and abstract (no external API required).

    Benchmark   – the headline contribution is a new benchmark, dataset,
                  testbed, evaluation suite or leaderboard.
    Evaluation  – the headline contribution is measuring / auditing / comparing
                  *existing* models or methods (empirical study, systematic
                  evaluation, analysis) without releasing a new resource.
    Other       – everything else (new methods, theory, applications, …).

Precedence is Benchmark > Evaluation > Other: a paper that both releases a
benchmark and evaluates models on it counts as a Benchmark paper.

Used by fetch_icml_openreview.py to fill the "Paper Type" column, and runnable
standalone to (re)label an existing spreadsheet in place:

    python icml_2026/src/classify_paper_type.py [path/to/ICML_2026.xlsx]
"""

import re
import sys
from collections import Counter
from pathlib import Path

PAPER_TYPES = ["Benchmark", "Evaluation", "Other"]

DATA_DIR = Path(__file__).parent.parent / "data"


# ---------------------------------------------------------------------------
# Benchmark / dataset patterns
# ---------------------------------------------------------------------------

# Resources that are almost always *contributed*, never merely used in passing.
_STRONG_NOUN = (
    r"(?:benchmark|benchmark suite|evaluation suite|evaluation benchmark|"
    r"test ?bed|testbed|leaderboard|challenge set|task suite|test suite|"
    r"evaluation harness|evaluation framework|evaluation protocol|arena)"
)
# Resources that are common in ordinary method papers too ("trained on a large
# dataset", "a corpus of text"), so they need corroborating evidence.
_WEAK_NOUN = r"(?:data ?set|corpus|corpora|data collection|annotated collection)"

_ANY_NOUN = rf"(?:{_STRONG_NOUN}|{_WEAK_NOUN})"

_RELEASE_VERB = (
    r"(?:introduce|present|propose|release|construct|curate|build|develop|"
    r"create|collect|assemble|contribute|open.?source|publish|"
    r"design and (?:release|build))"
)

# "we introduce … <resource>" inside one sentence.
_B_INTRO_STRONG = re.compile(
    rf"\b(?:we|this (?:paper|work|study)|our (?:paper|work))\b[^.]{{0,80}}?"
    rf"\b{_RELEASE_VERB}\w*\b[^.]{{0,90}}?\b{_STRONG_NOUN}s?\b",
    re.I,
)
_B_INTRO_WEAK = re.compile(
    rf"\b(?:we|this (?:paper|work|study)|our (?:paper|work))\b[^.]{{0,80}}?"
    rf"\b{_RELEASE_VERB}\w*\b[^.]{{0,90}}?\b(?:new |novel |large.scale |curated |"
    rf"annotated |human.annotated |synthetic |public )+{_WEAK_NOUN}s?\b",
    re.I,
)
# Apposition: "…, a comprehensive benchmark for X"
_B_APPOS = re.compile(
    rf"[,:—-]\s*(?:a|an|the)\s+(?:new |novel |first |large.scale |comprehensive |"
    rf"unified |challenging |curated |public |open |diverse |systematic |"
    rf"standardized |realistic |holistic |fine.grained |multilingual |"
    rf"multimodal )*{_STRONG_NOUN}s?\b",
    re.I,
)
_B_OWNED = re.compile(rf"\bour (?:new |proposed )?{_STRONG_NOUN}s?\b", re.I)
_B_FIRST = re.compile(rf"\bthe first {_ANY_NOUN}s?\b", re.I)
_B_TO_EVAL = re.compile(
    rf"\b{_ANY_NOUN}s?\b[^.]{{0,60}}?\b(?:to|that|designed to|which)\s+"
    rf"(?:systematically\s+)?(?:assess|evaluat\w*|measur\w*|benchmark\w*|"
    rf"test|probe|diagnos\w*|stress.test)\w*\b",
    re.I,
)

# Title signals.
_B_TITLE_NAME = re.compile(
    r"(?:^|[\s\-–—:(])[A-Za-z0-9\-]*bench\b|\bbenchmark(?:s|ing)?\b|"
    r"\btest ?bed\b|\bleaderboard\b|\barena\b|\bevaluation suite\b",
    re.I,
)
_B_TITLE_WEAK = re.compile(r"\bdata ?set\b|\bcorpus\b|\bsuite\b", re.I)

# "…, a multimodal dataset of / for / containing …" — apposition around a weak noun.
_B_APPOS_WEAK = re.compile(
    rf"[,:—-]\s*(?:a|an|the)\s+(?:\w+[\s-]+){{0,4}}{_WEAK_NOUN}s?\b\s*"
    rf"(?:of|for|with|containing|comprising|consisting|spanning|covering)\b",
    re.I,
)

# Phrases where a resource noun is *used*, not *released*.
_B_USAGE = re.compile(
    r"\b(?:on|across|over|using|with|from|against)\s+(?:\w+[\s-]+){0,3}"
    r"(?:standard|existing|public|popular|common|widely.used|established|"
    r"real.world|synthetic|downstream|diverse|challenging|\d+)?\s*"
    r"benchmarks?\b"
    r"|\bbenchmark (?:dataset|task|problem|suite|environment|instance)s?\b"
    r"|\bexisting benchmarks?\b|\bprior benchmarks?\b|\bstandard benchmarks?\b"
    r"|\bbenchmark(?:ed|ing)? against\b",
    re.I,
)


# ---------------------------------------------------------------------------
# Evaluation / study patterns
# ---------------------------------------------------------------------------

_E_TITLE = re.compile(
    r"^(?:[^:]{0,60}:\s*)?(?:an?\s+|the\s+)?(?:empirical|systematic|comparative|"
    r"large.scale|comprehensive|in.depth|critical|careful|rigorous)\s+"
    r"(?:study|analysis|evaluation|investigation|comparison|assessment|review|"
    r"look|examination)\b"
    r"|^(?:evaluating|assessing|measuring|auditing|quantifying|benchmarking|"
    r"revisiting|re.examining|dissecting|probing|characterizing|characterising|"
    r"investigating|examining|comparing|stress.testing)\b"
    r"|^(?:how|do|does|are|is|can|should|why|what|when|which)\b[^?]{0,110}\?"
    r"|\ban? (?:empirical|systematic|comparative|large.scale|comprehensive) "
    r"(?:study|analysis|evaluation|investigation|comparison)\b"
    r"|\b(?:evaluation|assessment|audit|analysis) of\b|\bcase study\b"
    r"|\bhow (?:well|much|far|good|robust|reliable|faithful|sensitive)\b",
    re.I,
)
_E_STRONG = re.compile(
    r"\bwe (?:systematically|comprehensively|empirically|critically|extensively|"
    r"carefully|rigorously)?\s*(?:evaluate|assess|audit|benchmark|compare|"
    r"measure|quantify|examine|investigate|analyz\w*|analys\w*)\b"
    r"|\bempirical (?:study|analysis|investigation|evaluation|assessment)\b"
    r"|\bsystematic (?:study|analysis|evaluation|review|investigation|comparison)\b"
    r"|\bcomprehensive (?:study|analysis|evaluation|comparison|assessment)\b"
    r"|\blarge.scale (?:study|analysis|evaluation|comparison)\b"
    r"|\bwe conduct (?:a|an|the|our)[^.]{0,40}(?:study|survey|analysis|evaluation|"
    r"investigation|comparison|audit)\b"
    r"|\bwe perform (?:a|an|the)[^.]{0,40}(?:study|analysis|investigation|evaluation)\b"
    r"|\bwe (?:present|provide|report) (?:a|an|the)[^.]{0,40}(?:study|analysis|"
    r"evaluation|comparison|audit)\b",
    re.I,
)
_E_MEDIUM = re.compile(
    r"\bour (?:analysis|study|experiments|findings|results|evaluation|audit) "
    r"(?:reveal|show|indicate|suggest|demonstrate|highlight)\w*"
    r"|\bwe study (?:how|why|whether|what|when|the extent)\b"
    r"|\bto (?:better )?understand (?:how|why|whether|what|when)\b"
    r"|\blittle is known\b|\bremains (?:poorly |largely )?(?:understood|unclear|"
    r"unexplored|underexplored)\b"
    r"|\bwe ask (?:whether|how|why)\b"
    r"|\bacross \d+ (?:models|llms|datasets|benchmarks|tasks|settings|"
    r"architectures|domains)\b"
    r"|\bwe (?:test|probe|study|evaluate) \d+ \w+"
    r"|\bthis (?:paper|work|study) (?:studies|investigates|examines|analyzes|"
    r"analyses|evaluates|revisits|characterizes)\b",
    re.I,
)
# A paper whose core contribution is a new artefact is not an evaluation paper.
_E_METHOD_BLOCK = re.compile(
    r"\bwe (?:propose|introduce|present|develop|design|devise)\b[^.]{0,90}?"
    r"\b(?:method|algorithm|framework|model|architecture|approach|technique|"
    r"objective|loss|regularizer|estimator|network|module|mechanism|policy|"
    r"solver|sampler|strategy|pipeline|system|layer|operator|scheme|remedy|"
    r"procedure|criterion|kernel|transform|decoder|encoder|agent|controller)\b"
    r"|\bour (?:method|algorithm|model|framework|approach|architecture) "
    r"(?:achieves|outperforms|improves|surpasses|attains|reduces|scales)\b",
    re.I,
)
_E_THEORY_BLOCK = re.compile(
    r"\bwe prove\b|\btheorem\b|\blemma\b|\bcorollary\b|"
    r"\bwe derive (?:a|an|the|tight|closed)\b|\bregret bound\b|"
    r"\bconvergence rate\b|\bsample complexity\b|\bminimax\b",
    re.I,
)

# ICML primary areas are snake_case, e.g. "general_machine_learning->evaluation".
_PA_EVAL = re.compile(r"->evaluation$|\bevaluation\b", re.I)


# ---------------------------------------------------------------------------
# Scorer
# ---------------------------------------------------------------------------

def score(title: str, abstract: str, primary_area: str = "", keywords: str = "") -> dict:
    t = title or ""
    a = abstract or ""
    pa = primary_area or ""
    kw = keywords or ""
    s = {k: 0.0 for k in PAPER_TYPES}

    # ---- Benchmark ----
    title_bench = bool(_B_TITLE_NAME.search(t)) or bool(_B_TITLE_WEAK.search(t))
    method_present = bool(_E_METHOD_BLOCK.search(a))
    intro_strong = bool(_B_INTRO_STRONG.search(a))
    intro_weak = bool(_B_INTRO_WEAK.search(a))

    if _B_TITLE_NAME.search(t):
        s["Benchmark"] += 4.0
    if _B_TITLE_WEAK.search(t):
        s["Benchmark"] += 2.0
    if intro_strong:
        s["Benchmark"] += 4.0
    elif intro_weak:
        s["Benchmark"] += 2.5
    if _B_APPOS.search(a):
        s["Benchmark"] += 2.5
    elif _B_APPOS_WEAK.search(a):
        s["Benchmark"] += 2.0
    if _B_OWNED.search(a):
        s["Benchmark"] += 2.0
    if _B_FIRST.search(a):
        s["Benchmark"] += 2.0
    if _B_TO_EVAL.search(a):
        s["Benchmark"] += 1.5
    if _B_TITLE_NAME.search(kw) or _B_TITLE_WEAK.search(kw):
        s["Benchmark"] += 1.0
    # "evaluated on standard benchmarks" is a usage phrase, not a contribution —
    # but a paper whose *title* names a benchmark is allowed to talk about them.
    if not title_bench:
        s["Benchmark"] -= min(len(_B_USAGE.findall(a)), 2) * 2.0
    # A method paper that also ships a benchmark is still a method paper, unless
    # the title itself advertises the resource.
    if method_present and not title_bench:
        s["Benchmark"] -= 5.0

    # ---- Evaluation ----
    title_eval = bool(_E_TITLE.search(t))
    n_strong = min(len(_E_STRONG.findall(a)), 3)
    n_medium = min(len(_E_MEDIUM.findall(a)), 3)

    if title_eval:
        s["Evaluation"] += 3.0
    s["Evaluation"] += n_strong * 2.0
    s["Evaluation"] += n_medium * 1.0
    if _PA_EVAL.search(pa):
        s["Evaluation"] += 2.0
    if method_present:
        s["Evaluation"] -= 3.0
    if _E_THEORY_BLOCK.search(a):
        s["Evaluation"] -= 2.5
    if title_bench:
        s["Evaluation"] -= 2.0
    # A title signal on its own (e.g. "Rethinking X") is not enough — the
    # abstract has to actually read like a study.
    if title_eval and n_strong == 0 and n_medium == 0:
        s["Evaluation"] -= 2.0

    return s


def classify(title: str, abstract: str, primary_area: str = "", keywords: str = "") -> str:
    s = score(title, abstract, primary_area, keywords)

    if s["Benchmark"] >= 5.0:
        return "Benchmark"
    if s["Evaluation"] >= 4.0:
        return "Evaluation"
    if s["Benchmark"] >= 3.5 and s["Benchmark"] > s["Evaluation"]:
        return "Benchmark"

    return "Other"


# ---------------------------------------------------------------------------
# Standalone re-labelling of an existing spreadsheet
# ---------------------------------------------------------------------------

def main():
    import openpyxl

    path = Path(sys.argv[1]) if len(sys.argv) > 1 else DATA_DIR / "ICML_2026.xlsx"
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    header = {c.value: i for i, c in enumerate(ws[1])}

    for name in ("Title", "Abstract", "Paper Type"):
        if name not in header:
            print(f"ERROR: column {name!r} not found in {path}")
            sys.exit(1)

    col = header["Paper Type"] + 1
    counts = Counter()
    for row in ws.iter_rows(min_row=2):
        label = classify(
            str(row[header["Title"]].value or ""),
            str(row[header["Abstract"]].value or ""),
            str(row[header["Primary Area"]].value or "") if "Primary Area" in header else "",
            str(row[header["Keywords"]].value or "") if "Keywords" in header else "",
        )
        ws.cell(row=row[0].row, column=col, value=label)
        counts[label] += 1

    wb.save(path)
    total = sum(counts.values())
    print(f"Re-labelled {total} papers in {path}")
    for k in PAPER_TYPES:
        print(f"  {k:<11} {counts[k]:5d}  ({counts[k] / total:.1%})")


if __name__ == "__main__":
    main()
