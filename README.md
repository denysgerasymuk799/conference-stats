# Conference Stats

Data-driven visualizations of accepted paper trends at top ML conferences — **ICLR 2026** and **ICML 2026**.

All charts are generated from raw OpenReview metadata — no manual labeling, no external APIs at inference time. The pipeline fetches paper titles, abstracts, keywords, and primary areas, then classifies each paper with a rule-based multi-signal scorer to produce publication-ready PNG charts.

---

## ICLR 2026 — Charts

### Paper Growth (2022 – 2026)

Accepted submissions have grown ~3× over five years, from ~1,100 in 2022 to over 4,000 in 2026. AI Safety & Alignment papers have kept pace, now exceeding 1,000 accepted papers.

![Number of papers](iclr_2026/plots/number_of_papers.png)

---

### Contribution Archetypes — All Papers

Each accepted paper is classified into one of five contribution types: **New Method**, **Empirical Study**, **Theory**, **Benchmark / Dataset**, or **Other**.

![Contribution archetypes](iclr_2026/plots/contribution_archetypes.png)

---

### Contribution Archetypes — AI Safety & Alignment Papers

Same classification, restricted to the AI safety subset.

![Contribution archetypes – AI safety](iclr_2026/plots/contribution_archetypes_ai_safety.png)

---

### Trending Topics — All Papers

Topic counts for ICLR 2026. A paper can match multiple topics; the denominator is the total number of accepted papers.

![Trending topics](iclr_2026/plots/trending_topics.png)

---

### Trending Topics — AI Safety & Alignment Papers

Filtered to papers whose title, keywords, or primary area contain safety-related terms (adversarial, alignment, fairness, privacy, hallucination, etc.).

![Trending topics – AI safety](iclr_2026/plots/trending_topics_ai_safety.png)

---

### Trending Topics — Healthcare & Life Sciences Papers

Filtered to papers whose title, keywords, or primary area contain healthcare-related terms (clinical, medical, biological, genomic, etc.).

![Trending topics – healthcare](iclr_2026/plots/trending_topics_healthcare.png)

---

## How It Works

### 1 — Data Collection (`fetch_iclr_openreview.py`)

Paper metadata is fetched from the [OpenReview API](https://openreview.net):
- **2022 – 2023**: OpenReview API v1
- **2024 – 2026**: OpenReview API v2

Each year's accepted papers are saved as an Excel file in `iclr_2026/data/`.

### 2 — Archetype Classification (`classify_archetypes.py`)

A multi-signal rule-based scorer assigns each paper a floating-point score across five archetypes. The classifier uses:

| Signal | How it works |
|---|---|
| **Strong regex hits** | Theorem/lemma/corollary for Theory; "we introduce … new benchmark" for Benchmark |
| **Verb–noun proximity** | "we propose" within 130 chars of a method noun, with no empirical noun intervening → New Method |
| **Primary area boost** | "datasets and benchmarks" primary area → hard-assign Benchmark |
| **Penalty terms** | Performance claims ("outperform", "state-of-the-art") penalize Empirical Study to avoid ablation papers being miscounted |

Results are written to `iclr_2026/data/archetypes_classified.json` and read by the plot script — no re-classification at plot time.

### 3 — Plot Generation (`generate_plot.py`)

Each chart is a standalone HTML page rendered to a 1980 × 1256 px PNG (2× device pixel ratio) via a headless Chromium browser (Playwright). Chart.js and chartjs-plugin-datalabels are vendored locally so the renderer works fully offline.

```
python iclr_2026/src/classify_archetypes.py          # pre-compute archetypes JSON
python iclr_2026/src/generate_plot.py <plot_name>    # render one chart
```

Available plot names:

| Name | Description |
|---|---|
| `number_of_papers` | Line chart — accepted papers per year |
| `contribution_archetypes` | Stacked bar — archetype breakdown, all papers |
| `contribution_archetypes_ai_safety` | Stacked bar — archetype breakdown, AI safety subset |
| `trending_topics` | Table — top topics, all ICLR 2026 papers |
| `trending_topics_ai_safety` | Table — top topics, AI safety papers |
| `trending_topics_healthcare` | Table — top topics, healthcare papers |

---

## ICML 2026 — Dataset

`icml_2026/data/ICML_2026.xlsx` holds all **6,341 accepted ICML 2026 main-conference papers** (Seoul, July 6–11 2026), in the same shape as the ICLR file plus two extra columns.

| Column | Notes |
|---|---|
| `Decision` | `Oral` (159) / `Spotlight` (377) / `Poster` (5,805) — presentation format in the official program |
| `Spotlight` | `Yes` for the 536 papers whose OpenReview venue is *ICML 2026 spotlight* (they present as either an oral or a spotlight talk) |
| `Title`, `Authors`, `Abstract`, `Primary Area`, `Keywords` | OpenReview camera-ready metadata |
| `Institutions` | Per-author affiliation, aligned with `Authors` (98% of authors resolved) |
| `Paper Type` | `Benchmark` (342) / `Evaluation` (99) / `Other` (5,900) |
| `Date (KST)`, `Start (KST)`, `End (KST)`, `Location`, `Poster #`, `Session` | Poster-session placement in Seoul local time; blank for the 725 papers with no in-person slot |
| `OpenReview URL`, `Virtual Site URL` | Links |

### Data Collection (`fetch_icml_openreview.py`)

```bash
python icml_2026/src/fetch_icml_openreview.py 2026            # writes data/ICML_2026.xlsx
python icml_2026/src/fetch_icml_openreview.py 2026 --refresh  # ignore data/cache/
```

Two public sources are merged on the OpenReview forum id:

| Source | Provides | Access note |
|---|---|---|
| `api2.openreview.net/notes/search` | Title, authors, abstract, primary area, keywords, venue label | The plain `/notes` endpoint now sits behind a Cloudflare Turnstile challenge for anonymous clients, so the script sweeps the still-open search endpoint over a stop-word list and unions the results (`a` + `the` + `we` already cover all 6,341 papers) |
| `icml.cc` virtual-site program JSON | Institutions, presentation type, room, poster number, session, times | OpenReview serves no affiliations without a login, so institutions come from the conference program. The live static snapshot is currently truncated to its first API page and the API behind it needs an icml.cc account, so the script falls back to the newest **complete** snapshot in the Wayback Machine and layers the live file on top |

The two sources agree exactly — the 6,341 accepted OpenReview notes and the 6,341 conference-track program entries are the same set, and their spotlight designations match on every paper.

### Paper-Type Classification (`classify_paper_type.py`)

The same style of offline multi-signal scorer used for ICLR archetypes, run over the title and abstract:

| Label | Assigned when |
|---|---|
| **Benchmark** | The headline contribution is a new benchmark, dataset, testbed or evaluation suite — a `…Bench`-style title, or a release verb ("we introduce / release / curate") near a resource noun |
| **Evaluation** | The headline contribution is measuring existing models — study/audit/comparison titles plus corroborating abstract evidence |
| **Other** | Everything else (new methods, theory, applications) |

Two rules do most of the disambiguation work: phrases that merely *use* benchmarks ("on standard benchmarks") are penalised rather than counted, and a paper that ships a benchmark alongside a new method is classified by its title — a method paper stays `Other`. The classifier favours precision over recall, so a small number of study papers land in `Other`. It can be re-run on its own to relabel the sheet in place:

```bash
python icml_2026/src/classify_paper_type.py
```

---

## Setup

```bash
pip install -r requirements.txt
playwright install chromium
```

**`requirements.txt`**
```
openreview-py>=2.2.0
pandas>=3.0.2
openpyxl>=3.1.5
playwright>=1.44.0
```

---

## Data Source

All paper metadata is sourced from [OpenReview.net](https://openreview.net) under its public API. No paper PDFs are downloaded or processed.
