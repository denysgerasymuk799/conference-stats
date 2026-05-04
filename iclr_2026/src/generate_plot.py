"""
Generate LinkedIn-ready charts from ICLR acceptance data.

Usage:
    python src/generate_plot.py <plot_name>

Available plots:
    iclr_trends  – Total, Healthcare, and AI Safety paper counts per year (2022–2026)
"""

import json
import re
import sys
import tempfile
import urllib.request
from pathlib import Path

import openpyxl
from playwright.sync_api import sync_playwright

# Vendor JS is cached in src/vendor/ (downloaded on first run)
_VENDOR_DIR = Path(__file__).parent / "vendor"
_VENDOR_DIR.mkdir(exist_ok=True)

_VENDOR = {
    "chartjs.min.js":    "https://cdn.jsdelivr.net/npm/chart.js@4.4.4/dist/chart.umd.min.js",
    "datalabels.min.js": "https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.2.0/dist/chartjs-plugin-datalabels.min.js",
}


def _vendor_js(filename: str) -> str:
    path = _VENDOR_DIR / filename
    if not path.exists():
        print(f"  Downloading {filename} …")
        urllib.request.urlretrieve(_VENDOR[filename], path)
    return path.read_text(encoding="utf-8")


DATA_DIR  = Path(__file__).parent.parent / "data"
PLOTS_DIR = Path(__file__).parent.parent / "plots"
PLOTS_DIR.mkdir(exist_ok=True)

# ---------------------------------------------------------------------------
# Paper categorisation
# Searches Title + Keywords + Primary Area (case-insensitive).
# ---------------------------------------------------------------------------

HEALTHCARE_TERMS = [
    # Clinical / medical
    "health", "medic", "clinic", "patient", "disease", "diagnos",
    "treatment", "therapy", "drug", "cancer", "tumor", "tumour",
    "hospital", "biomedic", "ehr", "radiology", "pathology",
    "covid", "pandemic", "surgical", "surgery", "pharmaceu",
    # Biology / life sciences
    # (also catches primary areas like "applications to physical sciences
    # (physics, chemistry, biology, etc.)" and "neuroscience & cognitive science")
    "biology", "biological", "bioinformatic", "protein", "molecular",
    "molecule", "genomic", "genome", "neurosci",
    # Biomedical signals & imaging modalities (from ICLR 2026 plan titles)
    "eeg", "ecg", "electrocard", "biosignal", "fmri", "mri",
    # Conditions, specialties & care settings
    "epilepsy", "cardiac", "cardio", "ultrasound", "physiolog",
    "anatomy", "anatomical", "mental health", "intensive care",
    "glucose", "respiratory", "wearable", "sleep", "spine", "spinal",
    "survival analysis",
]

AI_SAFETY_TERMS = [
    # Explicitly requested
    "privacy", "jailbreak", "explainab",
    # Core safety / alignment
    "ai safety", "alignment", "adversarial", "fairness", "trustworthy",
    "interpretab", "hallucination", "red team", "value alignment",
    "reward hacking", "decepti", "safety", "toxic", "harmful",
    # From primary area names
    # ("societal considerations…", "alignment, fairness, safety, privacy…",
    #  "Social Aspects of ML (eg, AI safety, fairness, privacy…)", etc.)
    "societal", "ethic",
    # Attack vectors & defences (from ICLR 2026 plan titles)
    "unlearn", "backdoor", "bias", "watermark", "poison", "attack", "vulnerab",
]


def _search_text(row: tuple, headers: dict) -> str:
    parts = []
    for col in ("Title", "Keywords", "Primary Area"):
        idx = headers.get(col)
        if idx is not None:
            parts.append(str(row[idx] or "").lower())
    return " ".join(parts)


def count_categories(path: Path) -> dict:
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    headers = {cell.value: i for i, cell in enumerate(ws[1])}
    total = healthcare = ai_safety = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        total += 1
        text = _search_text(row, headers)
        if any(t in text for t in HEALTHCARE_TERMS):
            healthcare += 1
        if any(t in text for t in AI_SAFETY_TERMS):
            ai_safety += 1
    return {"total": total, "healthcare": healthcare, "ai_safety": ai_safety}


# ---------------------------------------------------------------------------
# Contribution archetype classification
# Each paper is assigned exactly ONE archetype (priority-based):
#   Benchmark/Dataset → Theory → New Method → Empirical Study → Other
#
# Detection uses the abstract text (case-insensitive regex).
# Benchmark detection requires a creation verb AND a benchmark noun to appear
# within 100 characters of each other, preventing false positives from papers
# that merely *evaluate on* existing benchmarks.
# ---------------------------------------------------------------------------

_ARCH_BENCH_VERB = re.compile(
    r'\b(we introduce|we release|we create|we collect|we curate|we build|we construct)\b',
    re.I,
)
_ARCH_BENCH_NOUN = re.compile(
    r'\b(benchmark|dataset|testbed|evaluation suite|test suite|corpus|leaderboard)\b',
    re.I,
)
_ARCH_THEORY = re.compile(
    r'\b(theorem|we prove|provably|convergence guarantee|generalization bound|'
    r'regret bound|sample complexity|formal guarantee|information.theoretic|'
    r'pac learning|minimax optimal|tight bound|lower bound on|we establish)\b',
    re.I,
)
# Unambiguous method signals — no proximity check needed.
_ARCH_METHOD_DIRECT = re.compile(
    r'\b(our (method|algorithm|model|architecture|network|system|technique)|'
    r'we (fine.tune|finetune|pretrain|pre.train))\b',
    re.I,
)
# Verbs that may introduce a method OR an empirical study — need proximity validation.
_ARCH_METHOD_VERB = re.compile(
    r'\b(we propose|we introduce|we develop|we design|we present)\b',
    re.I,
)
# Method-type nouns that confirm the preceding verb is about a technical artifact.
_ARCH_METHOD_NOUN = re.compile(
    r'\b(method|algorithm|model|architecture|network|system|technique|'
    r'framework|module|mechanism|policy|agent|regularizer|loss function)\b',
    re.I,
)
# Empirical-type nouns: if one appears before the method noun in the lookahead
# window, the phrase is about a study/evaluation, not a new artifact.
_ARCH_EMPIRICAL_CONTEXT = re.compile(
    r'\b(study|studies|analysis|analyses|evaluat\w+|protocol|survey|'
    r'investigation|examination|audit|measurement|review)\b',
    re.I,
)
_ARCH_EMPIRICAL = re.compile(
    r'\b(we study|we investigate|we analyze|we examine|we characterize|we audit|'
    r'empirical study|empirical analysis|we find that|we observe that|'
    r'we show that|we demonstrate that)\b',
    re.I,
)


def _has_method_signal(a: str) -> bool:
    if _ARCH_METHOD_DIRECT.search(a):
        return True
    for m in _ARCH_METHOD_VERB.finditer(a):
        snippet = a[m.end(): min(len(a), m.end() + 120)]
        mn = _ARCH_METHOD_NOUN.search(snippet)
        if mn is None:
            continue
        en = _ARCH_EMPIRICAL_CONTEXT.search(snippet)
        # If an empirical noun appears before the method noun → evaluation framing
        if en and en.start() < mn.start():
            continue
        return True
    return False

ARCHETYPES = [
    "New Method",
    "Empirical Study",
    "Theory",
    "Benchmark / Dataset",
    "Other",
]

# Colors in the same order as ARCHETYPES (used in both legend and bars)
ARCHETYPE_COLORS = ["#5046e5", "#f97316", "#0891b2", "#10b981", "#9ca3af"]


def classify_archetype(abstract: str, primary_area: str) -> str:
    a  = abstract or ""
    pa = (primary_area or "").lower()

    # 1. Primary area explicitly marks dataset/benchmark papers
    if "datasets and benchmarks" in pa:
        return "Benchmark / Dataset"

    # 2. Theory
    if _ARCH_THEORY.search(a):
        return "Theory"

    # 3. New Method — takes priority over benchmark signals.
    #    "we propose an evaluation protocol" does NOT count; verb must be
    #    followed by a method-type noun before any empirical-type noun.
    if _has_method_signal(a):
        return "New Method"

    # 4. Benchmark / Dataset — only reaches here if no method signals
    for m in _ARCH_BENCH_VERB.finditer(a):
        snippet = a[max(0, m.start() - 100): m.start() + 150]
        if _ARCH_BENCH_NOUN.search(snippet):
            return "Benchmark / Dataset"

    # 5. Empirical Study
    if _ARCH_EMPIRICAL.search(a):
        return "Empirical Study"

    return "Other"


def count_archetypes(path: Path, filter_terms: list | None = None) -> dict:
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    headers = {cell.value: i for i, cell in enumerate(ws[1])}
    abs_idx = headers.get("Abstract")
    pa_idx  = headers.get("Primary Area")
    counts  = {a: 0 for a in ARCHETYPES}
    total   = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        if filter_terms is not None:
            search = _search_text(row, headers)
            if not any(t in search for t in filter_terms):
                continue
        total += 1
        abstract     = str(row[abs_idx] or "") if abs_idx is not None else ""
        primary_area = str(row[pa_idx]  or "") if pa_idx  is not None else ""
        counts[classify_archetype(abstract, primary_area)] += 1
    return {"total": total, **counts}


# ---------------------------------------------------------------------------
# HTML / Chart.js template
# Clean white style inspired by modern academic chart design.
# Three lines, one Y axis, no subtitle.
# Rendered at 2× device pixel ratio for crisp LinkedIn resolution.
# ---------------------------------------------------------------------------

_HTML_TEMPLATE = """\
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  html, body {{
    width: 990px;
    height: 628px;
    overflow: hidden;
    background: #ffffff;
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Helvetica Neue', Arial, sans-serif;
  }}
  .card {{
    width: 990px;
    height: 628px;
    padding: 44px 56px 36px 56px;
    background: #ffffff;
    display: flex;
    flex-direction: column;
  }}
  .title {{
    font-size: 30px;
    font-weight: 800;
    color: #111827;
    line-height: 1.2;
    letter-spacing: -0.5px;
    margin-bottom: 22px;
  }}
  .chart-wrap {{
    flex: 1;
    position: relative;
    min-height: 0;
  }}
  .rule {{
    height: 1px;
    background: #e5e7eb;
    margin-top: 14px;
    margin-bottom: 10px;
  }}
  .footer {{
    font-size: 12.5px;
    color: #6b7280;
    line-height: 1.55;
  }}
  .footer strong {{ font-weight: 700; color: #374151; }}
</style>
<script>{chartjs_js}</script>
<script>{datalabels_js}</script>
</head>
<body>
<div class="card">
  <div class="title">{title}</div>
  <div class="chart-wrap">
    <canvas id="myChart"></canvas>
  </div>
  <div class="rule"></div>
  <div class="footer">{footer}</div>
</div>
<script>
  Chart.register(ChartDataLabels);

  const YEARS  = {years_json};
  const TOTAL  = {total_json};
  const HEALTH = {health_json};
  const SAFETY = {safety_json};
  const last   = YEARS.length - 1;

  // Colour palette
  const C_TOTAL  = '#5046e5';   // indigo-violet (matches reference purple)
  const C_HEALTH = '#0891b2';   // cyan-600
  const C_SAFETY = '#f97316';   // orange-500

  const ctx = document.getElementById('myChart').getContext('2d');

  new Chart(ctx, {{
    type: 'line',
    data: {{
      labels: YEARS,
      datasets: [
        {{
          label: 'Total Accepted',
          data: TOTAL,
          borderColor: C_TOTAL,
          backgroundColor: 'transparent',
          borderWidth: 3,
          pointRadius: 7,
          pointHoverRadius: 9,
          pointBackgroundColor: C_TOTAL,
          pointBorderColor: '#fff',
          pointBorderWidth: 2.5,
          tension: 0.3,
          datalabels: {{
            display: ctx => ctx.dataIndex === last,
            anchor: 'center',
            align: 'right',
            offset: 14,
            color: C_TOTAL,
            font: {{ size: 15, weight: '800' }},
            formatter: v => v.toLocaleString(),
          }},
        }},
        {{
          label: 'AI Safety & Alignment',
          data: SAFETY,
          borderColor: C_SAFETY,
          backgroundColor: 'transparent',
          borderWidth: 3,
          pointRadius: 7,
          pointHoverRadius: 9,
          pointBackgroundColor: C_SAFETY,
          pointBorderColor: '#fff',
          pointBorderWidth: 2.5,
          tension: 0.3,
          datalabels: {{
            display: ctx => ctx.dataIndex === last,
            anchor: 'center',
            align: 'right',
            offset: 14,
            color: C_SAFETY,
            font: {{ size: 15, weight: '800' }},
            formatter: v => v.toLocaleString(),
          }},
        }},
        {{
          label: 'Healthcare & Life Sciences',
          data: HEALTH,
          borderColor: C_HEALTH,
          backgroundColor: 'transparent',
          borderWidth: 3,
          pointRadius: 7,
          pointHoverRadius: 9,
          pointBackgroundColor: C_HEALTH,
          pointBorderColor: '#fff',
          pointBorderWidth: 2.5,
          tension: 0.3,
          datalabels: {{
            display: ctx => ctx.dataIndex === last,
            anchor: 'center',
            align: 'right',
            offset: 14,
            color: C_HEALTH,
            font: {{ size: 15, weight: '800' }},
            formatter: v => v.toLocaleString(),
          }},
        }},
      ],
    }},
    options: {{
      responsive: true,
      maintainAspectRatio: false,
      animation: false,
      layout: {{ padding: {{ top: 16, right: 90, bottom: 4, left: 0 }} }},
      scales: {{
        y: {{
          beginAtZero: true,
          ticks: {{
            color: '#9ca3af',
            font: {{ size: 14 }},
            maxTicksLimit: 7,
            callback: v => v >= 1000 ? (v / 1000).toFixed(0) + 'k' : v,
          }},
          grid: {{ color: '#f3f4f6', lineWidth: 1.5 }},
          border: {{ display: false, dash: [0] }},
        }},
        x: {{
          ticks: {{
            color: '#111827',
            font: {{ size: 16, weight: '700' }},
            padding: 8,
          }},
          grid: {{ display: false }},
          border: {{ display: false }},
        }},
      }},
      plugins: {{
        legend: {{
          position: 'top',
          align: 'start',
          labels: {{
            color: '#374151',
            font: {{ size: 14.5, weight: '600' }},
            padding: 24,
            usePointStyle: true,
            pointStyleWidth: 12,
          }},
        }},
        tooltip: {{
          backgroundColor: '#1f2937',
          titleFont: {{ size: 13, weight: '700' }},
          bodyFont: {{ size: 13 }},
          padding: 12,
          callbacks: {{
            label: ctx => '  ' + ctx.dataset.label + ': ' + ctx.parsed.y.toLocaleString(),
          }},
        }},
      }},
    }},
  }});
</script>
</body>
</html>
"""


def build_html(title: str, footer: str,
               years: list, total: list, health: list, safety: list) -> str:
    return _HTML_TEMPLATE.format(
        title=title,
        footer=footer,
        years_json=json.dumps(years),
        total_json=json.dumps(total),
        health_json=json.dumps(health),
        safety_json=json.dumps(safety),
        chartjs_js=_vendor_js("chartjs.min.js"),
        datalabels_js=_vendor_js("datalabels.min.js"),
    )


# ---------------------------------------------------------------------------
# Playwright screenshot  (2× DPI for crisp LinkedIn resolution)
# ---------------------------------------------------------------------------

def screenshot_html(html: str, output_path: Path):
    with sync_playwright() as pw:
        browser = pw.chromium.launch()
        page = browser.new_page(
            viewport={"width": 990, "height": 628},
            device_scale_factor=2,
        )
        with tempfile.NamedTemporaryFile(suffix=".html", mode="w",
                                         encoding="utf-8", delete=False) as f:
            f.write(html)
            tmp_path = f.name
        page.goto(f"file://{tmp_path}", wait_until="load")
        page.screenshot(path=str(output_path), clip={"x": 0, "y": 0, "width": 990, "height": 628})
        browser.close()
    Path(tmp_path).unlink(missing_ok=True)


# ---------------------------------------------------------------------------
# Plot definitions
# ---------------------------------------------------------------------------

YEARS = [2022, 2023, 2024, 2025, 2026]


def plot_iclr_trends(output_path: Path):
    print("Loading data …")
    stats = {}
    for year in YEARS:
        path = DATA_DIR / f"ICLR_{year}.xlsx"
        if not path.exists():
            raise FileNotFoundError(f"Missing: {path}")
        print(f"  ICLR {year} …", end=" ", flush=True)
        stats[year] = count_categories(path)
        s = stats[year]
        print(f"total={s['total']:5d}  healthcare={s['healthcare']:4d}  ai_safety={s['ai_safety']:4d}")

    years_labels = [str(y) for y in YEARS]
    totals = [stats[y]["total"]      for y in YEARS]
    health = [stats[y]["healthcare"] for y in YEARS]
    safety = [stats[y]["ai_safety"]  for y in YEARS]

    html = build_html(
        title="ICLR Accepted Papers 2022–2026",
        footer=(
            "<strong>Healthcare & Life Sciences:</strong> papers with medical, clinical, biological, or life-science "
            "keywords in title, keywords, or primary area.<br>"
            "<strong>AI Safety & Alignment:</strong> papers with safety, alignment, fairness, interpretability, "
            "or adversarial keywords. Source: OpenReview."
        ),
        years=years_labels,
        total=totals,
        health=health,
        safety=safety,
    )

    print("Rendering chart …")
    screenshot_html(html, output_path)
    print(f"Saved → {output_path}")


# ---------------------------------------------------------------------------
# Stacked bar HTML / Chart.js template  (contribution archetypes)
# ---------------------------------------------------------------------------

_HTML_BAR_TEMPLATE = """\
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  html, body {{
    width: 990px;
    height: 628px;
    overflow: hidden;
    background: #ffffff;
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Helvetica Neue', Arial, sans-serif;
  }}
  .card {{
    width: 990px;
    height: 628px;
    padding: 30px 56px 36px 56px;
    background: #ffffff;
    display: flex;
    flex-direction: column;
  }}
  .title {{
    font-size: 30px;
    font-weight: 800;
    color: #111827;
    line-height: 1.2;
    letter-spacing: -0.5px;
    margin-bottom: 2px;
  }}
  .chart-wrap {{
    flex: 1;
    position: relative;
    min-height: 0;
  }}
  .rule {{
    height: 1px;
    background: #e5e7eb;
    margin-top: 14px;
    margin-bottom: 10px;
  }}
  .footer {{
    font-size: 12.5px;
    color: #6b7280;
    line-height: 1.55;
  }}
  .footer strong {{ font-weight: 700; color: #374151; }}
</style>
<script>{chartjs_js}</script>
<script>{datalabels_js}</script>
</head>
<body>
<div class="card">
  <div class="title">{title}</div>
  <div class="chart-wrap">
    <canvas id="myChart"></canvas>
  </div>
  <div class="rule"></div>
  <div class="footer">{footer}</div>
</div>
<script>
  Chart.register(ChartDataLabels);

  const YEARS    = {years_json};
  const DATASETS = {datasets_json};

  const ctx = document.getElementById('myChart').getContext('2d');

  new Chart(ctx, {{
    type: 'bar',
    data: {{
      labels: YEARS,
      datasets: DATASETS,
    }},
    options: {{
      responsive: true,
      maintainAspectRatio: false,
      animation: false,
      layout: {{ padding: {{ top: 40, right: 10, bottom: 0, left: 0 }} }},
      scales: {{
        x: {{
          stacked: true,
          ticks: {{
            color: '#111827',
            font: {{ size: 16, weight: '700' }},
            padding: 8,
          }},
          grid: {{ display: false }},
          border: {{ display: false }},
        }},
        y: {{
          stacked: true,
          min: 0,
          max: 100,
          ticks: {{
            color: '#9ca3af',
            font: {{ size: 14 }},
            stepSize: 20,
            callback: v => v + '%',
          }},
          grid: {{ color: '#f3f4f6', lineWidth: 1.5 }},
          border: {{ display: false }},
        }},
      }},
      plugins: {{
        legend: {{
          position: 'top',
          align: 'start',
          labels: {{
            color: '#374151',
            font: {{ size: 14, weight: '600' }},
            padding: 20,
            usePointStyle: true,
            pointStyle: 'rect',
            pointStyleWidth: 14,
          }},
        }},
        tooltip: {{
          backgroundColor: '#1f2937',
          titleFont: {{ size: 13, weight: '700' }},
          bodyFont: {{ size: 13 }},
          padding: 12,
          callbacks: {{
            label: ctx => '  ' + ctx.dataset.label + ': ' + ctx.parsed.y.toFixed(1) + '%',
          }},
        }},
        datalabels: {{
          display: ctx => ctx.parsed != null && ctx.parsed.y >= 6,
          color: '#ffffff',
          font: {{ size: 13, weight: '700' }},
          formatter: v => v.toFixed(0) + '%',
          anchor: 'center',
          align: 'center',
        }},
      }},
    }},
  }});
</script>
</body>
</html>
"""


def build_bar_html(title: str, footer: str, years: list,
                   arch_pct: dict) -> str:
    # arch_pct: {archetype_label: [pct_2022, pct_2023, ...]}
    datasets = []
    for label, color in zip(ARCHETYPES, ARCHETYPE_COLORS):
        datasets.append({
            "label":           label,
            "data":            arch_pct[label],
            "backgroundColor": color,
        })
    return _HTML_BAR_TEMPLATE.format(
        title=title,
        footer=footer,
        years_json=json.dumps(years),
        datasets_json=json.dumps(datasets),
        chartjs_js=_vendor_js("chartjs.min.js"),
        datalabels_js=_vendor_js("datalabels.min.js"),
    )


# ---------------------------------------------------------------------------
# Trending topics  (ICLR 2026 — keywords + title matching)
# ---------------------------------------------------------------------------

TOPIC_PATTERNS: list[tuple[str, re.Pattern]] = [
    ("Large Language Models", re.compile(
        r"large language model|(?<!\w)llm(?:s)?(?!\w)|"
        r"instruction (tuning|following)|in.context learning|"
        r"chain.of.thought|autoregressive (model|generation)|"
        r"foundation model|frontier model|"
        r"(?:gpt|gemini|llama|mistral|qwen|phi|falcon)(?:[- _\d]|$)",
        re.I,
    )),
    ("Reinforcement Learning", re.compile(
        r"reinforcement learning|(?<!\w)rl(?!\w)|"
        r"reward (model|function|shaping|learning)|"
        r"policy (gradient|optimization|learning)|"
        r"(?<!\w)(?:ppo|dpo|sac|rlhf|grpo)(?!\w)|actor.critic|"
        r"markov decision|(?<!\w)mdp(?!\w)|"
        r"multi.armed bandit|offline rl|"
        r"value function|temporal difference|imitation learning",
        re.I,
    )),
    ("Diffusion & Flow Models", re.compile(
        r"diffusion (model|process|network|policy)|"
        r"score (matching|based model)|denoising diffusion|"
        r"(?<!\w)(?:ddpm|ddim)(?!\w)|"
        r"flow matching|consistency model|rectified flow",
        re.I,
    )),
    ("Reasoning & Planning", re.compile(
        r"(?:mathematical|logical|commonsense|formal|causal|abstract|llm) reasoning|"
        r"chain.of.thought|theorem prov(?:ing|er)|"
        r"code generation|program synthesis|"
        r"symbolic reasoning|neuro.symbolic|test.time (scaling|compute)|"
        r"planning|multi.step reasoning",
        re.I,
    )),
    ("Multimodal & Vision-Language", re.compile(
        r"multimodal large language|multimodal learning|multi.modal|"
        r"vision.language (model|learning)|"
        r"(?<!\w)vlm(?:s)?(?!\w)|(?<!\w)clip(?!\w)|"
        r"visual question answering|(?<!\w)vqa(?!\w)|"
        r"image.text|text.to.image|image captioning|"
        r"video (generation|understanding|language)|"
        r"image generation|image editing|3d gaussian",
        re.I,
    )),
    ("Interpretability", re.compile(
        r"interpretabilit|mechanistic interpretabilit|"
        r"explainabilit|feature attribution|"
        r"saliency|attention visualization|"
        r"probing|circuit (analysis|discovery)|"
        r"superposition|sparse autoencoder|"
        r"internal representation|concept (bottleneck|learning)",
        re.I,
    )),
    # Uses same terms as number_of_papers so counts are directly comparable.
    # Primary Area searched too (author-supplied taxonomy), matching that plot's logic.
    ("Safety & Alignment", re.compile(
        r"privacy|jailbreak|explainab|ai safety|"
        r"(?<!\w)alignment(?!\w)|adversarial|fairness|trustworthy|"
        r"interpretab|hallucination|red.team|value alignment|"
        r"reward hacking|decepti|(?<!\w)safety(?!\w)|toxic|harmful|"
        r"societal|ethic|unlearn|backdoor|bias|watermark|poison|"
        r"(?<!\w)attack(?!\w)|vulnerab",
        re.I,
    )),
    ("Efficient ML & Quantization", re.compile(
        r"parameter.efficient|(?<!\w)lora(?!\w)|(?<!\w)peft(?!\w)|"
        r"low.rank (adaptation|fine.tuning)|"
        r"quantization|(?:model|network) pruning|"
        r"knowledge distillation|model compression|"
        r"mixture.of.experts|(?<!\w)moe(?!\w)|"
        r"efficient (fine.tuning|training|inference)|"
        r"test.time (adaptation|training)|continual learning",
        re.I,
    )),
    ("Agents & Robotics", re.compile(
        r"(?<!\w)agent(?:s)?(?!\w)|agentic|multi.agent|"
        r"robot(?:ics)? (?:learning|manipulation|navigation|control)|"
        r"embodied (?:ai|intelligence|navigation)|autonomous driving|"
        r"tool (?:use|call(?:ing)?)|(?:long.term|episodic|external|working) memory|"
        r"memory.augmented|llm (?:agent|planning|workflow)|"
        r"agentic (?:workflow|memory|behavior|framework)|"
        r"(?:task|web|gui).?agent",
        re.I,
    )),
    ("Graph & Structured Learning", re.compile(
        r"graph neural (network|learning)|(?<!\w)gnn(?:s)?(?!\w)|"
        r"graph (attention|transformer|convolutional|learning)|"
        r"node (classification|embedding)|link prediction|"
        r"knowledge graph|molecular graph|"
        r"optimal transport",
        re.I,
    )),
    ("Generative Models", re.compile(
        r"generative (model|adversarial|flow)|"
        r"(?<!\w)gan(?:s)?(?!\w)|variational auto.?encoder|(?<!\w)vae(?!\w)|"
        r"normalizing flow|energy.based model|"
        r"neural rendering|neural radiance",
        re.I,
    )),
    ("Representation & Self-supervised", re.compile(
        r"representation learning|self.supervised (learning|pre.?training)|"
        r"contrastive (learning|representation)|"
        r"(?<!\w)(?:simclr|moco|mae|byol)(?!\w)|"
        r"masked (autoencoder|image modeling)|"
        r"unsupervised (pre.?training|representation)|pretraining",
        re.I,
    )),
    ("Uncertainty & Bayesian", re.compile(
        r"uncertainty (quantification|estimation|calibration)|"
        r"bayesian (learning|inference|optimization|neural)|"
        r"conformal (prediction|inference)|calibration|"
        r"distribution shift|out.of.distribution|(?<!\w)ood(?!\w)",
        re.I,
    )),
    ("Federated & Privacy", re.compile(
        r"federated (learning|optimization|averaging)|(?<!\w)fedavg(?!\w)|"
        r"differential privacy|privacy.preserving|"
        r"secure (aggregation|computation)|"
        r"communication.efficient",
        re.I,
    )),
    ("Healthcare & Biology", re.compile(
        r"protein (structure|folding|design|language)|"
        r"molecular (generation|design|property)|"
        r"drug (discovery|design|interaction)|genomic|"
        r"medical (image|imaging|report)|clinical|"
        r"electronic health|biomedical|bioinformatic",
        re.I,
    )),
]


def count_topics(path: Path) -> tuple[int, list[tuple[str, int]]]:
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    h  = {c.value: i for i, c in enumerate(ws[1])}

    def col(row, name):
        idx = h.get(name)
        return str(row[idx] or "") if idx is not None else ""

    counts = {name: 0 for name, _ in TOPIC_PATTERNS}
    total  = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        total += 1
        # Primary Area is author-supplied taxonomy — include alongside title & keywords
        text  = (col(row, "Title") + " " + col(row, "Keywords") + " " + col(row, "Primary Area")).lower()
        for name, pat in TOPIC_PATTERNS:
            if pat.search(text):
                counts[name] += 1

    ranked = sorted(counts.items(), key=lambda x: x[1], reverse=True)
    return total, ranked


# ---------------------------------------------------------------------------
# AI-safety-specific sub-topic patterns
# Applied only to the filtered AI safety paper subset.
# ---------------------------------------------------------------------------

AI_SAFETY_TOPIC_PATTERNS: list[tuple[str, re.Pattern]] = [
    ("Interpretability", re.compile(
        r"interpretabilit|mechanistic interpretabilit|"
        r"explainabilit|explainable ai|"
        r"feature attribution|saliency|sparse autoencoder|"
        r"probing|circuit (analysis|discovery)|steering (vector|behavior)|"
        r"superposition|polysemanticity|activation patching",
        re.I,
    )),
    ("Alignment & RLHF", re.compile(
        r"(?<!\w)alignment(?!\w)|preference (alignment|optimization|learning)|"
        r"(?<!\w)(?:rlhf|dpo|grpo|ppo)(?!\w)|"
        r"value alignment|reward (model|learning|hacking)|"
        r"human (feedback|preference)|constitutional|safety alignment|"
        r"llm alignment|ai alignment|post.training",
        re.I,
    )),
    ("Adversarial Attacks", re.compile(
        r"adversarial (attack|example|training|robustness|perturbation)|"
        r"jailbreak|red.team|prompt injection|"
        r"certified (defense|robustness)|evasion attack",
        re.I,
    )),
    ("Privacy", re.compile(
        r"(?<!\w)privacy(?!\w)|differential privacy|"
        r"membership inference|memorization|"
        r"data (leakage|privacy)|federated learning",
        re.I,
    )),
    ("Fairness & Bias", re.compile(
        r"(?<!\w)fairness(?!\w)|(?<!\w)bias(?!\w)|"
        r"demographic parity|equal opportunity|"
        r"group fairness|debiasing|representation bias|social bias|"
        r"implicit bias",
        re.I,
    )),
    ("Hallucination & Factuality", re.compile(
        r"hallucination|factualit|truthfulness|"
        r"knowledge (editing|update)|model editing|"
        r"factual accuracy|grounding|confabulation|sycophancy",
        re.I,
    )),
    ("Machine Unlearning", re.compile(
        r"machine unlearning|(?<!\w)unlearning(?!\w)|"
        r"right to be forgotten|selective forgetting",
        re.I,
    )),
    ("Watermarking & Detection", re.compile(
        r"watermarking|(?<!\w)watermark(?!\w)|"
        r"ai.generated (content|text|image) detection|"
        r"deepfake detection|synthetic (text|content) detection",
        re.I,
    )),
    ("Robustness & OOD", re.compile(
        r"(?<!\w)robustness(?!\w)|"
        r"out.of.distribution|(?<!\w)ood(?!\w)|"
        r"distribution shift|domain generalization|"
        r"covariate shift|spurious correlation|"
        r"distributional robustness",
        re.I,
    )),
    ("Backdoor & Poisoning", re.compile(
        r"(?:data|training|gradient) poisoning|"
        r"backdoor (attack|defense|detection)|"
        r"trojan (attack|detection)|"
        r"(?<!\w)poisoning(?!\w)|(?<!\w)backdoor(?!\w)",
        re.I,
    )),
    ("Agent Safety", re.compile(
        r"(?:safe|constrained|aligned) (?:rl|reinforcement|agent)|"
        r"(?:safety|alignment) (?:for|of) (agent|llm agent|autonomous)|"
        r"agentic (safety|risk|alignment)|"
        r"agent (safety|alignment|guardrail)|"
        r"safe (exploration|policy|planning)",
        re.I,
    )),
    ("Societal & Ethical AI", re.compile(
        r"societal|(?<!\w)ethic\w*(?!\w)|responsible ai|"
        r"trustworthy ai|ai governance|"
        r"social impact|dual.use|harmful content|toxic",
        re.I,
    )),
]


def count_ai_safety_topics(path: Path) -> tuple[int, list[tuple[str, int]]]:
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    h  = {c.value: i for i, c in enumerate(ws[1])}

    def col(row, name):
        idx = h.get(name)
        return str(row[idx] or "") if idx is not None else ""

    counts = {name: 0 for name, _ in AI_SAFETY_TOPIC_PATTERNS}
    total  = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        # Use all fields to decide if this is an AI safety paper (broad filter)
        search_all = (col(row, "Title") + " " + col(row, "Keywords") + " "
                      + col(row, "Primary Area")).lower()
        if not any(t in search_all for t in AI_SAFETY_TERMS):
            continue
        total += 1
        # Use only title + keywords for sub-topic matching to avoid the primary
        # area name "alignment, fairness, safety, privacy, and societal…"
        # triggering every sub-topic pattern simultaneously.
        search_tkw = (col(row, "Title") + " " + col(row, "Keywords")).lower()
        for name, pat in AI_SAFETY_TOPIC_PATTERNS:
            if pat.search(search_tkw):
                counts[name] += 1

    ranked = sorted(counts.items(), key=lambda x: x[1], reverse=True)
    return total, ranked


# ---------------------------------------------------------------------------
# HTML table template  (trending topics)
# ---------------------------------------------------------------------------

_HTML_TABLE_TEMPLATE = """\
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  html, body {{
    width: 990px; height: 628px;
    overflow: hidden; background: #ffffff;
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Helvetica Neue', Arial, sans-serif;
  }}
  .card {{
    width: 990px; height: 628px;
    padding: 30px 56px 36px 56px;
    background: #ffffff;
    display: flex; flex-direction: column;
  }}
  .title {{
    font-size: 30px; font-weight: 800; color: #111827;
    line-height: 1.2; letter-spacing: -0.5px; margin-bottom: 14px;
  }}
  .tbl {{ width: 100%; border-collapse: collapse; flex: 1; }}
  .tbl thead th {{
    font-size: 11.5px; font-weight: 700; color: #9ca3af;
    text-transform: uppercase; letter-spacing: 0.5px;
    padding: 0 8px 8px 8px; border-bottom: 1.5px solid #e5e7eb;
    text-align: left;
  }}
  .tbl thead th.r {{ text-align: right; }}
  .tbl tbody tr:nth-child(even) {{ background: #f9fafb; }}
  .tbl tbody td {{
    padding: 0 8px; height: 43px; vertical-align: middle;
  }}
  .rank {{
    font-size: 14px; font-weight: 800; color: #5046e5;
    width: 30px;
  }}
  .topic-name {{
    font-size: 15px; font-weight: 600; color: #111827;
    white-space: nowrap;
  }}
  .bar-cell {{ width: 260px; padding: 0 12px 0 12px !important; }}
  .bar-bg {{
    height: 9px; background: #ede9fe; border-radius: 5px; overflow: hidden;
  }}
  .bar-fill {{
    height: 100%; border-radius: 5px; background: #5046e5;
  }}
  .count-cell {{
    width: 64px; text-align: right;
    font-size: 15px; font-weight: 700; color: #111827;
  }}
  .pct-cell {{
    width: 58px; text-align: right;
    font-size: 14px; color: #6b7280;
  }}
  .rule {{ height: 1px; background: #e5e7eb; margin-top: 10px; margin-bottom: 10px; }}
  .footer {{ font-size: 12.5px; color: #6b7280; line-height: 1.55; }}
  .footer strong {{ font-weight: 700; color: #374151; }}
</style>
</head>
<body>
<div class="card">
  <div class="title">{title}</div>
  <table class="tbl">
    <thead>
      <tr>
        <th style="width:30px">#</th>
        <th>Topic</th>
        <th class="bar-cell"></th>
        <th class="r" style="width:64px">Papers</th>
        <th class="r" style="width:58px">Share</th>
      </tr>
    </thead>
    <tbody>
{rows}
    </tbody>
  </table>
  <div class="rule"></div>
  <div class="footer">{footer}</div>
</div>
</body>
</html>
"""


def build_table_html(title: str, footer: str,
                     topics: list[tuple[str, int]], total: int) -> str:
    max_count = topics[0][1] if topics else 1
    row_html  = []
    for i, (name, count) in enumerate(topics, 1):
        pct      = count / total * 100
        bar_pct  = count / max_count * 100
        row_html.append(
            f'      <tr>'
            f'<td class="rank">{i}</td>'
            f'<td class="topic-name">{name}</td>'
            f'<td class="bar-cell">'
            f'<div class="bar-bg"><div class="bar-fill" style="width:{bar_pct:.1f}%"></div></div>'
            f'</td>'
            f'<td class="count-cell">{count:,}</td>'
            f'<td class="pct-cell">{pct:.1f}%</td>'
            f'</tr>'
        )
    return _HTML_TABLE_TEMPLATE.format(
        title=title,
        footer=footer,
        rows="\n".join(row_html),
    )


def plot_trending_topics(output_path: Path):
    path = DATA_DIR / "ICLR_2026.xlsx"
    if not path.exists():
        raise FileNotFoundError(f"Missing: {path}")
    print("Counting topics in ICLR 2026 …")
    total, ranked = count_topics(path)
    top10 = ranked[:10]
    for name, count in top10:
        print(f"  {count:5d} ({count/total*100:5.1f}%)  {name}")

    html = build_table_html(
        title="ICLR 2026 — Top 10 Trending Topics",
        footer=(
            f"Each paper matched against curated keyword &amp; title patterns. "
            f"A paper can match multiple topics — percentages are of {total:,} total accepted papers. "
            f"Source: OpenReview."
        ),
        topics=top10,
        total=total,
    )
    print("Rendering table …")
    screenshot_html(html, output_path)
    print(f"Saved → {output_path}")


def plot_trending_topics_ai_safety(output_path: Path):
    path = DATA_DIR / "ICLR_2026.xlsx"
    if not path.exists():
        raise FileNotFoundError(f"Missing: {path}")
    print("Counting AI safety sub-topics in ICLR 2026 …")
    total, ranked = count_ai_safety_topics(path)
    top10 = ranked[:10]
    print(f"  Total AI safety papers: {total}")
    for name, count in top10:
        print(f"  {count:5d} ({count/total*100:5.1f}%)  {name}")

    html = build_table_html(
        title="ICLR 2026 — Top 10 AI Safety Topics",
        footer=(
            f"Among {total:,} accepted AI Safety &amp; Alignment papers (matched by safety-related keywords, "
            f"title, or primary area). A paper can match multiple topics. Source: OpenReview."
        ),
        topics=top10,
        total=total,
    )
    print("Rendering table …")
    screenshot_html(html, output_path)
    print(f"Saved → {output_path}")


_CLASSIFIED_FILE = DATA_DIR / "archetypes_classified.json"


def _load_classified() -> dict:
    if not _CLASSIFIED_FILE.exists():
        raise FileNotFoundError(
            f"Missing {_CLASSIFIED_FILE}. "
            "Run: python iclr_2026/src/classify_archetypes.py"
        )
    return json.loads(_CLASSIFIED_FILE.read_text())


def plot_archetypes(output_path: Path):
    print(f"Loading pre-classified data from {_CLASSIFIED_FILE.name} …")
    data = _load_classified()
    years_labels = [str(y) for y in YEARS]
    arch_pct = {
        a: [round(data[str(y)]["archetypes"][a] / data[str(y)]["total"] * 100, 1)
            for y in YEARS]
        for a in ARCHETYPES
    }
    for y in YEARS:
        d = data[str(y)]
        parts = "  ".join(f"{a[:3]}={d['archetypes'][a]}" for a in ARCHETYPES)
        print(f"  {y}: total={d['total']}  {parts}")

    html = build_bar_html(
        title="ICLR Paper Archetypes 2022–2026",
        footer=(
            "<strong>New Method:</strong> abstract signals a novel algorithm, model, framework, or architecture. "
            "<strong>Empirical Study:</strong> primary contribution is analysis or measurement. "
            "<strong>Theory:</strong> theorem, proof, or formal bound. "
            "<strong>Benchmark / Dataset:</strong> introduces a new evaluation resource. "
            "Classified from abstracts via multi-signal scoring. Source: OpenReview."
        ),
        years=years_labels,
        arch_pct=arch_pct,
    )

    print("Rendering chart …")
    screenshot_html(html, output_path)
    print(f"Saved → {output_path}")


def plot_archetypes_ai_safety(output_path: Path):
    print(f"Loading pre-classified data (AI Safety subset) from {_CLASSIFIED_FILE.name} …")
    data = _load_classified()
    years_labels = [str(y) for y in YEARS]
    arch_pct = {
        a: [round(data[str(y)]["ai_safety_archetypes"][a] / data[str(y)]["ai_safety_total"] * 100, 1)
            for y in YEARS]
        for a in ARCHETYPES
    }
    for y in YEARS:
        d = data[str(y)]
        parts = "  ".join(f"{a[:3]}={d['ai_safety_archetypes'][a]}" for a in ARCHETYPES)
        print(f"  {y}: safety={d['ai_safety_total']}  {parts}")

    html = build_bar_html(
        title="ICLR AI Safety Paper Archetypes 2022–2026",
        footer=(
            "Subset of accepted papers matching AI Safety &amp; Alignment keywords (safety, alignment, fairness, "
            "adversarial, privacy, interpretability, etc.). "
            "<strong>New Method:</strong> novel algorithm, model, or framework. "
            "<strong>Empirical Study:</strong> analysis or measurement. "
            "<strong>Theory:</strong> theorem, proof, or formal bound. "
            "<strong>Benchmark / Dataset:</strong> new evaluation resource. Source: OpenReview."
        ),
        years=years_labels,
        arch_pct=arch_pct,
    )

    print("Rendering chart …")
    screenshot_html(html, output_path)
    print(f"Saved → {output_path}")


# ---------------------------------------------------------------------------
# Registry + main
# ---------------------------------------------------------------------------

PLOTS = {
    "number_of_papers": plot_iclr_trends,
    "contribution_archetypes": plot_archetypes,
    "contribution_archetypes_ai_safety": plot_archetypes_ai_safety,
    "trending_topics": plot_trending_topics,
    "trending_topics_ai_safety": plot_trending_topics_ai_safety,
}


def main():
    if len(sys.argv) < 2:
        print("Usage: python src/generate_plot.py <plot_name>")
        print("Available plots:", ", ".join(PLOTS))
        sys.exit(1)

    name = sys.argv[1]
    if name not in PLOTS:
        print(f"Unknown plot '{name}'. Available: {', '.join(PLOTS)}")
        sys.exit(1)

    PLOTS[name](PLOTS_DIR / f"{name}.png")


if __name__ == "__main__":
    main()
