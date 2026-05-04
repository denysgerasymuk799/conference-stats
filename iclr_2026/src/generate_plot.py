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
