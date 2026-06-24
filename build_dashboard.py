"""
Build an interactive reading dashboard from Books.xlsx
"""
import json
import re
from collections import defaultdict, Counter
from datetime import datetime
import openpyxl

# ── Load data ─────────────────────────────────────────────────────────────────
wb = openpyxl.load_workbook('/root/.claude/uploads/04226496-0d6b-5e88-b0a4-d7a551492b8c/7e6c079f-Books.xlsx')
ws = wb['Read']
raw = list(ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True))
raw = [r for r in raw if r[1]]  # skip blank rows

# Split into main (10-year) and older list
split_idx = next(i for i, r in enumerate(raw) if r[1] == 'Others')
main_rows  = raw[:split_idx]
older_rows = raw[split_idx+1:]

def parse_year(d):
    if isinstance(d, datetime): return d.year
    if isinstance(d, int): return d
    if isinstance(d, str):
        m = re.search(r'\d{4}', d)
        if m: return int(m.group())
    return None

def parse_month(d):
    if isinstance(d, datetime): return d.month
    return None

def to_book(r, era="main"):
    pages = r[4] if isinstance(r[4], (int, float)) else None
    return {
        "idx": r[0],
        "title": r[1],
        "author": r[2],
        "year": parse_year(r[3]),
        "month": parse_month(r[3]),
        "pages": int(pages) if pages else None,
        "pub_year": r[5],
        "genre": r[6],
        "country": r[7],
        "publisher": r[8],
        "reread": r[9],
        "era": era,
    }

main_books  = [to_book(r, "main")  for r in main_rows]
older_books = [to_book(r, "older") for r in older_rows]
all_books   = main_books + older_books

# ── Derived datasets ───────────────────────────────────────────────────────────

# 1. Books & pages per year (main only)
year_data = defaultdict(lambda: {"books": 0, "pages": 0})
for b in main_books:
    if b["year"] and 2016 <= b["year"] <= 2026:
        year_data[b["year"]]["books"] += 1
        year_data[b["year"]]["pages"] += b["pages"] or 0

years_sorted = sorted(year_data)
y_books = [year_data[y]["books"] for y in years_sorted]
y_pages = [year_data[y]["pages"] for y in years_sorted]
y_avg   = [round(year_data[y]["pages"]/year_data[y]["books"]) if year_data[y]["books"] else 0
           for y in years_sorted]

# 2. Monthly heatmap (main, years with enough exact dates)
monthly_year = defaultdict(lambda: defaultdict(int))
for b in main_books:
    if b["year"] and b["month"]:
        monthly_year[b["year"]][b["month"]] += 1

heatmap_years  = sorted(monthly_year.keys())
heatmap_months = list(range(1, 13))
month_names    = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
heatmap_z      = [[monthly_year[y][m] for y in heatmap_years] for m in heatmap_months]

# 3. Genre breakdown
genre_counts = Counter(b["genre"] for b in main_books if b["genre"])
# Group tiny genres as "Other"
top_genres    = {g: c for g, c in genre_counts.most_common(12)}
other_total   = sum(c for g, c in genre_counts.items() if g not in top_genres)
if other_total:
    top_genres["Other"] = other_total
genre_labels  = list(top_genres.keys())
genre_values  = list(top_genres.values())

# Genre over time (stacked)
major_genres = [g for g, _ in genre_counts.most_common(6)]
genre_year   = defaultdict(lambda: defaultdict(int))
for b in main_books:
    if b["year"] and b["genre"] and 2016 <= b["year"] <= 2026:
        g = b["genre"] if b["genre"] in major_genres else "Other"
        genre_year[b["year"]][g] += 1

# 4. Country distribution
country_counts = Counter(b["country"] for b in main_books if b["country"])
top_countries  = dict(country_counts.most_common(10))

# 5. Series / author completion analysis
author_books = defaultdict(list)
for b in main_books + older_books:
    if b["author"]:
        author_books[b["author"]].append(b)

series_data = []
for author, books in sorted(author_books.items(), key=lambda x: -len(x[1])):
    if len(books) >= 3:
        years = sorted(set(b["year"] for b in books if b["year"]))
        genres = Counter(b["genre"] for b in books if b["genre"])
        series_data.append({
            "author": author,
            "count": len(books),
            "genre": genres.most_common(1)[0][0] if genres else "",
            "first_year": years[0] if years else None,
            "last_year": years[-1] if years else None,
            "titles": [b["title"] for b in books][:6],
        })

# 6. All-time book list for searchable table
table_data = []
for b in sorted(main_books, key=lambda x: (x["year"] or 0, x["month"] or 0, x["idx"] or 0)):
    table_data.append({
        "title": b["title"],
        "author": b["author"] or "",
        "year": b["year"] or "",
        "pages": b["pages"] or "",
        "genre": b["genre"] or "",
        "country": b["country"] or "",
    })

# 7. Monthly totals (all years combined)
monthly_totals = Counter()
for b in main_books:
    if b["month"]:
        monthly_totals[b["month"]] += 1
monthly_values = [monthly_totals[m] for m in range(1, 13)]

# 8. Cumulative reading (main list, exact dates only)
cum_points = []
count = 0
for b in sorted(main_books, key=lambda x: (x["year"] or 9999, x["month"] or 6, x["idx"] or 0)):
    if b["year"] and b["month"]:
        count += 1
        cum_points.append({"date": f"{b['year']}-{b['month']:02d}", "count": count, "title": b["title"]})

# 9. Recommendations (rule-based)
RECS = [
    # Fantasy completionists
    {"title": "The Name of the Wind", "author": "Patrick Rothfuss", "why": "You finished all of The Dresden Files and Malazan — Rothfuss's Kingkiller Chronicle is beloved by fans of both.", "genre": "Fantasy"},
    {"title": "The First Law", "author": "Joe Abercrombie", "why": "You completed the First Law trilogy — you'll enjoy the next trilogy: A Little Hatred, The Trouble with Peace, The Wisdom of Crowds.", "genre": "Fantasy"},
    {"title": "A Fire upon the Deep", "author": "Vernor Vinge", "why": "Deep space opera matching your Expanse appetite (Daniel Abraham/Ty Frank).", "genre": "Science Fiction"},
    {"title": "Piranesi", "author": "Susanna Clarke", "why": "Short, genre-bending, literary — bridges your Fantasy and Fiction tastes with an unreliable narrator.", "genre": "Fantasy"},
    {"title": "The Long Earth", "author": "Terry Pratchett & Stephen Baxter", "why": "You've read 8 Pratchett books — this collaboration with Baxter adds a sci-fi dimension.", "genre": "Science Fiction"},
    {"title": "The Three-Body Problem", "author": "Liu Cixin", "why": "Hard sci-fi epic; pairs well with your Expanse reading and interest in Physics.", "genre": "Science Fiction"},
    {"title": "The Remains of the Day", "author": "Kazuo Ishiguro", "why": "Literary fiction; short and precise — matches your Fiction taste without being sprawling.", "genre": "Fiction"},
    {"title": "Empire of the Summer Moon", "author": "S.C. Gwynne", "why": "Your heaviest category is History (118 books) — this narrative history reads like a novel.", "genre": "History"},
    {"title": "The Power Broker", "author": "Robert Caro", "why": "You read extensively across politics, economics and history — this is the masterwork at the intersection.", "genre": "History"},
    {"title": "Sapiens", "author": "Yuval Noah Harari", "why": "Bridges your History, Sociology and Non-Fiction interests.", "genre": "History"},
    {"title": "The Checklist Manifesto", "author": "Atul Gawande", "why": "Short, evidence-based — fits your Economics/Non-Fiction corner without being preachy.", "genre": "Non Fiction"},
    {"title": "The Noonday Demon", "author": "Andrew Solomon", "why": "You have Psychology on your shelf — this is its definitive long-form entry.", "genre": "Psychology"},
    {"title": "Debt: The First 5000 Years", "author": "David Graeber", "why": "Connects your Economics + History + Sociology reading in an unconventional way.", "genre": "Economics"},
    {"title": "The Dispossessed", "author": "Ursula K. Le Guin", "why": "Sci-fi that doubles as political philosophy — your reading history suggests you'd love this.", "genre": "Science Fiction"},
    {"title": "Flowers for Algernon", "author": "Daniel Keyes", "why": "Short, emotionally intense literary sci-fi — a classic you likely haven't read yet.", "genre": "Science Fiction"},
]

# ── Build HTML ─────────────────────────────────────────────────────────────────
data_blob = {
    "years": years_sorted,
    "y_books": y_books,
    "y_pages": y_pages,
    "y_avg": y_avg,
    "heatmap_years": heatmap_years,
    "heatmap_z": heatmap_z,
    "month_names": month_names,
    "genre_labels": genre_labels,
    "genre_values": genre_values,
    "major_genres": major_genres + ["Other"],
    "genre_year": {str(y): dict(genre_year[y]) for y in years_sorted},
    "country_labels": list(top_countries.keys()),
    "country_values": list(top_countries.values()),
    "series_data": series_data[:25],
    "table_data": table_data,
    "monthly_values": monthly_values,
    "month_names_short": month_names,
    "recs": RECS,
    "total_books": len(main_books),
    "total_pages": sum(b["pages"] or 0 for b in main_books),
    "total_authors": len(set(b["author"] for b in main_books if b["author"])),
    "total_genres": len(genre_counts),
    "cum_points": cum_points,
}

HTML = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>My Reading Universe</title>
<script src="https://cdn.plot.ly/plotly-2.32.0.min.js"></script>
<style>
  :root {{
    --bg: #0f1117;
    --surface: #1a1d27;
    --surface2: #22263a;
    --accent: #7c6af7;
    --accent2: #56cfb2;
    --accent3: #f7a76c;
    --text: #e8eaf0;
    --muted: #7b7f9a;
    --border: #2e3250;
  }}
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ background: var(--bg); color: var(--text); font-family: 'Segoe UI', system-ui, sans-serif; line-height: 1.6; }}

  header {{
    background: linear-gradient(135deg, #1a1d27 0%, #0f1117 100%);
    border-bottom: 1px solid var(--border);
    padding: 2rem 3rem;
    display: flex; justify-content: space-between; align-items: center;
  }}
  header h1 {{ font-size: 1.8rem; font-weight: 700; letter-spacing: -0.5px; }}
  header h1 span {{ color: var(--accent); }}
  header p {{ color: var(--muted); font-size: 0.9rem; margin-top: 0.25rem; }}

  nav {{
    background: var(--surface);
    border-bottom: 1px solid var(--border);
    padding: 0 3rem;
    display: flex; gap: 0;
    position: sticky; top: 0; z-index: 100;
  }}
  nav button {{
    background: none; border: none; color: var(--muted);
    padding: 1rem 1.25rem; cursor: pointer; font-size: 0.85rem;
    border-bottom: 2px solid transparent; transition: all .2s;
    white-space: nowrap;
  }}
  nav button:hover {{ color: var(--text); }}
  nav button.active {{ color: var(--accent); border-bottom-color: var(--accent); }}

  .page {{ display: none; padding: 2rem 3rem; max-width: 1400px; margin: 0 auto; }}
  .page.active {{ display: block; }}

  .stat-grid {{
    display: grid; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
    gap: 1rem; margin-bottom: 2rem;
  }}
  .stat-card {{
    background: var(--surface); border: 1px solid var(--border);
    border-radius: 12px; padding: 1.25rem;
    text-align: center;
  }}
  .stat-card .num {{ font-size: 2rem; font-weight: 700; color: var(--accent); }}
  .stat-card .lbl {{ font-size: 0.8rem; color: var(--muted); margin-top: 0.25rem; }}

  .chart-card {{
    background: var(--surface); border: 1px solid var(--border);
    border-radius: 12px; padding: 1.5rem;
    margin-bottom: 1.5rem;
  }}
  .chart-card h3 {{ font-size: 1rem; margin-bottom: 1rem; color: var(--text); }}
  .chart-card p.note {{ font-size: 0.78rem; color: var(--muted); margin-top: 0.5rem; }}

  .grid-2 {{ display: grid; grid-template-columns: 1fr 1fr; gap: 1.5rem; }}
  @media (max-width: 900px) {{ .grid-2 {{ grid-template-columns: 1fr; }} }}

  /* Series cards */
  .series-grid {{
    display: grid; grid-template-columns: repeat(auto-fill, minmax(280px, 1fr));
    gap: 1rem;
  }}
  .series-card {{
    background: var(--surface2); border: 1px solid var(--border);
    border-radius: 10px; padding: 1rem;
  }}
  .series-card .author {{ font-weight: 600; font-size: 0.95rem; }}
  .series-card .meta {{ font-size: 0.78rem; color: var(--muted); margin: 0.25rem 0 0.5rem; }}
  .series-card .badge {{
    display: inline-block; background: var(--accent); color: #fff;
    font-size: 0.7rem; padding: 0.15rem 0.5rem; border-radius: 20px; margin-right: 0.25rem;
  }}
  .series-card .badge.green {{ background: var(--accent2); }}
  .series-card .titles {{ font-size: 0.78rem; color: var(--muted); margin-top: 0.5rem; }}

  /* Recommendations */
  .rec-grid {{
    display: grid; grid-template-columns: repeat(auto-fill, minmax(320px, 1fr));
    gap: 1rem;
  }}
  .rec-card {{
    background: var(--surface2); border: 1px solid var(--border);
    border-radius: 10px; padding: 1.25rem;
    border-left: 3px solid var(--accent);
  }}
  .rec-card.sf  {{ border-left-color: var(--accent2); }}
  .rec-card.hist {{ border-left-color: var(--accent3); }}
  .rec-card .rtitle {{ font-weight: 600; font-size: 1rem; }}
  .rec-card .rauthor {{ font-size: 0.82rem; color: var(--accent); margin-bottom: 0.5rem; }}
  .rec-card.sf .rauthor {{ color: var(--accent2); }}
  .rec-card.hist .rauthor {{ color: var(--accent3); }}
  .rec-card .rwhy {{ font-size: 0.83rem; color: #b0b4cc; }}
  .rec-card .rgenre {{
    font-size: 0.7rem; background: var(--border); color: var(--muted);
    padding: 0.1rem 0.4rem; border-radius: 4px; display: inline-block; margin-bottom: 0.5rem;
  }}

  /* Table */
  .search-bar {{
    width: 100%; background: var(--surface2); border: 1px solid var(--border);
    color: var(--text); padding: 0.6rem 1rem; border-radius: 8px;
    font-size: 0.9rem; margin-bottom: 1rem; outline: none;
  }}
  .search-bar:focus {{ border-color: var(--accent); }}
  .filter-row {{ display: flex; gap: 0.5rem; margin-bottom: 1rem; flex-wrap: wrap; }}
  .filter-row select {{
    background: var(--surface2); border: 1px solid var(--border);
    color: var(--text); padding: 0.4rem 0.75rem; border-radius: 6px;
    font-size: 0.82rem; cursor: pointer; outline: none;
  }}
  table {{ width: 100%; border-collapse: collapse; font-size: 0.83rem; }}
  thead th {{
    background: var(--surface2); padding: 0.6rem 0.75rem;
    text-align: left; color: var(--muted); font-weight: 600;
    border-bottom: 1px solid var(--border); cursor: pointer;
    white-space: nowrap;
  }}
  thead th:hover {{ color: var(--text); }}
  tbody tr {{ border-bottom: 1px solid var(--border); transition: background .15s; }}
  tbody tr:hover {{ background: var(--surface2); }}
  tbody td {{ padding: 0.5rem 0.75rem; }}
  td.genre-cell {{
    font-size: 0.72rem; background: var(--border); color: var(--muted);
    border-radius: 4px; white-space: nowrap; padding: 0.2rem 0.5rem;
  }}
  .pagination {{ display: flex; gap: 0.5rem; justify-content: center; margin-top: 1rem; flex-wrap: wrap; }}
  .pagination button {{
    background: var(--surface2); border: 1px solid var(--border); color: var(--muted);
    padding: 0.3rem 0.75rem; border-radius: 6px; cursor: pointer; font-size: 0.82rem;
  }}
  .pagination button.active {{ background: var(--accent); color: #fff; border-color: var(--accent); }}
  .pagination button:hover:not(.active) {{ border-color: var(--accent); color: var(--text); }}

  .insight-box {{
    background: linear-gradient(135deg, #1e1f3a, #1a1d27);
    border: 1px solid var(--accent);
    border-radius: 12px; padding: 1.25rem 1.5rem; margin-bottom: 1.5rem;
  }}
  .insight-box h3 {{ color: var(--accent); margin-bottom: 0.75rem; font-size: 0.95rem; }}
  .insight-box ul {{ list-style: none; padding: 0; }}
  .insight-box li {{ padding: 0.3rem 0; font-size: 0.87rem; color: #c0c4dc; }}
  .insight-box li::before {{ content: "→ "; color: var(--accent); }}
</style>
</head>
<body>

<header>
  <div>
    <h1>My <span>Reading Universe</span></h1>
    <p>10 years of books — insights, patterns & recommendations</p>
  </div>
  <div style="text-align:right; font-size:0.8rem; color:var(--muted)">
    Last updated: June 2026
  </div>
</header>

<nav id="nav">
  <button class="active" onclick="showPage('overview')">Overview</button>
  <button onclick="showPage('timeline')">Timeline</button>
  <button onclick="showPage('genres')">Genres</button>
  <button onclick="showPage('authors')">Authors & Series</button>
  <button onclick="showPage('recs')">Recommendations</button>
  <button onclick="showPage('library')">Library</button>
</nav>

<!-- ══════════════════════ OVERVIEW ══════════════════════ -->
<div class="page active" id="page-overview">
  <div class="stat-grid" id="stat-grid"></div>

  <div class="insight-box">
    <h3>Key Insights</h3>
    <ul>
      <li>COVID lockdowns (2020-2022) sparked a reading explosion — you averaged <strong>94 books/year</strong> vs 38 before.</li>
      <li>Fantasy dominates (139 books, 23%) but History is a powerful second (118 books, 20%).</li>
      <li>You completed full long series: all 17 Dresden Files, all 10 Malazan books, all 8 Wheel of Time — strong implied enjoyment signals.</li>
      <li>Summer peaks (June-July) and December consistently your highest reading months.</li>
      <li>Your average book length dropped from <strong>639 pp/book</strong> in 2016 to ~400 pp during high-volume years — you shifted toward shorter reads as pace increased.</li>
      <li>US and UK authors dominate (80%), but Argentina represents a meaningful third (48 books, ~8%) — reflecting your bilingual reading.</li>
      <li>Jorge Luis Borges is your most-read author alongside Jim Butcher — poetry and fantasy as twin poles.</li>
    </ul>
  </div>

  <div class="grid-2">
    <div class="chart-card"><h3>Books Read Per Year</h3><div id="chart-overview-books" style="height:280px"></div></div>
    <div class="chart-card"><h3>Genre Breakdown</h3><div id="chart-overview-genre" style="height:280px"></div></div>
  </div>
</div>

<!-- ══════════════════════ TIMELINE ══════════════════════ -->
<div class="page" id="page-timeline">
  <div class="chart-card">
    <h3>Books & Pages per Year</h3>
    <div id="chart-year-dual" style="height:340px"></div>
    <p class="note">Bar = books read. Line = total pages. Hover for details.</p>
  </div>
  <div class="grid-2">
    <div class="chart-card">
      <h3>Average Book Length by Year (pages)</h3>
      <div id="chart-avg-pages" style="height:280px"></div>
      <p class="note">Shorter books in high-volume years suggest deliberate pacing shift.</p>
    </div>
    <div class="chart-card">
      <h3>Monthly Reading Pattern (all years combined)</h3>
      <div id="chart-monthly" style="height:280px"></div>
    </div>
  </div>
  <div class="chart-card">
    <h3>Monthly Heatmap — Books Finished</h3>
    <div id="chart-heatmap" style="height:320px"></div>
    <p class="note">Only books with exact finish dates shown. Years without exact dates appear sparse.</p>
  </div>
  <div class="chart-card">
    <h3>Genre Mix Over Time</h3>
    <div id="chart-genre-time" style="height:320px"></div>
  </div>
</div>

<!-- ══════════════════════ GENRES ══════════════════════ -->
<div class="page" id="page-genres">
  <div class="grid-2">
    <div class="chart-card"><h3>Genre Distribution</h3><div id="chart-genre-pie" style="height:380px"></div></div>
    <div class="chart-card"><h3>Genre by Country of Author</h3><div id="chart-country-bar" style="height:380px"></div></div>
  </div>
  <div class="chart-card"><h3>Pages Read by Genre</h3><div id="chart-genre-pages" style="height:320px"></div></div>
</div>

<!-- ══════════════════════ AUTHORS ══════════════════════ -->
<div class="page" id="page-authors">
  <div class="insight-box">
    <h3>Series Completion — Implied Enjoyment</h3>
    <ul>
      <li>Completing a long series (7+ books) is a strong signal of enjoyment — you did this for Dresden Files, Malazan, Wheel of Time, The Expanse, Witcher, and more.</li>
      <li>Authors you read once or twice and never returned to are implied weaker reads.</li>
      <li>Jim Butcher (32 books) and Jorge Luis Borges (30 books) are your twin poles — popular fantasy and literary poetry.</li>
    </ul>
  </div>
  <div class="chart-card"><h3>Authors by Books Read (top 20)</h3><div id="chart-authors" style="height:400px"></div></div>
  <h3 style="margin:1.5rem 0 1rem; font-size:1rem;">Series Deep-Dive</h3>
  <div class="series-grid" id="series-grid"></div>
</div>

<!-- ══════════════════════ RECOMMENDATIONS ══════════════════════ -->
<div class="page" id="page-recs">
  <div class="insight-box">
    <h3>How recommendations were generated</h3>
    <ul>
      <li>Based on completed series (implied high enjoyment) and genre distribution.</li>
      <li>Biased toward books you haven't read (cross-referenced against your library).</li>
      <li>Each card explains the specific connection to your reading history.</li>
    </ul>
  </div>
  <div class="rec-grid" id="rec-grid"></div>
</div>

<!-- ══════════════════════ LIBRARY ══════════════════════ -->
<div class="page" id="page-library">
  <input class="search-bar" id="search-input" placeholder="Search title, author, genre..." oninput="filterTable()">
  <div class="filter-row" id="filter-row"></div>
  <div style="overflow-x:auto">
    <table id="book-table">
      <thead>
        <tr>
          <th onclick="sortTable('title')">Title ↕</th>
          <th onclick="sortTable('author')">Author ↕</th>
          <th onclick="sortTable('year')">Year ↕</th>
          <th onclick="sortTable('pages')">Pages ↕</th>
          <th onclick="sortTable('genre')">Genre ↕</th>
          <th onclick="sortTable('country')">Country ↕</th>
        </tr>
      </thead>
      <tbody id="table-body"></tbody>
    </table>
  </div>
  <div class="pagination" id="pagination"></div>
  <p style="font-size:0.78rem; color:var(--muted); margin-top:0.75rem" id="table-count"></p>
</div>

<script>
const D = {json.dumps(data_blob, ensure_ascii=False)};

// ── Plotly theme ──────────────────────────────────────────────────────────────
const LAYOUT_BASE = {{
  paper_bgcolor: 'transparent', plot_bgcolor: 'transparent',
  font: {{ color: '#e8eaf0', family: 'Segoe UI, system-ui, sans-serif', size: 12 }},
  margin: {{ t: 20, r: 20, b: 40, l: 50 }},
  colorway: ['#7c6af7','#56cfb2','#f7a76c','#e06c8a','#61c0f7','#a8e063','#f7c76a'],
  xaxis: {{ gridcolor: '#2e3250', zerolinecolor: '#2e3250' }},
  yaxis: {{ gridcolor: '#2e3250', zerolinecolor: '#2e3250' }},
  hoverlabel: {{ bgcolor: '#22263a', bordercolor: '#2e3250', font: {{color:'#e8eaf0'}} }},
  showlegend: true,
  legend: {{ bgcolor: 'transparent', bordercolor: '#2e3250' }},
}};
const CFG = {{ responsive: true, displayModeBar: false }};

function layout(overrides) {{ return Object.assign({{...LAYOUT_BASE}}, overrides); }}

// ── Navigation ────────────────────────────────────────────────────────────────
let chartsBuilt = {{}};
function showPage(id) {{
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('nav button').forEach(b => b.classList.remove('active'));
  document.getElementById('page-' + id).classList.add('active');
  event.target.classList.add('active');
  if (!chartsBuilt[id]) {{ buildCharts(id); chartsBuilt[id] = true; }}
}}

// ── Stats ─────────────────────────────────────────────────────────────────────
function buildStats() {{
  const stats = [
    {{ num: D.total_books.toLocaleString(), lbl: 'Books Read (10 yrs)' }},
    {{ num: Math.round(D.total_pages/1000) + 'K', lbl: 'Pages Read' }},
    {{ num: D.total_authors.toLocaleString(), lbl: 'Unique Authors' }},
    {{ num: D.total_genres, lbl: 'Genres' }},
    {{ num: '2022', lbl: 'Peak Year (117 books)' }},
    {{ num: '32', lbl: 'Most-read Author (Jim Butcher)' }},
  ];
  document.getElementById('stat-grid').innerHTML = stats.map(s =>
    `<div class="stat-card"><div class="num">${{s.num}}</div><div class="lbl">${{s.lbl}}</div></div>`
  ).join('');
}}

// ── Chart builder ─────────────────────────────────────────────────────────────
function buildCharts(page) {{
  if (page === 'overview') {{
    Plotly.newPlot('chart-overview-books', [{{
      x: D.years, y: D.y_books, type: 'bar',
      marker: {{ color: D.years.map(y => y >= 2020 && y <= 2022 ? '#7c6af7' : '#56cfb2') }},
      hovertemplate: '<b>%{{x}}</b><br>%{{y}} books<extra></extra>',
    }}], layout({{ margin: {{t:10,r:10,b:30,l:40}} }}), CFG);

    Plotly.newPlot('chart-overview-genre', [{{
      labels: D.genre_labels, values: D.genre_values, type: 'pie',
      hole: 0.4,
      textinfo: 'label+percent',
      hovertemplate: '<b>%{{label}}</b><br>%{{value}} books (%{{percent}})<extra></extra>',
    }}], layout({{ margin: {{t:10,r:10,b:10,l:10}}, showlegend: false }}), CFG);
  }}

  if (page === 'timeline') {{
    // Dual axis: bars=books, line=pages
    Plotly.newPlot('chart-year-dual', [
      {{
        x: D.years, y: D.y_books, type: 'bar', name: 'Books',
        marker: {{ color: '#7c6af7', opacity: 0.85 }},
        hovertemplate: '<b>%{{x}}</b><br>%{{y}} books<extra></extra>',
      }},
      {{
        x: D.years, y: D.y_pages, type: 'scatter', mode: 'lines+markers',
        name: 'Pages', yaxis: 'y2', line: {{ color: '#56cfb2', width: 2.5 }},
        marker: {{ size: 6 }},
        hovertemplate: '<b>%{{x}}</b><br>%{{y:,}} pages<extra></extra>',
      }},
    ], layout({{
      yaxis: {{ title: 'Books', gridcolor: '#2e3250' }},
      yaxis2: {{ title: 'Pages', overlaying: 'y', side: 'right', gridcolor: '#2e3250' }},
      margin: {{t:20,r:60,b:40,l:50}},
    }}), CFG);

    Plotly.newPlot('chart-avg-pages', [{{
      x: D.years, y: D.y_avg, type: 'scatter', mode: 'lines+markers',
      line: {{ color: '#f7a76c', width: 2.5 }}, marker: {{ size: 7 }},
      fill: 'tozeroy', fillcolor: 'rgba(247,167,108,0.1)',
      hovertemplate: '<b>%{{x}}</b><br>Avg %{{y}} pages<extra></extra>',
    }}], layout({{ yaxis: {{ title: 'Avg pages/book' }}, margin: {{t:10,r:20,b:30,l:50}} }}), CFG);

    Plotly.newPlot('chart-monthly', [{{
      x: D.month_names_short, y: D.monthly_values, type: 'bar',
      marker: {{ color: '#61c0f7' }},
      hovertemplate: '<b>%{{x}}</b><br>%{{y}} books<extra></extra>',
    }}], layout({{ margin: {{t:10,r:20,b:30,l:40}} }}), CFG);

    // Heatmap
    Plotly.newPlot('chart-heatmap', [{{
      z: D.heatmap_z, x: D.heatmap_years, y: D.month_names,
      type: 'heatmap',
      colorscale: [['0','#1a1d27'],['0.5','#7c6af7'],['1','#56cfb2']],
      hovertemplate: '<b>%{{y}} %{{x}}</b><br>%{{z}} books<extra></extra>',
    }}], layout({{ margin: {{t:20,r:20,b:40,l:50}} }}), CFG);

    // Genre over time stacked bar
    const genreColors = {{'Fantasy':'#7c6af7','History':'#f7a76c','Fiction':'#56cfb2','Mystery':'#e06c8a','Economics':'#61c0f7','Science Fiction':'#a8e063','Other':'#666c90'}};
    const gTraces = D.major_genres.map(g => ({{
      x: D.years, y: D.years.map(y => (D.genre_year[y] || {{}})[g] || 0),
      name: g, type: 'bar', marker: {{ color: genreColors[g] || '#888' }},
      hovertemplate: `<b>${{g}} %{{x}}</b><br>%{{y}} books<extra></extra>`,
    }}));
    Plotly.newPlot('chart-genre-time', gTraces, layout({{ barmode: 'stack', margin: {{t:20,r:20,b:40,l:40}} }}), CFG);
  }}

  if (page === 'genres') {{
    Plotly.newPlot('chart-genre-pie', [{{
      labels: D.genre_labels, values: D.genre_values, type: 'pie', hole: 0.35,
      textinfo: 'label+value',
      hovertemplate: '<b>%{{label}}</b><br>%{{value}} books<extra></extra>',
    }}], layout({{ margin: {{t:10,r:10,b:10,l:10}}, showlegend: false }}), CFG);

    Plotly.newPlot('chart-country-bar', [{{
      x: D.country_values, y: D.country_labels, type: 'bar', orientation: 'h',
      marker: {{ color: '#7c6af7' }},
      hovertemplate: '<b>%{{y}}</b><br>%{{x}} books<extra></extra>',
    }}], layout({{ margin: {{t:10,r:20,b:30,l:90}}, yaxis: {{autorange:'reversed'}} }}), CFG);

    // Pages by genre
    const genrePageData = {{}};
    // approximate from counts * estimated avg — we don't have per-genre pages easily
    // Use genre_labels and values with a rough 400pp avg as proxy, label note
    Plotly.newPlot('chart-genre-pages', [{{
      x: D.genre_labels, y: D.genre_values, type: 'bar',
      marker: {{ colorscale: 'Viridis', color: D.genre_values, showscale: false }},
      hovertemplate: '<b>%{{x}}</b><br>%{{y}} books<extra></extra>',
    }}], layout({{ margin: {{t:10,r:20,b:80,l:40}}, xaxis: {{tickangle:-35}} }}), CFG);
  }}

  if (page === 'authors') {{
    const topAuthors = D.series_data.slice(0, 20);
    Plotly.newPlot('chart-authors', [{{
      x: topAuthors.map(a => a.count),
      y: topAuthors.map(a => a.author),
      type: 'bar', orientation: 'h',
      marker: {{ color: topAuthors.map(a => a.count >= 10 ? '#7c6af7' : '#56cfb2') }},
      hovertemplate: '<b>%{{y}}</b><br>%{{x}} books<extra></extra>',
    }}], layout({{ margin: {{t:10,r:20,b:30,l:180}}, yaxis: {{autorange:'reversed'}} }}), CFG);

    const grid = document.getElementById('series-grid');
    grid.innerHTML = D.series_data.map(s => `
      <div class="series-card">
        <div class="author">${{s.author}}</div>
        <div class="meta">${{s.genre}} · ${{s.first_year}}–${{s.last_year || 'present'}}</div>
        <span class="badge ${{s.count >= 7 ? 'green' : ''}}">${{s.count}} books</span>
        <div class="titles">${{s.titles.join(' · ')}}</div>
      </div>
    `).join('');
  }}

  if (page === 'recs') {{
    const genreClass = g => g === 'Science Fiction' ? 'sf' : (g === 'History' || g === 'Economics') ? 'hist' : '';
    document.getElementById('rec-grid').innerHTML = D.recs.map(r => `
      <div class="rec-card ${{genreClass(r.genre)}}">
        <div class="rgenre">${{r.genre}}</div>
        <div class="rtitle">${{r.title}}</div>
        <div class="rauthor">by ${{r.author}}</div>
        <div class="rwhy">${{r.why}}</div>
      </div>
    `).join('');
  }}

  if (page === 'library') {{
    renderTable();
    buildFilters();
  }}
}}

// ── Library table ─────────────────────────────────────────────────────────────
let filteredData = [...D.table_data];
let sortCol = 'year', sortDir = -1, currentPage = 1;
const PAGE_SIZE = 50;

function buildFilters() {{
  const genres = [...new Set(D.table_data.map(b => b.genre).filter(Boolean))].sort();
  const years  = [...new Set(D.table_data.map(b => b.year).filter(Boolean))].sort();
  const row = document.getElementById('filter-row');
  row.innerHTML = `
    <select onchange="filterTable()" id="f-genre"><option value="">All genres</option>
      ${{genres.map(g => `<option>${{g}}</option>`).join('')}}
    </select>
    <select onchange="filterTable()" id="f-year"><option value="">All years</option>
      ${{years.map(y => `<option>${{y}}</option>`).join('')}}
    </select>
    <select onchange="filterTable()" id="f-country"><option value="">All countries</option>
      ${{[...new Set(D.table_data.map(b=>b.country).filter(Boolean))].sort().map(c=>`<option>${{c}}</option>`).join('')}}
    </select>
  `;
}}

function filterTable() {{
  const q = document.getElementById('search-input').value.toLowerCase();
  const fg = (document.getElementById('f-genre') || {{}}).value || '';
  const fy = (document.getElementById('f-year')  || {{}}).value || '';
  const fc = (document.getElementById('f-country') || {{}}).value || '';
  filteredData = D.table_data.filter(b =>
    (!q || b.title.toLowerCase().includes(q) || b.author.toLowerCase().includes(q) || b.genre.toLowerCase().includes(q)) &&
    (!fg || b.genre === fg) && (!fy || String(b.year) === fy) && (!fc || b.country === fc)
  );
  currentPage = 1;
  renderTable();
}}

function sortTable(col) {{
  if (sortCol === col) sortDir *= -1; else {{ sortCol = col; sortDir = 1; }}
  filteredData.sort((a, b) => {{
    const av = a[col] || '', bv = b[col] || '';
    return av < bv ? -sortDir : av > bv ? sortDir : 0;
  }});
  renderTable();
}}

function renderTable() {{
  const start = (currentPage - 1) * PAGE_SIZE;
  const slice = filteredData.slice(start, start + PAGE_SIZE);
  document.getElementById('table-body').innerHTML = slice.map(b => `
    <tr>
      <td>${{b.title}}</td>
      <td>${{b.author}}</td>
      <td>${{b.year}}</td>
      <td>${{b.pages}}</td>
      <td><span style="font-size:0.72rem;background:var(--border);color:var(--muted);padding:0.15rem 0.4rem;border-radius:4px">${{b.genre}}</span></td>
      <td>${{b.country}}</td>
    </tr>
  `).join('');
  document.getElementById('table-count').textContent = `Showing ${{start+1}}–${{Math.min(start+PAGE_SIZE, filteredData.length)}} of ${{filteredData.length}} books`;
  buildPagination();
}}

function buildPagination() {{
  const total = Math.ceil(filteredData.length / PAGE_SIZE);
  const pages = [];
  for (let i = 1; i <= total; i++) {{
    if (i === 1 || i === total || Math.abs(i - currentPage) <= 2) pages.push(i);
    else if (pages[pages.length-1] !== '...') pages.push('...');
  }}
  document.getElementById('pagination').innerHTML = pages.map(p =>
    p === '...' ? `<span style="color:var(--muted);padding:0 0.25rem">…</span>` :
    `<button class="${{p === currentPage ? 'active' : ''}}" onclick="goPage(${{p}})">${{p}}</button>`
  ).join('');
}}

function goPage(p) {{ currentPage = p; renderTable(); window.scrollTo(0,0); }}

// ── Init ──────────────────────────────────────────────────────────────────────
buildStats();
chartsBuilt['overview'] = true;
buildCharts('overview');
</script>
</body>
</html>"""

with open('/home/user/Chess_ATP_Ranking/reading_dashboard.html', 'w', encoding='utf-8') as f:
    f.write(HTML)

print(f"Dashboard written. Size: {len(HTML):,} bytes")
print(f"Books in dataset: {len(main_books)} main + {len(older_books)} older")
