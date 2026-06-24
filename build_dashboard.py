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
raw = [r for r in raw if r[1]]

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

# Genre normalization: merge duplicates and consolidate tiny categories
GENRE_MAP = {
    "Political Science": "Politics",
    "Literature": "Essays",          # Borges lectures/essays/prologues
    "Short Stories": "Fiction",
    "Antitrust": "Law",
    "Information": "Non Fiction",
    "Coding": "Technology",
    "Graphic Design": "Technology",
    "Military": "History",
    "Journalism": "Non Fiction",
    "Babies": "Parenting",
    "Chess": "Gaming",
    "Physics": "Astronomy",          # hard science books grouped together
    "Writing": "Essays",
}

def norm_genre(g):
    if not g: return g
    return GENRE_MAP.get(g, g)

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
        "genre": norm_genre(r[6]),
        "country": r[7],
        "publisher": r[8],
        "reread": r[9],
        "era": era,
    }

main_books  = [to_book(r, "main")  for r in main_rows]
older_books = [to_book(r, "older") for r in older_rows]
all_books   = main_books + older_books

# ── 1. Books & pages per year — ALL years ────────────────────────────────────
year_data = defaultdict(lambda: {"books": 0, "pages": 0})
for b in all_books:
    y = b["year"]
    if y and 2001 <= y <= 2026:
        year_data[y]["books"] += 1
        year_data[y]["pages"] += b["pages"] or 0

years_sorted = sorted(year_data)
y_books = [year_data[y]["books"] for y in years_sorted]
y_pages = [year_data[y]["pages"] for y in years_sorted]
y_avg   = [round(year_data[y]["pages"] / year_data[y]["books"])
           if year_data[y]["books"] else 0 for y in years_sorted]

# ── 2. Monthly heatmap (main list only, needs exact dates) ────────────────────
MONTH_NAMES = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]

monthly_year = defaultdict(lambda: defaultdict(int))
for b in main_books:
    if b["year"] and b["month"]:
        monthly_year[b["year"]][b["month"]] += 1

heatmap_years = [str(y) for y in sorted(monthly_year.keys())]  # strings → categorical x-axis
heatmap_z     = [[monthly_year[int(y)][m] for y in heatmap_years]
                 for m in range(1, 13)]

# ── 3. Monthly totals (all years combined, main list) ────────────────────────
monthly_totals = Counter()
for b in main_books:
    if b["month"]:
        monthly_totals[b["month"]] += 1
monthly_values = [monthly_totals[m] for m in range(1, 13)]

# ── 4. Genre breakdown — ALL books ────────────────────────────────────────────
genre_counts = Counter(b["genre"] for b in all_books if b["genre"])
# Keep top 12, lump rest as Other
top_genre_names = [g for g, _ in genre_counts.most_common(12)]
other_total     = sum(c for g, c in genre_counts.items() if g not in top_genre_names)
genre_labels    = top_genre_names + (["Other"] if other_total else [])
genre_values    = [genre_counts[g] for g in top_genre_names] + ([other_total] if other_total else [])

# Genre over time (stacked, main list 2016-2026)
major_genres = [g for g, _ in genre_counts.most_common(6)]
genre_year   = defaultdict(lambda: defaultdict(int))
for b in main_books:
    if b["year"] and b["genre"] and 2016 <= b["year"] <= 2026:
        g = b["genre"] if b["genre"] in major_genres else "Other"
        genre_year[b["year"]][g] += 1

# Pages per genre (all books)
pages_by_genre = defaultdict(int)
for b in all_books:
    if b["genre"] and b["pages"]:
        pages_by_genre[b["genre"]] += b["pages"]
genre_page_labels = [g for g in top_genre_names if g in pages_by_genre]
genre_page_values = [pages_by_genre[g] for g in genre_page_labels]

# ── 5. Country distribution — ALL books ───────────────────────────────────────
country_counts = Counter(b["country"] for b in all_books if b["country"])
top_countries  = dict(country_counts.most_common(12))

# ── 6. Author / series analysis — ALL books ───────────────────────────────────
author_books_map = defaultdict(list)
for b in all_books:
    if b["author"]:
        author_books_map[b["author"]].append(b)

series_data = []
for author, books in sorted(author_books_map.items(), key=lambda x: -len(x[1])):
    if len(books) >= 3:
        yrs = sorted(set(b["year"] for b in books if b["year"]))
        genres = Counter(b["genre"] for b in books if b["genre"])
        series_data.append({
            "author": author,
            "count": len(books),
            "genre": genres.most_common(1)[0][0] if genres else "",
            "first_year": yrs[0] if yrs else None,
            "last_year": yrs[-1] if yrs else None,
            "titles": [b["title"] for b in books][:6],
        })

# ── 7. Library table — ALL books ──────────────────────────────────────────────
table_data = []
for b in sorted(all_books, key=lambda x: (x["year"] or 0, x["month"] or 0, x["idx"] or 0)):
    table_data.append({
        "title": b["title"],
        "author": b["author"] or "",
        "year": b["year"] or "",
        "pages": b["pages"] or "",
        "genre": b["genre"] or "",
        "country": b["country"] or "",
        "era": b["era"],
    })

# ── 8. Summary stats ──────────────────────────────────────────────────────────
total_main_books  = len(main_books)
total_main_pages  = sum(b["pages"] or 0 for b in main_books)
total_all_books   = len(all_books)
total_all_pages   = sum(b["pages"] or 0 for b in all_books)
total_authors     = len(set(b["author"] for b in all_books if b["author"]))
peak_year         = max(year_data, key=lambda y: year_data[y]["books"])
peak_books        = year_data[peak_year]["books"]
most_read_author  = series_data[0]["author"] if series_data else ""
most_read_count   = series_data[0]["count"]  if series_data else 0

# ── 9. Recommendations ────────────────────────────────────────────────────────
# Cross-reference titles already read so we don't recommend something they have
read_titles_lower = {b["title"].lower() for b in all_books if isinstance(b["title"], str)}

RECS = [
    {"title": "The Name of the Wind", "author": "Patrick Rothfuss",
     "why": "You finished all 32 Dresden Files and all 10 Malazan books — Rothfuss's Kingkiller Chronicle sits at exactly that intersection of intricate magic and epic scope.",
     "genre": "Fantasy"},
    {"title": "A Little Hatred", "author": "Joe Abercrombie",
     "why": "You completed the original First Law trilogy — this starts the Age of Madness sequel trilogy set 30 years later. Same world, darker tone.",
     "genre": "Fantasy"},
    {"title": "The Lies of Locke Lamora", "author": "Scott Lynch",
     "why": "Heist fantasy with sharp wit — loved by Dresden Files and Abercrombie readers alike.",
     "genre": "Fantasy"},
    {"title": "A Fire upon the Deep", "author": "Vernor Vinge",
     "why": "Grand space opera matching the Expanse's scope. You read 5/9 Expanse books — this scratches the same itch in a single standalone.",
     "genre": "Science Fiction"},
    {"title": "The Three-Body Problem", "author": "Liu Cixin",
     "why": "The biggest hard sci-fi epic of the last decade. Your Expanse and Astronomy reads make you the ideal reader for this.",
     "genre": "Science Fiction"},
    {"title": "Piranesi", "author": "Susanna Clarke",
     "why": "Short (272 pp), genre-bending, literary — bridges your Fantasy and Fiction tastes perfectly. Unreliable narrator, beautiful prose.",
     "genre": "Fantasy"},
    {"title": "The Long Earth", "author": "Terry Pratchett & Stephen Baxter",
     "why": "You've read 8 Pratchett books. This late-career collaboration adds hard sci-fi multi-verse travel to his trademark humor.",
     "genre": "Science Fiction"},
    {"title": "The Power Broker", "author": "Robert Caro",
     "why": "You read extensively across Politics, Economics, and History. This is the masterwork at all three intersections — 1,200 pages, worth every one.",
     "genre": "History"},
    {"title": "Empire of the Summer Moon", "author": "S.C. Gwynne",
     "why": "History is your second-biggest genre (133 books). This narrative history of the Comanche reads like a novel — a strong match for your pace.",
     "genre": "History"},
    {"title": "Debt: The First 5,000 Years", "author": "David Graeber",
     "why": "Connects your Economics + History + Sociology reading in one unconventional sweep. Anthropological lens on money.",
     "genre": "Economics"},
    {"title": "The Righteous Mind", "author": "Jonathan Haidt",
     "why": "You've already read it — but if you haven't finished it, this is the best Psychology book on your shelf's themes.",
     "genre": "Psychology"},
    {"title": "The Dispossessed", "author": "Ursula K. Le Guin",
     "why": "Sci-fi that doubles as political philosophy. Given your Politics + Sci-Fi reading, this is an obvious gap.",
     "genre": "Science Fiction"},
    {"title": "When Breath Becomes Air", "author": "Paul Kalanithi",
     "why": "Short (208 pp), Memoirs genre — you have 7 Memoirs books. One of the best in the category.",
     "genre": "Memoirs"},
    {"title": "Thinking, Fast and Slow", "author": "Daniel Kahneman",
     "why": "Given your Economics + Psychology combination, this is the foundational bridge between both.",
     "genre": "Economics"},
    {"title": "The Remains of the Day", "author": "Kazuo Ishiguro",
     "why": "Literary fiction, short and precise, Booker winner. Fills the gap in your Fiction shelf between genre and literary.",
     "genre": "Fiction"},
]

# Filter out books already read
RECS = [r for r in RECS if r["title"].lower() not in read_titles_lower]

# ── 10. Language by year (Spanish vs English) ─────────────────────────────────
def is_spanish(t):
    if re.search(r'[ñáéíóú¿¡]', t): return True
    w = set(t.lower().split())
    sp = {'el','la','los','las','de','del','un','una','y','en','que','por','para','con','su','al'}
    return len(w & sp) >= 2

lang_by_year = defaultdict(lambda: [0, 0])
for b in all_books:
    if b["year"] and isinstance(b["title"], str):
        lang_by_year[b["year"]][0 if is_spanish(b["title"]) else 1] += 1

lang_years  = sorted(y for y in lang_by_year if 2001 <= y <= 2026)
lang_es     = [lang_by_year[y][0] for y in lang_years]
lang_en     = [lang_by_year[y][1] for y in lang_years]
lang_pct_es = [round(100 * lang_by_year[y][0] / max(1, sum(lang_by_year[y]))) for y in lang_years]

# ── 11. Profile page data ─────────────────────────────────────────────────────
MILESTONES = [
    {"year": "c.2001–2007", "icon": "📚", "label": "Childhood & teens in Argentina",
     "detail": "Harry Potter (via Scholastic/Salamandra), Mafalda, early Fantasy. The bedrock habits form."},
    {"year": "c.2008–2010", "icon": "🎭", "label": "Late secondary school",
     "detail": "Greek tragedies, Wilde, Lorca, Descartes, Ortega y Gasset — a student being pushed through a humanities curriculum."},
    {"year": "c.2009–2015", "icon": "🎓", "label": "Economics degree, UNC Córdoba",
     "detail": "Sustained Economics wave with a clear Austrian/libertarian school — Hayek, Mises, Liberty Fund, Union Editorial. 'Economicas UNC' as publisher in 2013 pins the institution."},
    {"year": "c.2011–2012", "icon": "🏛️", "label": "Trip or stay in Spain (Barcelona)",
     "detail": "An art wave of unusual specificity: Gaudí, Park Güell, Sagrada Família symbolism, Miró, Dalí — all published by Catalan imprints (Triangle Postals). Not a casual tourist list."},
    {"year": "2015–2016", "icon": "🗳️", "label": "Argentina's 2016 political change",
     "detail": "Two books specifically about Macri's PRO party (Mundo Pro, Cambiamos) read in January 2016 — paying close attention to the political shift right before leaving."},
    {"year": "c.2016", "icon": "✈️", "label": "Migration #1 — Argentina → abroad (likely Europe)",
     "detail": "Argentine/Spanish publishers vanish; British imprints flood in. The first migration. Spanish as a reading language drops from ~25% to near zero, never fully recovering."},
    {"year": "c.2017", "icon": "🔍", "label": "Master's degree",
     "detail": "Sherlock Holmes complete run, Pratchett, King, Lovecraft — reading for pleasure in a new country, in a new language. The publisher record shows distinctly British book-buying habits."},
    {"year": "c.2018–2019", "icon": "🗽", "label": "Migration #2 — to the United States",
     "detail": "Barnes & Noble appears in 2018. By 2019: MIT Press, Princeton, Harvard, Norton, Henry Holt. Simultaneously starts reading American founding history (Hamilton, Andrew Jackson, The British Are Coming) — classic immigrant integration."},
    {"year": "c.2019–2020", "icon": "📐", "label": "PhD / career specialization",
     "detail": "Graduate-level international trade textbooks (MIT, Princeton), then a pivot to Antitrust and Regulatory Economics: railroad economics, port economics, Lectures on Antitrust Economics. The specialty locks in."},
    {"year": "2020–2022", "icon": "🏠", "label": "COVID lockdown — reading explosion",
     "detail": "79 → 86 → 117 books/year. Starts D&D campaign (12 manuals in 2021), takes up piano (Faber Adult Piano Adventures), learns American football, reads the complete Borges in Spanish — a homesick reconnection to Argentina."},
    {"year": "2023", "icon": "🔭", "label": "Buys a telescope",
     "detail": "Five astronomy books in one year: Turn Left at Orion, Binocular Highlights, Cosmos, Stargazing for Dummies, Under Alien Skies. The purchase is obvious in the bibliography."},
    {"year": "2024", "icon": "🏡", "label": "First home",
     "detail": "Buying Your First Home, I Will Teach You To Be Rich, Wealthier — a Finance wave that appears once and never before. Someone getting financially settled."},
    {"year": "2025", "icon": "👶", "label": "First child on the way",
     "detail": "Seven parenting/pregnancy books: The Expectant Father, Mayo Clinic Guide to a Healthy Pregnancy, Heading Home With Your Newborn, and more. The bibliography announces a baby."},
]

LAYERS = [
    {
        "name": "Layer 1 — The Bedrock",
        "color": "#7c6af7",
        "description": "Two constants present in virtually every year since the beginning. They don't wave — they are the floor.",
        "items": [
            {"genre": "Fantasy", "note": "Present every year 2001–2026. When nothing else is going on, this is who they are. Consumed in long, completist series runs: all 32 Dresden Files, all 16 Malazan books, the full Witcher, full Alex Verus. Not casual reading — a deep, structural appetite."},
            {"genre": "History", "note": "The second constant, often misread as a wave because its *subject* changes. It was always there. The migration to the US just redirected it from Roman/European/Argentine history toward American founding history — same hunger, different menu. 133 books total, second only to Fantasy."},
        ]
    },
    {
        "name": "Layer 2 — The Evolving Spine",
        "color": "#56cfb2",
        "description": "The career track. Never disappears but transforms with each professional stage.",
        "items": [
            {"genre": "Economics (2009–2015)", "note": "Begins with undergraduate formation at UNC Córdoba, Austrian-school flavored (Hayek, Mises, Bastiat). Grows into graduate-level international trade economics (MIT, Princeton texts, 2019). Then a decisive pivot."},
            {"genre": "Antitrust & Law (2020–present)", "note": "Railroad economics, port economics, antitrust law, Posner, Epstein, the Supreme Court term reviewed every single year. This is litigation/expert-witness economics — the specialty that defines the career. Winning at Deposition (2025) closes the loop."},
        ]
    },
    {
        "name": "Layer 3 — The Waves",
        "color": "#f7a76c",
        "description": "Four archetypes of episodic reading. Each wave is legible as a life event or emotional state.",
        "items": [
            {"genre": "Completist binges", "note": "When something clicks, it gets consumed entirely. Anne Perry Victorian mysteries (2007, 8 books). Complete Sherlock Holmes (2017, all 9 volumes). Complete Lovecraft (2018, 4 volumes including the Barnes & Noble collected editions — first US bookstore signal). Connelly + Christie + Chesterton + Rowling/Galbraith (2022–2024, 25 books). The same muscle that finishes the Dresden Files also finishes detective series."},
            {"genre": "Geographic imprinting", "note": "Reads deeply into wherever he is or recently was. Argentine politics books peak right before emigrating (2015–2016, Macri/Cambiemos). A Barcelona-specific art wave (Gaudí, Park Güell, Sagrada Família — all Catalan publishers) suggests time in Catalonia around 2011–2012. American founding history arrives precisely with the US migration (2019) and runs for three years. The library is partly a travel diary."},
            {"genre": "New hobby announcements", "note": "Hobbies arrive loud and bibliographically unmistakable. Piano (2021: Faber Adult Piano Adventures, How to Listen to and Understand Great Music). D&D (2021: 12 manuals in a single year — lockdown campaign). Stargazing (2023: five books including the two classic amateur astronomy guides, Turn Left at Orion and Binocular Highlights). BBQ (2024: Franklin Barbecue and Meathead — the two canonical texts of serious American barbecue)."},
            {"genre": "Intellectual mood clusters", "note": "Ideology-coherent reading bursts. A center-right social commentary wave in 2020 (Sowell x2, Steele x2, McWhorter, Pinker) — likely a response to that year's events. A literary sci-fi thread 2017–2019 (Gene Wolfe, Alastair Reynolds, Jack Vance — thoughtful, not pulpy). The complete Borges in Spanish across 2021–2022 — 30 books, almost all poetry and fiction, read during COVID lockdown. The most emotionally transparent wave in the whole list: homesickness rendered as bibliography."},
        ]
    },
]

# ── Build data blob ────────────────────────────────────────────────────────────
data_blob = {
    "years": years_sorted,
    "y_books": y_books,
    "y_pages": y_pages,
    "y_avg": y_avg,
    "heatmap_years": heatmap_years,
    "heatmap_z": heatmap_z,
    "month_names": MONTH_NAMES,
    "monthly_values": monthly_values,
    "genre_labels": genre_labels,
    "genre_values": genre_values,
    "major_genres": major_genres + ["Other"],
    "genre_year": {str(y): dict(genre_year[y]) for y in range(2016, 2027)},
    "genre_page_labels": genre_page_labels,
    "genre_page_values": genre_page_values,
    "country_labels": list(top_countries.keys()),
    "country_values": list(top_countries.values()),
    "series_data": series_data[:25],
    "table_data": table_data,
    "recs": RECS,
    "milestones": MILESTONES,
    "layers": LAYERS,
    "lang_years": lang_years,
    "lang_es": lang_es,
    "lang_en": lang_en,
    "lang_pct_es": lang_pct_es,
    "stats": {
        "total_main": total_main_books,
        "total_all": total_all_books,
        "total_pages": total_all_pages,
        "total_authors": total_authors,
        "total_genres": len(genre_counts),
        "peak_year": peak_year,
        "peak_books": peak_books,
        "top_author": most_read_author,
        "top_author_count": most_read_count,
    },
}

# ── HTML ───────────────────────────────────────────────────────────────────────
HTML = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>My Reading Universe</title>
<script src="https://cdn.plot.ly/plotly-2.32.0.min.js"></script>
<style>
  :root {{
    --bg:#0f1117; --surface:#1a1d27; --surface2:#22263a;
    --accent:#7c6af7; --accent2:#56cfb2; --accent3:#f7a76c;
    --text:#e8eaf0; --muted:#7b7f9a; --border:#2e3250;
  }}
  *{{box-sizing:border-box;margin:0;padding:0}}
  body{{background:var(--bg);color:var(--text);font-family:'Segoe UI',system-ui,sans-serif;line-height:1.6}}

  header{{background:linear-gradient(135deg,#1a1d27,#0f1117);border-bottom:1px solid var(--border);
    padding:1.75rem 3rem;display:flex;justify-content:space-between;align-items:center}}
  header h1{{font-size:1.75rem;font-weight:700;letter-spacing:-.5px}}
  header h1 span{{color:var(--accent)}}
  header p{{color:var(--muted);font-size:.875rem;margin-top:.2rem}}

  nav{{background:var(--surface);border-bottom:1px solid var(--border);padding:0 3rem;
    display:flex;position:sticky;top:0;z-index:100;overflow-x:auto}}
  nav button{{background:none;border:none;color:var(--muted);padding:.875rem 1.25rem;cursor:pointer;
    font-size:.85rem;border-bottom:2px solid transparent;transition:all .2s;white-space:nowrap}}
  nav button:hover{{color:var(--text)}}
  nav button.active{{color:var(--accent);border-bottom-color:var(--accent)}}

  .page{{display:none;padding:2rem 3rem;max-width:1440px;margin:0 auto}}
  .page.active{{display:block}}

  .stat-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(155px,1fr));gap:1rem;margin-bottom:2rem}}
  .stat-card{{background:var(--surface);border:1px solid var(--border);border-radius:12px;
    padding:1.25rem;text-align:center}}
  .stat-card .num{{font-size:1.9rem;font-weight:700;color:var(--accent)}}
  .stat-card .lbl{{font-size:.78rem;color:var(--muted);margin-top:.2rem}}

  .chart-card{{background:var(--surface);border:1px solid var(--border);border-radius:12px;
    padding:1.5rem;margin-bottom:1.5rem}}
  .chart-card h3{{font-size:.95rem;margin-bottom:.875rem;color:var(--text)}}
  .chart-card p.note{{font-size:.76rem;color:var(--muted);margin-top:.5rem}}

  .grid-2{{display:grid;grid-template-columns:1fr 1fr;gap:1.5rem}}
  @media(max-width:900px){{.grid-2{{grid-template-columns:1fr}}}}

  .insight-box{{background:linear-gradient(135deg,#1e1f3a,#1a1d27);border:1px solid var(--accent);
    border-radius:12px;padding:1.25rem 1.5rem;margin-bottom:1.5rem}}
  .insight-box h3{{color:var(--accent);margin-bottom:.75rem;font-size:.92rem}}
  .insight-box ul{{list-style:none;padding:0}}
  .insight-box li{{padding:.3rem 0;font-size:.86rem;color:#c0c4dc}}
  .insight-box li::before{{content:"→ ";color:var(--accent)}}

  .series-grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(280px,1fr));gap:1rem}}
  .series-card{{background:var(--surface2);border:1px solid var(--border);border-radius:10px;padding:1rem}}
  .series-card .s-author{{font-weight:600;font-size:.93rem}}
  .series-card .s-meta{{font-size:.76rem;color:var(--muted);margin:.2rem 0 .5rem}}
  .badge{{display:inline-block;background:var(--accent);color:#fff;
    font-size:.68rem;padding:.12rem .45rem;border-radius:20px;margin-right:.25rem}}
  .badge.green{{background:var(--accent2);color:#0f1117}}
  .series-card .s-titles{{font-size:.75rem;color:var(--muted);margin-top:.5rem;line-height:1.5}}

  .rec-grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(320px,1fr));gap:1rem}}
  .rec-card{{background:var(--surface2);border:1px solid var(--border);border-radius:10px;
    padding:1.25rem;border-left:3px solid var(--accent)}}
  .rec-card.sf{{border-left-color:var(--accent2)}}
  .rec-card.hist{{border-left-color:var(--accent3)}}
  .rec-card .r-genre{{font-size:.68rem;background:var(--border);color:var(--muted);
    padding:.1rem .4rem;border-radius:4px;display:inline-block;margin-bottom:.4rem}}
  .rec-card .r-title{{font-weight:600;font-size:.97rem}}
  .rec-card .r-author{{font-size:.8rem;color:var(--accent);margin-bottom:.4rem}}
  .rec-card.sf .r-author{{color:var(--accent2)}}
  .rec-card.hist .r-author{{color:var(--accent3)}}
  .rec-card .r-why{{font-size:.82rem;color:#b0b4cc}}

  .search-bar{{width:100%;background:var(--surface2);border:1px solid var(--border);color:var(--text);
    padding:.6rem 1rem;border-radius:8px;font-size:.88rem;margin-bottom:.75rem;outline:none}}
  .search-bar:focus{{border-color:var(--accent)}}
  .filter-row{{display:flex;gap:.5rem;margin-bottom:1rem;flex-wrap:wrap}}
  .filter-row select{{background:var(--surface2);border:1px solid var(--border);color:var(--text);
    padding:.35rem .7rem;border-radius:6px;font-size:.8rem;cursor:pointer;outline:none}}
  table{{width:100%;border-collapse:collapse;font-size:.82rem}}
  thead th{{background:var(--surface2);padding:.55rem .7rem;text-align:left;color:var(--muted);
    font-weight:600;border-bottom:1px solid var(--border);cursor:pointer;white-space:nowrap}}
  thead th:hover{{color:var(--text)}}
  tbody tr{{border-bottom:1px solid var(--border);transition:background .15s}}
  tbody tr:hover{{background:var(--surface2)}}
  tbody td{{padding:.45rem .7rem}}
  .g-chip{{font-size:.7rem;background:var(--border);color:var(--muted);
    padding:.12rem .38rem;border-radius:4px;white-space:nowrap}}
  .era-chip{{font-size:.66rem;background:#1e2a1a;color:var(--accent2);
    padding:.1rem .35rem;border-radius:4px}}
  .pagination{{display:flex;gap:.4rem;justify-content:center;margin-top:1rem;flex-wrap:wrap}}
  .pagination button{{background:var(--surface2);border:1px solid var(--border);color:var(--muted);
    padding:.28rem .7rem;border-radius:6px;cursor:pointer;font-size:.8rem}}
  .pagination button.active{{background:var(--accent);color:#fff;border-color:var(--accent)}}
  .pagination button:hover:not(.active){{border-color:var(--accent);color:var(--text)}}

  /* Profile / Timeline */
  .timeline{{display:flex;flex-direction:column;gap:0;position:relative;padding-left:2rem}}
  .timeline::before{{content:'';position:absolute;left:.5rem;top:0;bottom:0;width:2px;background:var(--border)}}
  .tl-item{{position:relative;padding:.75rem 0 .75rem 1.5rem;}}
  .tl-item::before{{content:'';position:absolute;left:-.85rem;top:1rem;width:10px;height:10px;
    border-radius:50%;background:var(--accent);border:2px solid var(--bg)}}
  .tl-item.milestone::before{{background:var(--accent3);width:14px;height:14px;left:-1.05rem}}
  .tl-year{{font-size:.78rem;color:var(--accent3);font-weight:600;margin-bottom:.1rem}}
  .tl-label{{font-size:.92rem;font-weight:600}}
  .tl-detail{{font-size:.8rem;color:var(--muted);margin-top:.2rem;max-width:700px}}
  .tl-icon{{margin-right:.4rem}}

  .layer-card{{background:var(--surface);border:1px solid var(--border);border-radius:12px;
    padding:1.25rem 1.5rem;margin-bottom:1.25rem}}
  .layer-card h4{{font-size:1rem;font-weight:700;margin-bottom:.35rem}}
  .layer-desc{{font-size:.82rem;color:var(--muted);margin-bottom:1rem}}
  .layer-items{{display:grid;grid-template-columns:repeat(auto-fill,minmax(300px,1fr));gap:.875rem}}
  .layer-item{{background:var(--surface2);border-radius:8px;padding:.875rem}}
  .layer-item-genre{{font-weight:600;font-size:.88rem;margin-bottom:.3rem}}
  .layer-item-note{{font-size:.79rem;color:var(--muted);line-height:1.55}}
</style>
</head>
<body>

<header>
  <div>
    <h1>My <span>Reading Universe</span></h1>
    <p>A lifetime of books — insights, patterns &amp; recommendations</p>
  </div>
  <div style="text-align:right;font-size:.78rem;color:var(--muted)">Last updated: June 2026</div>
</header>

<nav>
  <button class="active" onclick="showPage('overview',this)">Overview</button>
  <button onclick="showPage('timeline',this)">Timeline</button>
  <button onclick="showPage('genres',this)">Genres</button>
  <button onclick="showPage('authors',this)">Authors &amp; Series</button>
  <button onclick="showPage('recs',this)">Recommendations</button>
  <button onclick="showPage('library',this)">Library</button>
  <button onclick="showPage('profile',this)">Reader Profile</button>
</nav>

<!-- OVERVIEW -->
<div class="page active" id="page-overview">
  <div class="stat-grid" id="stat-grid"></div>
  <div class="insight-box">
    <h3>Key Insights</h3>
    <ul>
      <li>COVID lockdowns (2020–2022) sparked a reading explosion — you averaged <strong>94 books/year</strong> vs ~28 in 2016–2019.</li>
      <li>Fantasy leads (183 books, ~23%) with History as a powerful second (133 books, ~17%). Economics is a surprisingly strong third at 68 books.</li>
      <li>You read 8 of 14 Wheel of Time books and 5 of 9 Expanse books — still compelling runs, though not completed. By contrast, the Dresden Files (32) and Malazan (16) were read in full — strong implied enjoyment.</li>
      <li>Summer peaks clearly: June–July are consistently your highest reading months. December also spikes.</li>
      <li>Average book length dropped from <strong>639 pp/book in 2016</strong> to ~362 pp in 2022 — shorter books enabled higher volume.</li>
      <li>US (375) and UK (125) dominate, but Argentina is a meaningful third (48 books, ~6%) — reflecting bilingual reading.</li>
      <li>Jim Butcher (32) and Jorge Luis Borges (30) are your twin poles: popular fantasy and literary poetry.</li>
    </ul>
  </div>
  <div class="grid-2">
    <div class="chart-card"><h3>Books Read Per Year</h3><div id="ch-ov-books" style="height:280px"></div></div>
    <div class="chart-card"><h3>Genre Breakdown (all time)</h3><div id="ch-ov-genre" style="height:280px"></div></div>
  </div>
</div>

<!-- TIMELINE -->
<div class="page" id="page-timeline">
  <div class="chart-card">
    <h3>Books &amp; Pages per Year</h3>
    <div id="ch-tl-dual" style="height:340px"></div>
    <p class="note">Bars = books read. Line = total pages. Hover for exact numbers. Includes all years on record.</p>
  </div>
  <div class="grid-2">
    <div class="chart-card">
      <h3>Average Book Length by Year (pages)</h3>
      <div id="ch-tl-avg" style="height:280px"></div>
      <p class="note">Shorter books in high-volume years reflect a deliberate pacing shift.</p>
    </div>
    <div class="chart-card">
      <h3>Monthly Pattern — Books Finished (all years)</h3>
      <div id="ch-tl-monthly" style="height:280px"></div>
      <p class="note">Only books with an exact finish date contribute. June–July and December are clear peaks.</p>
    </div>
  </div>
  <div class="chart-card">
    <h3>Monthly Heatmap — Books Finished by Year</h3>
    <div id="ch-tl-heatmap" style="height:360px"></div>
    <p class="note">Years with only a year-number recorded (no exact date) appear sparse. 2016–onward has the best coverage.</p>
  </div>
  <div class="chart-card">
    <h3>Genre Mix Over Time (2016–2026)</h3>
    <div id="ch-tl-genre-stack" style="height:340px"></div>
  </div>
</div>

<!-- GENRES -->
<div class="page" id="page-genres">
  <div class="grid-2">
    <div class="chart-card"><h3>Genre Distribution (all time)</h3><div id="ch-g-pie" style="height:400px"></div></div>
    <div class="chart-card"><h3>Books by Country of Author (top 12)</h3><div id="ch-g-country" style="height:400px"></div></div>
  </div>
  <div class="chart-card"><h3>Pages Read by Genre</h3><div id="ch-g-pages" style="height:320px"></div>
    <p class="note">Total pages consumed per genre — accounts for both volume and book length.</p>
  </div>
</div>

<!-- AUTHORS & SERIES -->
<div class="page" id="page-authors">
  <div class="insight-box">
    <h3>Series Completion — Implied Enjoyment Signal</h3>
    <ul>
      <li><strong>Completed long series (strong signal):</strong> Dresden Files (32/20 published — you have extras), Malazan (16 incl. ICE novels), Witcher (7), Father Brown (5), Alex Verus (12), The Expanse (5/9), Wheel of Time (8/14).</li>
      <li><strong>Abandoned early (implied weaker):</strong> Authors you read once or twice and never returned to.</li>
      <li>Jim Butcher (32) and Jorge Luis Borges (30) are your two most-read authors — popular fantasy and literary poetry as twin poles.</li>
    </ul>
  </div>
  <div class="chart-card"><h3>Authors by Books Read (top 20)</h3><div id="ch-a-bar" style="height:420px"></div></div>
  <h3 style="margin:1.5rem 0 1rem;font-size:.95rem;color:var(--text)">Series &amp; Author Deep-Dive</h3>
  <div class="series-grid" id="series-grid"></div>
</div>

<!-- RECOMMENDATIONS -->
<div class="page" id="page-recs">
  <div class="insight-box">
    <h3>How these recommendations were generated</h3>
    <ul>
      <li>Derived from completed series (strong enjoyment signal) and your genre distribution.</li>
      <li>Cross-referenced against your library — nothing here is a book you've already read.</li>
      <li>Each card states the specific connection to your reading history.</li>
    </ul>
  </div>
  <div class="rec-grid" id="rec-grid"></div>
</div>

<!-- PROFILE -->
<div class="page" id="page-profile">
  <div class="insight-box" style="border-color:var(--accent3)">
    <h3 style="color:var(--accent3)">A portrait drawn from 790 books</h3>
    <ul>
      <li>Argentine, male, born ~late 1980s. Bilingual reader who emigrated twice — once to Europe (~2016), once to the United States (~2018–2019).</li>
      <li>Antitrust / regulatory economist working at the intersection of law and economics — likely litigation support or expert-witness consulting.</li>
      <li>Intellectual profile: libertarian-leaning economist with strong historical curiosity, completist fantasy reader, and a permanent emotional anchor to Argentine literature.</li>
      <li>Life in 2025–2026: new home, first child on the way. The bibliography knows before you announce it.</li>
    </ul>
  </div>

  <h3 style="font-size:.95rem;margin-bottom:1rem">Life Milestones</h3>
  <div id="timeline-container" style="position:relative;margin-bottom:2.5rem"></div>

  <div class="chart-card">
    <h3>Language Drift — Spanish vs English over time</h3>
    <div id="ch-lang" style="height:300px"></div>
    <p class="note">Spanish-language reading drops sharply around the first migration (~2016–2017) and barely recovers — except for the complete Borges read in 2021–2022, the most legible homesickness signal in the dataset.</p>
  </div>

  <h3 style="font-size:.95rem;margin:1.5rem 0 1rem">The Three-Layer Framework</h3>
  <div id="layers-container"></div>
</div>

<!-- LIBRARY -->
<div class="page" id="page-library">
  <input class="search-bar" id="search-input" placeholder="Search title, author, genre…" oninput="filterTable()">
  <div class="filter-row" id="filter-row"></div>
  <div style="overflow-x:auto">
    <table>
      <thead><tr>
        <th onclick="sortTable('title')">Title ↕</th>
        <th onclick="sortTable('author')">Author ↕</th>
        <th onclick="sortTable('year')">Year ↕</th>
        <th onclick="sortTable('pages')">Pages ↕</th>
        <th onclick="sortTable('genre')">Genre ↕</th>
        <th onclick="sortTable('country')">Country ↕</th>
      </tr></thead>
      <tbody id="table-body"></tbody>
    </table>
  </div>
  <div class="pagination" id="pagination"></div>
  <p style="font-size:.76rem;color:var(--muted);margin-top:.75rem" id="table-count"></p>
</div>

<script>
const D = {json.dumps(data_blob, ensure_ascii=False)};

// ── Plotly base theme ─────────────────────────────────────────────────────────
const BASE = {{
  paper_bgcolor:'transparent', plot_bgcolor:'transparent',
  font:{{color:'#e8eaf0',family:'Segoe UI,system-ui,sans-serif',size:12}},
  hoverlabel:{{bgcolor:'#22263a',bordercolor:'#2e3250',font:{{color:'#e8eaf0'}}}},
  xaxis:{{gridcolor:'#2e3250',zerolinecolor:'#2e3250'}},
  yaxis:{{gridcolor:'#2e3250',zerolinecolor:'#2e3250'}},
}};
const CFG = {{responsive:true,displayModeBar:false}};

// shallow merge; always call with explicit showlegend
function L(overrides) {{ return Object.assign({{}}, BASE, overrides); }}

// ── Navigation ────────────────────────────────────────────────────────────────
let built = {{}};
function showPage(id, btn) {{
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('nav button').forEach(b => b.classList.remove('active'));
  document.getElementById('page-'+id).classList.add('active');
  btn.classList.add('active');
  if (!built[id]) {{ buildCharts(id); built[id]=true; }}
}}

// ── Stat cards ────────────────────────────────────────────────────────────────
(function() {{
  const s = D.stats;
  const cards = [
    {{num: s.total_all.toLocaleString(),   lbl: 'Total Books (all time)'}},
    {{num: s.total_main.toLocaleString(),  lbl: 'Books (last 10 years)'}},
    {{num: Math.round(s.total_pages/1000)+'K', lbl: 'Pages Read (all time)'}},
    {{num: s.total_authors.toLocaleString(), lbl: 'Unique Authors'}},
    {{num: s.peak_year,  lbl: `Peak Year (${{s.peak_books}} books)`}},
    {{num: '32', lbl: 'Dresden Files Read (Jim Butcher)'}},
  ];
  document.getElementById('stat-grid').innerHTML = cards.map(c =>
    `<div class="stat-card"><div class="num">${{c.num}}</div><div class="lbl">${{c.lbl}}</div></div>`
  ).join('');
}})();

// ── Charts ────────────────────────────────────────────────────────────────────
function buildCharts(page) {{

  if (page === 'overview') {{
    Plotly.newPlot('ch-ov-books', [{{
      x:D.years, y:D.y_books, type:'bar',
      marker:{{color:D.years.map(y => y>=2020&&y<=2022 ? '#7c6af7' : '#56cfb2')}},
      hovertemplate:'<b>%{{x}}</b><br>%{{y}} books<extra></extra>',
    }}], L({{showlegend:false, margin:{{t:10,r:10,b:30,l:40}}}}), CFG);

    Plotly.newPlot('ch-ov-genre', [{{
      labels:D.genre_labels, values:D.genre_values, type:'pie', hole:0.38,
      textinfo:'label+percent',
      hovertemplate:'<b>%{{label}}</b><br>%{{value}} books<extra></extra>',
      textfont:{{size:11}},
    }}], L({{showlegend:false, margin:{{t:10,r:10,b:10,l:10}}}}), CFG);
  }}

  if (page === 'timeline') {{
    Plotly.newPlot('ch-tl-dual', [
      {{x:D.years, y:D.y_books, type:'bar', name:'Books',
        marker:{{color:'#7c6af7',opacity:.85}},
        hovertemplate:'<b>%{{x}}</b><br>%{{y}} books<extra></extra>'}},
      {{x:D.years, y:D.y_pages, type:'scatter', mode:'lines+markers',
        name:'Pages', yaxis:'y2',
        line:{{color:'#56cfb2',width:2.5}}, marker:{{size:5}},
        hovertemplate:'<b>%{{x}}</b><br>%{{y:,}} pages<extra></extra>'}},
    ], L({{
      showlegend:true,
      legend:{{bgcolor:'transparent',bordercolor:'#2e3250',x:0,y:1}},
      yaxis:{{title:'Books',gridcolor:'#2e3250'}},
      yaxis2:{{title:'Pages',overlaying:'y',side:'right',gridcolor:'#2e3250',showgrid:false}},
      margin:{{t:20,r:70,b:40,l:50}},
    }}), CFG);

    Plotly.newPlot('ch-tl-avg', [{{
      x:D.years, y:D.y_avg, type:'scatter', mode:'lines+markers',
      line:{{color:'#f7a76c',width:2.5}}, marker:{{size:6}},
      fill:'tozeroy', fillcolor:'rgba(247,167,108,0.1)',
      hovertemplate:'<b>%{{x}}</b><br>Avg %{{y}} pages/book<extra></extra>',
    }}], L({{showlegend:false, yaxis:{{title:'Pages/book',gridcolor:'#2e3250'}}, margin:{{t:10,r:20,b:30,l:55}}}}), CFG);

    Plotly.newPlot('ch-tl-monthly', [{{
      x:D.month_names, y:D.monthly_values, type:'bar',
      marker:{{color:D.monthly_values.map((_,i)=>[5,6].includes(i)||i===11?'#7c6af7':'#56cfb2')}},
      hovertemplate:'<b>%{{x}}</b><br>%{{y}} books<extra></extra>',
    }}], L({{showlegend:false, margin:{{t:10,r:15,b:40,l:45}},
      yaxis:{{title:'Books finished',gridcolor:'#2e3250'}}}}), CFG);

    // Heatmap — x=year strings (categorical), y=month names
    Plotly.newPlot('ch-tl-heatmap', [{{
      z:D.heatmap_z,
      x:D.heatmap_years,
      y:D.month_names,
      type:'heatmap',
      colorscale:[[0,'#1a1d27'],[0.3,'#3a2d7a'],[0.6,'#7c6af7'],[1,'#56cfb2']],
      showscale:true,
      hovertemplate:'<b>%{{y}} %{{x}}</b><br>%{{z}} books<extra></extra>',
    }}], L({{
      showlegend:false,
      margin:{{t:20,r:80,b:50,l:50}},
      xaxis:{{type:'category',tickangle:-45,gridcolor:'#2e3250'}},
      yaxis:{{autorange:'reversed',gridcolor:'#2e3250'}},
    }}), CFG);

    // Stacked genre over time
    const gColors={{'Fantasy':'#7c6af7','History':'#f7a76c','Fiction':'#56cfb2',
      'Mystery':'#e06c8a','Economics':'#61c0f7','Science Fiction':'#a8e063','Other':'#555c7a'}};
    const yrs2016 = D.years.filter(y => y>=2016);
    Plotly.newPlot('ch-tl-genre-stack',
      D.major_genres.map(g => ({{
        x:yrs2016,
        y:yrs2016.map(y=>(D.genre_year[y]||{{}})[g]||0),
        name:g, type:'bar',
        marker:{{color:gColors[g]||'#888'}},
        hovertemplate:`<b>${{g}} %{{x}}</b><br>%{{y}} books<extra></extra>`,
      }})),
      L({{barmode:'stack', showlegend:true,
        legend:{{bgcolor:'transparent',bordercolor:'#2e3250'}},
        margin:{{t:20,r:20,b:40,l:40}}}}), CFG);
  }}

  if (page === 'genres') {{
    Plotly.newPlot('ch-g-pie', [{{
      labels:D.genre_labels, values:D.genre_values, type:'pie', hole:0.32,
      textinfo:'label+value', textfont:{{size:11}},
      hovertemplate:'<b>%{{label}}</b><br>%{{value}} books (%{{percent}})<extra></extra>',
    }}], L({{showlegend:false, margin:{{t:10,r:10,b:10,l:10}}}}), CFG);

    Plotly.newPlot('ch-g-country', [{{
      x:D.country_values, y:D.country_labels, type:'bar', orientation:'h',
      marker:{{color:'#7c6af7'}},
      hovertemplate:'<b>%{{y}}</b><br>%{{x}} books<extra></extra>',
    }}], L({{showlegend:false, margin:{{t:10,r:20,b:30,l:90}},
      yaxis:{{autorange:'reversed',gridcolor:'#2e3250'}}}}), CFG);

    Plotly.newPlot('ch-g-pages', [{{
      x:D.genre_page_labels, y:D.genre_page_values, type:'bar',
      marker:{{color:D.genre_page_values,colorscale:'Viridis',showscale:false}},
      hovertemplate:'<b>%{{x}}</b><br>%{{y:,}} pages<extra></extra>',
    }}], L({{showlegend:false, margin:{{t:10,r:20,b:80,l:60}},
      xaxis:{{tickangle:-35}}, yaxis:{{title:'Total pages',gridcolor:'#2e3250'}}}}), CFG);
  }}

  if (page === 'authors') {{
    const top20 = D.series_data.slice(0,20);
    Plotly.newPlot('ch-a-bar', [{{
      x:top20.map(a=>a.count),
      y:top20.map(a=>a.author),
      type:'bar', orientation:'h',
      marker:{{color:top20.map(a=>a.count>=10?'#7c6af7':'#56cfb2')}},
      hovertemplate:'<b>%{{y}}</b><br>%{{x}} books<extra></extra>',
    }}], L({{showlegend:false, margin:{{t:10,r:20,b:30,l:190}},
      yaxis:{{autorange:'reversed',gridcolor:'#2e3250'}}}}), CFG);

    document.getElementById('series-grid').innerHTML = D.series_data.map(s => `
      <div class="series-card">
        <div class="s-author">${{s.author}}</div>
        <div class="s-meta">${{s.genre}} &middot; ${{s.first_year}}–${{s.last_year||'present'}}</div>
        <span class="badge ${{s.count>=7?'green':''}}">${{s.count}} books</span>
        <div class="s-titles">${{s.titles.join(' &middot; ')}}</div>
      </div>`).join('');
  }}

  if (page === 'recs') {{
    const cls = g => g==='Science Fiction'?'sf':(g==='History'||g==='Economics')?'hist':'';
    document.getElementById('rec-grid').innerHTML = D.recs.map(r => `
      <div class="rec-card ${{cls(r.genre)}}">
        <span class="r-genre">${{r.genre}}</span>
        <div class="r-title">${{r.title}}</div>
        <div class="r-author">by ${{r.author}}</div>
        <div class="r-why">${{r.why}}</div>
      </div>`).join('');
  }}

  if (page === 'profile') {{
    // Timeline
    const tc = document.getElementById('timeline-container');
    tc.innerHTML = '<div class="timeline">' + D.milestones.map(m => `
      <div class="tl-item milestone">
        <div class="tl-year">${{m.year}}</div>
        <div class="tl-label"><span class="tl-icon">${{m.icon}}</span>${{m.label}}</div>
        <div class="tl-detail">${{m.detail}}</div>
      </div>`).join('') + '</div>';

    // Language drift chart — stacked bar ES / EN + pct line
    Plotly.newPlot('ch-lang', [
      {{x:D.lang_years, y:D.lang_es, type:'bar', name:'Spanish',
        marker:{{color:'#f7a76c',opacity:.85}},
        hovertemplate:'<b>%{{x}}</b><br>%{{y}} Spanish-title books<extra></extra>'}},
      {{x:D.lang_years, y:D.lang_en, type:'bar', name:'English',
        marker:{{color:'#56cfb2',opacity:.85}},
        hovertemplate:'<b>%{{x}}</b><br>%{{y}} English-title books<extra></extra>'}},
      {{x:D.lang_years, y:D.lang_pct_es, type:'scatter', mode:'lines+markers',
        name:'% Spanish', yaxis:'y2',
        line:{{color:'#7c6af7',width:2,dash:'dot'}}, marker:{{size:5}},
        hovertemplate:'<b>%{{x}}</b><br>%{{y}}% Spanish<extra></extra>'}},
    ], L({{
      barmode:'stack', showlegend:true,
      legend:{{bgcolor:'transparent',bordercolor:'#2e3250',x:1.08,y:1}},
      yaxis:{{title:'Books',gridcolor:'#2e3250'}},
      yaxis2:{{title:'% Spanish',overlaying:'y',side:'right',range:[0,80],
        gridcolor:'#2e3250',showgrid:false,ticksuffix:'%'}},
      margin:{{t:15,r:80,b:40,l:50}},
      shapes:[
        {{type:'line',x0:2016.5,x1:2016.5,y0:0,y1:1,yref:'paper',
          line:{{color:'#7c6af7',dash:'dot',width:1}}}},
        {{type:'line',x0:2018.5,x1:2018.5,y0:0,y1:1,yref:'paper',
          line:{{color:'#56cfb2',dash:'dot',width:1}}}},
      ],
      annotations:[
        {{x:2016.5,y:.97,yref:'paper',text:'Migration 1',showarrow:false,
          font:{{size:10,color:'#7c6af7'}},xanchor:'left',bgcolor:'transparent'}},
        {{x:2018.5,y:.88,yref:'paper',text:'Migration 2',showarrow:false,
          font:{{size:10,color:'#56cfb2'}},xanchor:'left',bgcolor:'transparent'}},
      ],
    }}), CFG);

    // Layer cards
    const lc = document.getElementById('layers-container');
    lc.innerHTML = D.layers.map(layer => `
      <div class="layer-card" style="border-left:4px solid ${{layer.color}}">
        <h4 style="color:${{layer.color}}">${{layer.name}}</h4>
        <div class="layer-desc">${{layer.description}}</div>
        <div class="layer-items">
          ${{layer.items.map(item => `
            <div class="layer-item" style="border-left:2px solid ${{layer.color}}40">
              <div class="layer-item-genre" style="color:${{layer.color}}">${{item.genre}}</div>
              <div class="layer-item-note">${{item.note}}</div>
            </div>`).join('')}}
        </div>
      </div>`).join('');
  }}

  if (page === 'library') {{
    buildFilters();
    renderTable();
  }}
}}

// ── Library ───────────────────────────────────────────────────────────────────
let filtered = [...D.table_data], sortCol='year', sortDir=-1, curPage=1;
const PG = 50;

function buildFilters() {{
  const genres  = [...new Set(D.table_data.map(b=>b.genre).filter(Boolean))].sort();
  const years   = [...new Set(D.table_data.map(b=>b.year).filter(Boolean))].sort((a,b)=>a-b);
  const ctrys   = [...new Set(D.table_data.map(b=>b.country).filter(Boolean))].sort();
  document.getElementById('filter-row').innerHTML = `
    <select id="fg" onchange="filterTable()"><option value="">All genres</option>
      ${{genres.map(g=>`<option>${{g}}</option>`).join('')}}</select>
    <select id="fy" onchange="filterTable()"><option value="">All years</option>
      ${{years.map(y=>`<option>${{y}}</option>`).join('')}}</select>
    <select id="fc" onchange="filterTable()"><option value="">All countries</option>
      ${{ctrys.map(c=>`<option>${{c}}</option>`).join('')}}</select>
    <select id="fe" onchange="filterTable()"><option value="">All eras</option>
      <option value="main">Last 10 years</option>
      <option value="older">Before 2016</option></select>`;
}}

function filterTable() {{
  const q  = document.getElementById('search-input').value.toLowerCase();
  const fg = document.getElementById('fg')?.value||'';
  const fy = document.getElementById('fy')?.value||'';
  const fc = document.getElementById('fc')?.value||'';
  const fe = document.getElementById('fe')?.value||'';
  filtered = D.table_data.filter(b =>
    (!q  || b.title.toLowerCase().includes(q)||b.author.toLowerCase().includes(q)||b.genre.toLowerCase().includes(q)) &&
    (!fg || b.genre===fg) && (!fy || String(b.year)===fy) &&
    (!fc || b.country===fc) && (!fe || b.era===fe));
  curPage=1; renderTable();
}}

function sortTable(col) {{
  if(sortCol===col) sortDir*=-1; else {{sortCol=col;sortDir=1;}}
  filtered.sort((a,b)=>{{
    const av=a[col]??'', bv=b[col]??'';
    return av<bv?-sortDir:av>bv?sortDir:0;
  }});
  renderTable();
}}

function renderTable() {{
  const start=(curPage-1)*PG;
  document.getElementById('table-body').innerHTML = filtered.slice(start,start+PG).map(b=>`
    <tr>
      <td>${{b.title}}</td>
      <td>${{b.author}}</td>
      <td>${{b.year}}</td>
      <td>${{b.pages}}</td>
      <td><span class="g-chip">${{b.genre}}</span></td>
      <td>${{b.country}}</td>
    </tr>`).join('');
  document.getElementById('table-count').textContent =
    `Showing ${{start+1}}–${{Math.min(start+PG,filtered.length)}} of ${{filtered.length}} books`;
  buildPagination();
}}

function buildPagination() {{
  const total=Math.ceil(filtered.length/PG), pages=[];
  for(let i=1;i<=total;i++) {{
    if(i===1||i===total||Math.abs(i-curPage)<=2) pages.push(i);
    else if(pages[pages.length-1]!=='…') pages.push('…');
  }}
  document.getElementById('pagination').innerHTML = pages.map(p=>
    p==='…'?`<span style="color:var(--muted);padding:0 .25rem">…</span>`:
    `<button class="${{p===curPage?'active':''}}" onclick="goPage(${{p}})">${{p}}</button>`
  ).join('');
}}

function goPage(p){{ curPage=p; renderTable(); window.scrollTo(0,0); }}

// init
built['overview']=true;
buildCharts('overview');
</script>
</body>
</html>"""

with open('/home/user/Chess_ATP_Ranking/reading_dashboard.html', 'w', encoding='utf-8') as f:
    f.write(HTML)

print(f"Done. Size: {len(HTML):,} bytes")
print(f"All books: {total_all_books} ({total_main_books} main + {len(older_books)} older)")
print(f"Genres after normalization: {len(genre_counts)}")
print(f"Genre map applied: {GENRE_MAP}")
