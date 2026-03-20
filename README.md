# Automated Performance Reporter

Interactive Streamlit application for campaign performance reporting. Upload campaign data (Excel, SharePoint, or Redshift), explore it through filterable dashboards with KPI cards, trend charts, and hierarchy tables, then generate a branded PPTX deck in one click.

**This project expands on the [reporting-automation](../reporting-automation/) pipeline** — which generates static weekly reports via CLI — by adding a real-time interactive UI, dynamic filtering, multi-source ingestion, AI-powered ad-hoc queries, and on-demand deck generation.

---

## What's New vs. reporting-automation

| Capability | reporting-automation (CLI) | This app (Streamlit) |
|-----------|---------------------------|---------------------|
| Interface | `python main.py` | Interactive web UI |
| Data source | Single Excel or Redshift | Excel upload, SharePoint link, or Redshift |
| Filtering | Config YAML (LOB, year) | Real-time sidebar: LOB, date range, channel, market |
| Tables | Static Excel export (6 tabs) | Hierarchical LOB > Channel > Platform with MoM/WoW deltas |
| Charts | 3 matplotlib PNGs | Dynamic Plotly with selectable secondary metrics |
| Insights | Rule-based (5 bullets) | Rule-based + AI Sandbox for natural-language queries |
| Deck output | 5-slide PPTX via template | On-demand branded PPTX (KPI + trend + channel + product + insights) |
| Data context | Config YAML | Optional context document upload (describes dataset structure) |
| Persistence | None | Disk cache survives browser refresh |

---

## Project Structure

```
reporting-app/
├── app.py               # Main application (single-file Streamlit app)
├── requirements.txt     # Python dependencies
├── run.sh               # One-command launcher (installs deps if needed)
├── .gitignore
├── .app_config.json     # Persisted user settings (auto-created, gitignored)
├── .data_cache/         # Cached dataset + metadata (auto-created, gitignored)
└── README.md
```

---

## Setup

**Python 3.9 or higher required.**

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Run the app
python3 -m streamlit run app.py
```

Or use the launcher script:
```bash
chmod +x run.sh
./run.sh
```

The app opens at **http://localhost:8501**.

### Dependencies

| Package | Purpose |
|---------|---------|
| `streamlit` | Web UI framework |
| `pandas` | Data manipulation |
| `openpyxl` | Excel file support |
| `python-pptx` | PPTX deck generation |
| `plotly` | Interactive charts |
| `numpy` | Numerical operations |
| `openai` | AI Sandbox (GPT-4o for ad-hoc queries) |
| `psycopg2-binary` | Redshift connectivity (optional) |
| `requests` | SharePoint link fetching |

### Optional: AI Sandbox

The AI Sandbox tab lets users ask natural-language questions about the data and get auto-generated charts and tables. It requires an OpenAI API key.

- Enter your key in the AI Sandbox tab, or
- Set the `OPENAI_API_KEY` environment variable

If no key is provided, the rest of the app works normally — the AI Sandbox is the only feature that requires it.

---

## How to Use

### 1. Load Data

Three options in the sidebar:

- **Upload Excel** — Drag/drop an `.xlsx` file. Multi-sheet workbooks are supported; select the target sheet.
- **SharePoint Link** — Paste a public sharing URL ("Anyone with the link"). The app converts it to a download URL automatically.
- **Redshift** — Enter connection details and a SQL query.

Data is cached to disk so it survives browser refreshes.

### 2. Configure

- **Client Name** — Used in the deck title slide
- **Report Title** — Deck heading
- **Data Context Document** *(optional)* — Upload a `.txt`, `.md`, `.docx`, or `.xlsx` file that describes the dataset structure, column definitions, and business rules. This context is injected into AI Sandbox queries.

### 3. Filter

Use the filter bar to narrow by:
- Line of Business (LOB)
- Date range
- Channel

### 4. Explore Tabs

| Tab | Content |
|-----|---------|
| **Trends** | Monthly trend chart (dual-axis, selectable secondary metric), MoM narrative, channel stacked bar, monthly hierarchy table (LOB > Channel > Platform with MoM deltas), weekly hierarchy table (WoW deltas) |
| **Channel** | Channel summary table + horizontal bar chart |
| **Platform** | Platform summary table + spend vs. conversions scatter |
| **Product** | Product/segment CPA table + bar chart, Channel > Platform > Creative hierarchy |
| **Insights** | Auto-generated rule-based performance insights |
| **All Data** | Searchable full dataset with CSV export |
| **AI Sandbox** | Natural-language query interface (requires OpenAI key) |

### 5. Generate Deck

Scroll below the tabs to the **Generate Deck** section. Click **Generate Performance Deck** to build a branded PPTX, then **Download PPTX**.

---

## Column Mapping

The app automatically normalizes common column name variants:

| Recognized names | Mapped to |
|-----------------|-----------|
| `spend`, `Spend` | `Spend` |
| `impressions`, `Impressions` | `Impressions` |
| `clicks`, `Clicks` | `Clicks` |
| `conversion`, `conversions`, `Conversion` | `Conversion` |
| `site_visits`, `visits`, `Site_Visits` | `Site_Visits` |
| `lob`, `LOB` | `LOB` |
| `channel`, `Channel` | `Channel` |
| `platform`, `Platform` | `Platform` |
| `date`, `Date` | `Date` |
| `market`, `Market` | `Market` |
| `product`, `Product` | `Product` |
| `campaign`, `Campaign` | `Campaign` |
| `business_audience`, `audience` | `Audience` |
| `creative`, `ad_name`, `Ad_Name`, `ad`, `Ad` | `Creative` |

If your source file uses different names, update the `rename_map` dictionary near line 1130 in `app.py`.

---

## Brand / Style

| Element | Value |
|---------|-------|
| Primary color | Navy `#1A2B4A` |
| Accent color | Red `#DF552C` |
| Font | Avenir (fallback Calibri) |
| Background | Light gray `#F7F8FA` |
| Sidebar | Navy with white text |

---

## Relationship to reporting-automation

This app can be used **alongside or instead of** the existing CLI pipeline:

- **Same data source** — Both read from `all_Consolidated.xlsx` (or equivalent). The app's column mapping covers the same fields as `config/field_map.yaml`.
- **Complementary outputs** — The CLI generates scheduled static reports; this app is for ad-hoc exploration and on-demand decks.
- **Shared KPI logic** — Both calculate Spend, Impressions, Clicks, Conversions, CTR, CPM, CPA, CPC, and Site Visits. Formulas are equivalent.
- **No shared code dependencies** — The app is self-contained in a single file. It can be deployed independently.

---

## Troubleshooting

| Issue | Solution |
|-------|----------|
| `command not found: streamlit` | Use `python3 -m streamlit run app.py` instead |
| App slow to load | Clear cache: delete `.data_cache/` folder and refresh |
| SharePoint link fails | Ensure the link is "Anyone with the link" (not restricted). Check for auth pop-ups. |
| Charts not rendering | Install `kaleido`: `pip install kaleido` |
| AI Sandbox errors | Check your OpenAI API key is valid and has credits |
| PPTX generation fails | Ensure `python-pptx` is installed: `pip install python-pptx` |
