"""
Campaign Performance Reporting App
"""

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import os
import sys

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Campaign Performance Reporter",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Styling ───────────────────────────────────────────────────────────────────
st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Nunito+Sans:wght@300;400;600;700;800&display=swap');

  html, body, [class*="css"] {
    font-family: 'Nunito Sans', 'Avenir', 'Helvetica Neue', sans-serif;
  }

  /* Hide Streamlit branding */
  #MainMenu {visibility: hidden;}
  footer {visibility: hidden;}
  header {visibility: hidden;}

  /* App background */
  .stApp { background-color: #F7F8FA; }

  /* Sidebar */
  [data-testid="stSidebar"] {
    background: #1A2B4A;
  }
  [data-testid="stSidebar"] * {
    color: #FFFFFF !important;
  }
  [data-testid="stSidebar"] .stSelectbox label,
  [data-testid="stSidebar"] .stMultiSelect label,
  [data-testid="stSidebar"] .stFileUploader label {
    color: #FFFFFF !important;
    font-weight: 600;
  }

  /* Header bar */
  .app-header {
    background: #1A2B4A;
    padding: 1.5rem 2rem;
    margin: -1rem -1rem 2rem -1rem;
    border-bottom: 4px solid #E81F76;
  }
  .app-header h1 {
    color: #FFFFFF;
    font-size: 1.6rem;
    font-weight: 800;
    margin: 0;
    letter-spacing: -0.02em;
  }
  .app-header p {
    color: #8C9DB5;
    font-size: 0.85rem;
    margin: 0.25rem 0 0 0;
  }

  /* KPI cards */
  .kpi-card {
    background: #FFFFFF;
    border-radius: 8px;
    padding: 1.25rem 1.5rem;
    box-shadow: 0 1px 4px rgba(0,0,0,0.08);
    border-top: 3px solid #1A2B4A;
  }
  .kpi-card.accent { border-top-color: #E81F76; }
  .kpi-value {
    font-size: 2rem;
    font-weight: 800;
    color: #1A2B4A;
    line-height: 1;
  }
  .kpi-value.magenta { color: #E81F76; }
  .kpi-label {
    font-size: 0.75rem;
    font-weight: 600;
    color: #8C8C8C;
    text-transform: uppercase;
    letter-spacing: 0.05em;
    margin-top: 0.4rem;
  }
  .kpi-delta {
    font-size: 0.8rem;
    color: #22C55E;
    font-weight: 600;
    margin-top: 0.2rem;
  }
  .kpi-delta.down { color: #EF4444; }

  /* Section headers */
  .section-header {
    font-size: 0.7rem;
    font-weight: 700;
    color: #8C8C8C;
    text-transform: uppercase;
    letter-spacing: 0.1em;
    padding-bottom: 0.5rem;
    border-bottom: 2px solid #E81F76;
    margin-bottom: 1rem;
  }

  /* Insight cards */
  .insight-card {
    background: #1A2B4A;
    border-radius: 8px;
    padding: 1rem 1.25rem;
    color: white;
    margin-bottom: 0.75rem;
  }
  .insight-card .insight-title {
    font-size: 0.7rem;
    font-weight: 700;
    color: #E81F76;
    text-transform: uppercase;
    letter-spacing: 0.08em;
  }
  .insight-card .insight-text {
    font-size: 0.9rem;
    font-weight: 600;
    color: #FFFFFF;
    margin-top: 0.25rem;
    line-height: 1.4;
  }

  /* Table styling */
  .styled-table {
    width: 100%;
    border-collapse: collapse;
    font-size: 0.85rem;
  }
  .styled-table th {
    background: #1A2B4A;
    color: white;
    padding: 0.6rem 0.75rem;
    text-align: left;
    font-size: 0.7rem;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.05em;
  }
  .styled-table td {
    padding: 0.55rem 0.75rem;
    border-bottom: 1px solid #F0F0F0;
    color: #333333;
  }
  .styled-table tr:nth-child(even) td { background: #F9FAFB; }
  .styled-table tr:hover td { background: #EEF2FF; }
  .highlight-row td { background: #FFF0F6 !important; font-weight: 700; }

  /* Buttons */
  .stButton > button {
    background: #E81F76 !important;
    color: white !important;
    border: none !important;
    border-radius: 6px !important;
    font-weight: 700 !important;
    font-family: 'Nunito Sans', 'Avenir', sans-serif !important;
    padding: 0.5rem 1.5rem !important;
    letter-spacing: 0.03em;
  }
  .stButton > button:hover {
    background: #C4105E !important;
    transform: translateY(-1px);
    box-shadow: 0 4px 12px rgba(232, 31, 118, 0.3) !important;
  }

  /* Download button */
  .stDownloadButton > button {
    background: #1A2B4A !important;
    color: white !important;
    border: none !important;
    border-radius: 6px !important;
    font-weight: 700 !important;
    width: 100%;
  }

  /* Tabs */
  .stTabs [data-baseweb="tab-list"] {
    gap: 0;
    border-bottom: 2px solid #E0E0E0;
  }
  .stTabs [data-baseweb="tab"] {
    font-weight: 600;
    font-size: 0.85rem;
    color: #8C8C8C;
    padding: 0.6rem 1.2rem;
    border-bottom: 2px solid transparent;
  }
  .stTabs [aria-selected="true"] {
    color: #1A2B4A !important;
    border-bottom: 2px solid #E81F76 !important;
    background: transparent !important;
  }

  /* Status indicators */
  .status-good { color: #22C55E; font-weight: 700; }
  .status-warn { color: #F59E0B; font-weight: 700; }
  .status-bad  { color: #EF4444; font-weight: 700; }

  /* Connection status */
  .conn-badge {
    display: inline-block;
    padding: 0.2rem 0.6rem;
    border-radius: 20px;
    font-size: 0.7rem;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.05em;
  }
  .conn-badge.connected { background: #DCFCE7; color: #166534; }
  .conn-badge.disconnected { background: #FEE2E2; color: #991B1B; }

  /* Progress bar override */
  .stProgress > div > div { background-color: #E81F76 !important; }

</style>
""", unsafe_allow_html=True)


# ── Helpers ───────────────────────────────────────────────────────────────────

def fmt_currency(v, decimals=0):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return "—"
    if v >= 1_000_000:
        return f"${v/1_000_000:.2f}MM"
    if v >= 1_000:
        return f"${v/1_000:.0f}K"
    return f"${v:,.{decimals}f}"

def fmt_pct(v, decimals=2):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return "—"
    return f"{v*100:.{decimals}f}%"

def fmt_int(v):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return "—"
    v = int(v)
    if v >= 1_000_000:
        return f"{v/1_000_000:.2f}MM"
    if v >= 1_000:
        return f"{v/1_000:.0f}K"
    return f"{v:,}"

def safe_div(n, d, scale=1.0):
    if d and d != 0:
        return (n or 0) / d * scale
    return None

def compute_kpis(df):
    spend       = df["Spend"].sum()
    impressions = df["Impressions"].sum()
    clicks      = df["Clicks"].sum()
    visits      = df["Site_Visits"].sum()
    conversions = df["Conversion"].sum()
    return {
        "spend": spend,
        "impressions": impressions,
        "clicks": clicks,
        "visits": visits,
        "conversions": conversions,
        "ctr": safe_div(clicks, impressions),
        "cpc": safe_div(spend, clicks),
        "cpm": safe_div(spend, impressions, 1000),
        "cvr": safe_div(conversions, clicks),
        "cpa": safe_div(spend, conversions),
        "site_visit_rate": safe_div(visits, clicks),
    }

def generate_insights(df, kpis):
    insights = []
    # Channel analysis
    if "Channel" in df.columns:
        ch = df.groupby("Channel").agg(
            Spend=("Spend","sum"), Conversions=("Conversion","sum"), Clicks=("Clicks","sum")
        ).reset_index()
        ch = ch[ch["Spend"] > 0].copy()
        if not ch.empty:
            top_conv = ch.sort_values("Conversions", ascending=False).iloc[0]
            top_eff = ch[ch["Conversions"] > 0].copy()
            if not top_eff.empty:
                top_eff["CPA"] = top_eff["Spend"] / top_eff["Conversions"]
                best_cpa = top_eff.sort_values("CPA").iloc[0]
                insights.append({
                    "title": "Top Converting Channel",
                    "text": f"{top_conv['Channel']} drives the most conversions ({fmt_int(top_conv['Conversions'])}). "
                            f"Best CPA: {best_cpa['Channel']} at {fmt_currency(best_cpa['CPA'], 2)}."
                })
            zero_conv = ch[(ch["Spend"]/ch["Spend"].sum() >= 0.05) & (ch["Conversions"] == 0)]
            if not zero_conv.empty:
                names = ", ".join(zero_conv["Channel"].tolist())
                insights.append({
                    "title": "⚠ Zero Conversion Channels",
                    "text": f"{names} account for ≥5% of spend but show zero conversions. Review attribution or reallocate budget."
                })

    # Platform analysis
    if "Platform" in df.columns:
        pl = df.groupby("Platform").agg(
            Spend=("Spend","sum"), Conversions=("Conversion","sum")
        ).reset_index()
        top_platform = pl.sort_values("Spend", ascending=False).iloc[0]
        insights.append({
            "title": "Dominant Platform",
            "text": f"{top_platform['Platform']} leads with {fmt_currency(top_platform['Spend'])} spend "
                    f"({top_platform['Spend']/pl['Spend'].sum()*100:.0f}% of total)."
        })

    # CPA insight
    if kpis["cpa"]:
        insights.append({
            "title": "Overall Efficiency",
            "text": f"Average CPA: {fmt_currency(kpis['cpa'], 2)} | CTR: {fmt_pct(kpis['ctr'])} | CVR: {fmt_pct(kpis['cvr'])}"
        })

    return insights[:4]


# ── Data loading ──────────────────────────────────────────────────────────────

@st.cache_data(show_spinner=False)
def load_excel(file_bytes, filename):
    xl = pd.read_excel(BytesIO(file_bytes), sheet_name=None)
    return xl


def try_redshift(host, port, db, user, password, query):
    """Connect to Redshift and execute a query. Returns (df, error)."""
    try:
        import psycopg2
        conn = psycopg2.connect(
            host=host, port=int(port), dbname=db, user=user, password=password,
            connect_timeout=10
        )
        df = pd.read_sql(query, conn)
        conn.close()
        return df, None
    except ImportError:
        return None, "psycopg2 not installed. Run: pip install psycopg2-binary"
    except Exception as e:
        return None, str(e)


# ── Deck generation ───────────────────────────────────────────────────────────

def generate_deck(df, config):
    """Generate PPTX deck from dataframe + config. Returns BytesIO."""
    try:
        from pptx import Presentation
        from pptx.util import Inches, Pt, Emu
        from pptx.dml.color import RGBColor
        from pptx.enum.text import PP_ALIGN
        from pptx.util import Pt
        import math
    except ImportError:
        return None, "python-pptx not installed. Run: pip install python-pptx"

    NAVY    = RGBColor(0x1A, 0x2B, 0x4A)
    MAGENTA = RGBColor(0xE8, 0x1F, 0x76)
    WHITE   = RGBColor(0xFF, 0xFF, 0xFF)
    LGRAY   = RGBColor(0xF4, 0xF4, 0xF4)
    CHARCOAL= RGBColor(0x33, 0x33, 0x33)
    MGRAY   = RGBColor(0x8C, 0x8C, 0x8C)

    W = Inches(13.33)
    H = Inches(7.5)

    prs = Presentation()
    prs.slide_width  = W
    prs.slide_height = H

    blank_layout = prs.slide_layouts[6]

    def add_rect(slide, x, y, w, h, fill=None, line=None, line_width=Pt(0)):
        from pptx.util import Pt
        shape = slide.shapes.add_shape(1, x, y, w, h)
        shape.line.fill.background()
        if fill:
            shape.fill.solid()
            shape.fill.fore_color.rgb = fill
        else:
            shape.fill.background()
        if line:
            shape.line.color.rgb = line
            shape.line.width = line_width
        else:
            shape.line.fill.background()
        return shape

    def add_text(slide, text, x, y, w, h, font_size=Pt(14), bold=False,
                 color=None, align=PP_ALIGN.LEFT, italic=False, wrap=True):
        txBox = slide.shapes.add_textbox(x, y, w, h)
        tf = txBox.text_frame
        tf.word_wrap = wrap
        p = tf.paragraphs[0]
        p.alignment = align
        run = p.add_run()
        run.text = text
        run.font.size = font_size
        run.font.bold = bold
        run.font.italic = italic
        run.font.color.rgb = color or CHARCOAL
        try:
            run.font.name = "Avenir"
        except Exception:
            run.font.name = "Calibri"
        return txBox

    def header_bar(slide, title):
        bar = add_rect(slide, 0, 0, W, Inches(0.45), fill=NAVY)
        add_text(slide, title.upper(),
                 Inches(0.4), Inches(0.08), Inches(12), Inches(0.3),
                 font_size=Pt(11), bold=True, color=WHITE)

    kpis = compute_kpis(df)
    lob  = config.get("lob", "All")
    client = config.get("client", "Client")
    date_range = config.get("date_range", "")

    # ── SLIDE 1: Cover ────────────────────────────────────────────────────────
    s1 = prs.slides.add_slide(blank_layout)
    add_rect(s1, 0, 0, W, H, fill=NAVY)
    add_rect(s1, 0, H - Inches(0.08), W, Inches(0.08), fill=MAGENTA)
    add_text(s1, client.upper(), Inches(0.5), Inches(0.4), Inches(5), Inches(0.35),
             font_size=Pt(12), bold=True, color=MAGENTA)
    add_text(s1, lob, Inches(0.5), Inches(2.8), Inches(12), Inches(1.2),
             font_size=Pt(44), bold=True, color=WHITE, align=PP_ALIGN.LEFT)
    add_text(s1, "Campaign Performance Summary",
             Inches(0.5), Inches(4.0), Inches(10), Inches(0.6),
             font_size=Pt(22), bold=False, color=WHITE)
    add_text(s1, date_range or "Full Campaign Flight",
             Inches(0.5), Inches(4.65), Inches(8), Inches(0.4),
             font_size=Pt(16), bold=False, color=MGRAY)
    add_text(s1, date_range or "",
             W - Inches(3.2), H - Inches(0.4), Inches(2.8), Inches(0.3),
             font_size=Pt(9), bold=False, color=MGRAY, align=PP_ALIGN.RIGHT)

    # ── SLIDE 2: KPI Summary ──────────────────────────────────────────────────
    s2 = prs.slides.add_slide(blank_layout)
    add_rect(s2, 0, 0, W, H, fill=WHITE)
    header_bar(s2, "Performance Summary")

    kpi_items = [
        (fmt_currency(kpis["spend"]),  "Total Spend",      NAVY),
        (fmt_int(kpis["impressions"]), "Impressions",      NAVY),
        (fmt_int(kpis["conversions"]), "Conversions",      MAGENTA),
        (fmt_currency(kpis["cpa"], 2) if kpis["cpa"] else "—", "Avg CPA", MAGENTA),
    ]
    card_w = Inches(2.8)
    card_h = Inches(1.6)
    gap    = Inches(0.25)
    start_x = Inches(0.6)
    start_y = Inches(0.7)

    for i, (val, label, color) in enumerate(kpi_items):
        cx = start_x + i * (card_w + gap)
        add_rect(s2, cx, start_y, card_w, card_h, fill=LGRAY)
        add_rect(s2, cx, start_y, Inches(0.06), card_h, fill=color)
        add_text(s2, val, cx + Inches(0.18), start_y + Inches(0.25),
                 card_w - Inches(0.25), Inches(0.8),
                 font_size=Pt(32), bold=True, color=color)
        add_text(s2, label, cx + Inches(0.18), start_y + Inches(1.1),
                 card_w - Inches(0.25), Inches(0.35),
                 font_size=Pt(10), bold=False, color=MGRAY)

    # Efficiency row
    eff_items = [
        (fmt_pct(kpis["ctr"]),  "CTR"),
        (fmt_pct(kpis["cvr"]),  "CVR"),
        (fmt_currency(kpis["cpc"], 2) if kpis["cpc"] else "—", "CPC"),
        (fmt_currency(kpis["cpm"], 2) if kpis["cpm"] else "—", "CPM"),
    ]
    ey = start_y + card_h + Inches(0.3)
    mini_w = Inches(2.8)
    for i, (val, label) in enumerate(eff_items):
        ex = start_x + i * (mini_w + gap)
        add_rect(s2, ex, ey, mini_w, Inches(0.9), fill=WHITE)
        add_rect(s2, ex, ey + Inches(0.78), mini_w, Inches(0.04), fill=NAVY)
        add_text(s2, val, ex + Inches(0.1), ey + Inches(0.05),
                 mini_w - Inches(0.2), Inches(0.5),
                 font_size=Pt(22), bold=True, color=NAVY)
        add_text(s2, label, ex + Inches(0.1), ey + Inches(0.52),
                 mini_w - Inches(0.2), Inches(0.25),
                 font_size=Pt(9), bold=False, color=MGRAY)

    # ── SLIDE 3: Channel breakdown ────────────────────────────────────────────
    if "Channel" in df.columns:
        s3 = prs.slides.add_slide(blank_layout)
        add_rect(s3, 0, 0, W, H, fill=WHITE)
        header_bar(s3, "Channel Performance")

        ch_df = df.groupby("Channel").agg(
            Spend=("Spend","sum"), Conversions=("Conversion","sum"), Clicks=("Clicks","sum")
        ).reset_index().sort_values("Spend", ascending=False).head(8)

        max_spend = ch_df["Spend"].max() or 1
        bar_area_w = Inches(7.5)
        bar_area_x = Inches(2.2)
        row_h = Inches(0.55)
        row_gap = Inches(0.1)
        start_y = Inches(0.6)

        for i, row in ch_df.iterrows():
            ry = start_y + (ch_df.index.get_loc(i)) * (row_h + row_gap)
            # Label
            add_text(s3, row["Channel"], Inches(0.4), ry + Inches(0.12),
                     Inches(1.75), Inches(0.35), font_size=Pt(10), bold=False, color=CHARCOAL)
            # Spend bar
            bar_len = bar_area_w * (row["Spend"] / max_spend)
            add_rect(s3, bar_area_x, ry, bar_len, Inches(0.28), fill=NAVY)
            add_text(s3, fmt_currency(row["Spend"]),
                     bar_area_x + bar_len + Inches(0.1), ry,
                     Inches(1.2), Inches(0.28), font_size=Pt(9), bold=True, color=NAVY)
            # Conv label
            conv_text = f"{fmt_int(row['Conversions'])} conv" if row['Conversions'] > 0 else "0 conv"
            add_text(s3, conv_text,
                     bar_area_x + bar_len + Inches(0.1), ry + Inches(0.3),
                     Inches(1.2), Inches(0.22), font_size=Pt(8), bold=False, color=MGRAY)

    # ── SLIDE 4: Platform breakdown ───────────────────────────────────────────
    if "Platform" in df.columns:
        s4 = prs.slides.add_slide(blank_layout)
        add_rect(s4, 0, 0, W, H, fill=WHITE)
        header_bar(s4, "Platform Performance")

        pl_df = df.groupby("Platform").agg(
            Spend=("Spend","sum"), Conversions=("Conversion","sum"),
            Clicks=("Clicks","sum"), Impressions=("Impressions","sum")
        ).reset_index().sort_values("Spend", ascending=False).head(6)
        pl_df["CPA"] = pl_df.apply(
            lambda r: r["Spend"]/r["Conversions"] if r["Conversions"] > 0 else None, axis=1
        )

        max_spend = pl_df["Spend"].max() or 1
        bar_area_w = Inches(6.0)
        bar_area_x = Inches(2.4)
        row_h = Inches(0.6)
        row_gap = Inches(0.12)
        start_y = Inches(0.6)

        for idx, (_, row) in enumerate(pl_df.iterrows()):
            ry = start_y + idx * (row_h + row_gap)
            add_text(s4, row["Platform"], Inches(0.4), ry + Inches(0.15),
                     Inches(1.9), Inches(0.35), font_size=Pt(10), bold=True, color=CHARCOAL)
            bar_len = bar_area_w * (row["Spend"] / max_spend)
            color = MAGENTA if idx == 0 else NAVY
            add_rect(s4, bar_area_x, ry, bar_len, Inches(0.3), fill=color)
            add_text(s4, fmt_currency(row["Spend"]),
                     bar_area_x + bar_len + Inches(0.1), ry,
                     Inches(1.4), Inches(0.3), font_size=Pt(9), bold=True, color=color)
            cpa_str = fmt_currency(row["CPA"], 2) + " CPA" if row["CPA"] else "No direct conv."
            add_text(s4, cpa_str,
                     bar_area_x + bar_len + Inches(0.1), ry + Inches(0.32),
                     Inches(1.8), Inches(0.24), font_size=Pt(8), bold=False, color=MGRAY)

    # ── SLIDE 5: Insights ─────────────────────────────────────────────────────
    s5 = prs.slides.add_slide(blank_layout)
    add_rect(s5, 0, 0, W, H, fill=LGRAY)
    add_rect(s5, 0, 0, W, Inches(0.45), fill=MAGENTA)
    add_text(s5, "KEY INSIGHTS",
             Inches(0.4), Inches(0.08), Inches(12), Inches(0.3),
             font_size=Pt(11), bold=True, color=WHITE)

    insights = generate_insights(df, kpis)
    card_h_ins = Inches(1.4)
    card_gap = Inches(0.2)
    cols = 2
    ins_w = (W - Inches(1.0) - Inches(0.2)) / cols
    start_y_ins = Inches(0.65)

    for idx, ins in enumerate(insights):
        col = idx % cols
        row = idx // cols
        cx = Inches(0.5) + col * (ins_w + Inches(0.2))
        cy = start_y_ins + row * (card_h_ins + card_gap)
        add_rect(s5, cx, cy, ins_w, card_h_ins, fill=NAVY)
        add_text(s5, ins["title"].upper(),
                 cx + Inches(0.2), cy + Inches(0.15),
                 ins_w - Inches(0.4), Inches(0.3),
                 font_size=Pt(8), bold=True, color=MAGENTA)
        add_text(s5, ins["text"],
                 cx + Inches(0.2), cy + Inches(0.45),
                 ins_w - Inches(0.4), Inches(0.85),
                 font_size=Pt(11), bold=False, color=WHITE)

    # Closing slide
    s6 = prs.slides.add_slide(blank_layout)
    add_rect(s6, 0, 0, W, H, fill=NAVY)
    add_rect(s6, 0, H - Inches(0.08), W, Inches(0.08), fill=MAGENTA)
    add_text(s6, "Campaign Analytics",
             Inches(0.5), Inches(2.5), Inches(12), Inches(0.8),
             font_size=Pt(36), bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_text(s6, client.upper() if client else "",
             Inches(0.5), Inches(3.4), Inches(12), Inches(0.5),
             font_size=Pt(18), bold=False, color=MGRAY, align=PP_ALIGN.CENTER)
    ds = f"Data: {lob} | {date_range}" if date_range else f"Data: {lob}"
    add_text(s6, ds,
             Inches(0.5), H - Inches(0.6), Inches(12), Inches(0.3),
             font_size=Pt(9), bold=False, color=MGRAY, align=PP_ALIGN.CENTER)

    buf = BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf, None


# ── Sidebar ───────────────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown("""
    <div style="padding:0.5rem 0 1.5rem 0; border-bottom:1px solid rgba(255,255,255,0.15); margin-bottom:1.5rem;">
        <div style="font-size:1.2rem;font-weight:800;color:#FFFFFF;letter-spacing:-0.02em;">
            📊 Campaign Reporter
        </div>
        <div style="font-size:0.75rem;color:#8C9DB5;margin-top:0.25rem;">
            Campaign Analytics
        </div>
    </div>
    """, unsafe_allow_html=True)

    data_source = st.radio(
        "DATA SOURCE",
        ["Upload Excel", "Redshift"],
        index=0,
    )

    st.markdown("---")
    st.markdown("**REPORT SETTINGS**")

    client_name = st.text_input("Client Name", value="Chubb", key="client")
    report_title = st.text_input("Report Title", value="Campaign Performance Summary")

    df_raw = None
    redshift_connected = False

    if data_source == "Upload Excel":
        uploaded = st.file_uploader(
            "Drop Excel file here",
            type=["xlsx", "xls"],
            help="Supports multi-sheet workbooks",
        )
        if uploaded:
            with st.spinner("Reading file..."):
                sheets = load_excel(uploaded.read(), uploaded.name)
            sheet_names = list(sheets.keys())
            selected_sheet = st.selectbox("Sheet", sheet_names)
            df_raw = sheets[selected_sheet]
            st.success(f"✓ {len(df_raw):,} rows loaded")

    else:
        st.markdown("**REDSHIFT CONNECTION**")
        rs_host = st.text_input("Host", placeholder="your-cluster.region.redshift.amazonaws.com")
        rs_port = st.text_input("Port", value="5439")
        rs_db   = st.text_input("Database", placeholder="analytics")
        rs_user = st.text_input("Username")
        rs_pass = st.text_input("Password", type="password")
        rs_query = st.text_area(
            "SQL Query",
            value="SELECT * FROM campaign_data WHERE lob = 'Personal Risk Services' LIMIT 10000",
            height=100,
        )
        if st.button("Connect & Run Query"):
            if all([rs_host, rs_db, rs_user, rs_pass]):
                with st.spinner("Connecting to Redshift..."):
                    df_raw, err = try_redshift(rs_host, rs_port, rs_db, rs_user, rs_pass, rs_query)
                if err:
                    st.error(f"Connection failed: {err}")
                else:
                    st.session_state["redshift_df"] = df_raw
                    st.success(f"✓ {len(df_raw):,} rows returned")
            else:
                st.warning("Fill in all connection fields.")

        if "redshift_df" in st.session_state:
            df_raw = st.session_state["redshift_df"]
            redshift_connected = True
            st.markdown('<span class="conn-badge connected">● CONNECTED</span>', unsafe_allow_html=True)


# ── Main ──────────────────────────────────────────────────────────────────────

st.markdown("""
<div class="app-header">
    <h1>Campaign Performance Reporter</h1>
    <p>Upload an Excel workbook or connect to Redshift to generate a performance summary deck</p>
</div>
""", unsafe_allow_html=True)

if df_raw is None:
    # Landing state
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("""
        <div class="kpi-card">
            <div class="kpi-value">01</div>
            <div class="kpi-label">Load Your Data</div>
            <div style="font-size:0.85rem;color:#555;margin-top:0.5rem;">
                Upload an Excel workbook or connect to Redshift using the sidebar.
            </div>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown("""
        <div class="kpi-card accent">
            <div class="kpi-value magenta">02</div>
            <div class="kpi-label">Filter & Configure</div>
            <div style="font-size:0.85rem;color:#555;margin-top:0.5rem;">
                Filter by LOB, date range, channel, or market. Set your report title and client name.
            </div>
        </div>
        """, unsafe_allow_html=True)
    with col3:
        st.markdown("""
        <div class="kpi-card">
            <div class="kpi-value">03</div>
            <div class="kpi-label">Generate & Download</div>
            <div style="font-size:0.85rem;color:#555;margin-top:0.5rem;">
                One click generates a branded PPTX performance summary deck ready for presentation.
            </div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.info("← Get started by selecting a data source in the sidebar.", icon="👈")

else:
    # ── Column mapping ─────────────────────────────────────────────────────────
    # Normalize common column name variants
    rename_map = {
        "spend": "Spend", "Spend": "Spend",
        "impressions": "Impressions", "Impressions": "Impressions",
        "clicks": "Clicks", "Clicks": "Clicks",
        "site_visits": "Site_Visits", "Site_Visits": "Site_Visits", "visits": "Site_Visits",
        "conversion": "Conversion", "conversions": "Conversion", "Conversion": "Conversion",
        "lob": "LOB", "LOB": "LOB",
        "channel": "Channel", "Channel": "Channel",
        "platform": "Platform", "Platform": "Platform",
        "date": "Date", "Date": "Date",
        "market": "Market", "Market": "Market",
        "product": "Product", "Product": "Product",
        "campaign": "Campaign", "Campaign": "Campaign",
    }
    df = df_raw.rename(columns={k: v for k, v in rename_map.items() if k in df_raw.columns})

    # Ensure numeric cols
    for col in ["Spend", "Impressions", "Clicks", "Site_Visits", "Conversion"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # ── Filters ────────────────────────────────────────────────────────────────
    st.markdown('<div class="section-header">Filters</div>', unsafe_allow_html=True)
    filter_cols = st.columns(4)

    lob_options = ["All"]
    if "LOB" in df.columns:
        lob_options += sorted(df["LOB"].dropna().unique().tolist())

    with filter_cols[0]:
        selected_lob = st.selectbox("Line of Business", lob_options)

    date_range_label = ""
    if "Date" in df.columns:
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
        min_d, max_d = df["Date"].min(), df["Date"].max()
        with filter_cols[1]:
            start_d = st.date_input("Start Date", value=min_d, min_value=min_d, max_value=max_d)
        with filter_cols[2]:
            end_d   = st.date_input("End Date",   value=max_d, min_value=min_d, max_value=max_d)
        date_range_label = f"{start_d.strftime('%b %Y')} – {end_d.strftime('%b %Y')}"

    channel_opts = ["All"]
    if "Channel" in df.columns:
        channel_opts += sorted(df["Channel"].dropna().unique().tolist())
    with filter_cols[3]:
        selected_channel = st.selectbox("Channel", channel_opts)

    # Apply filters
    dff = df.copy()
    if selected_lob != "All" and "LOB" in dff.columns:
        dff = dff[dff["LOB"] == selected_lob]
    if "Date" in dff.columns and "start_d" in dir():
        dff = dff[(dff["Date"] >= pd.Timestamp(start_d)) & (dff["Date"] <= pd.Timestamp(end_d))]
    if selected_channel != "All" and "Channel" in dff.columns:
        dff = dff[dff["Channel"] == selected_channel]

    st.markdown(f"**{len(dff):,} rows** after filters", unsafe_allow_html=False)
    st.markdown("<hr style='border:none;border-top:1px solid #E0E0E0;margin:0.5rem 0 1.5rem 0'>", unsafe_allow_html=True)

    # ── KPI Banner ─────────────────────────────────────────────────────────────
    kpis = compute_kpis(dff)

    k1, k2, k3, k4, k5, k6 = st.columns(6)
    kpi_cards = [
        (k1, fmt_currency(kpis["spend"]),       "Total Spend",      "navy"),
        (k2, fmt_int(kpis["impressions"]),       "Impressions",      "navy"),
        (k3, fmt_int(kpis["clicks"]),            "Clicks",           "navy"),
        (k4, fmt_int(kpis["conversions"]),       "Conversions",      "magenta"),
        (k5, fmt_pct(kpis["ctr"]),               "CTR",              "navy"),
        (k6, fmt_currency(kpis["cpa"],2) if kpis["cpa"] else "—", "CPA", "magenta"),
    ]
    for col, val, label, style in kpi_cards:
        accent = "accent" if style == "magenta" else ""
        val_class = "kpi-value magenta" if style == "magenta" else "kpi-value"
        with col:
            st.markdown(f"""
            <div class="kpi-card {accent}">
                <div class="{val_class}">{val}</div>
                <div class="kpi-label">{label}</div>
            </div>
            """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Tabs ───────────────────────────────────────────────────────────────────
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "📊 Channel", "🖥 Platform", "📦 Product", "📈 Trends", "💡 Insights"
    ])

    with tab1:
        st.markdown('<div class="section-header">Channel Performance</div>', unsafe_allow_html=True)
        if "Channel" in dff.columns:
            ch_df = dff.groupby("Channel").agg(
                Spend=("Spend","sum"),
                Impressions=("Impressions","sum"),
                Clicks=("Clicks","sum"),
                Conversions=("Conversion","sum"),
            ).reset_index().sort_values("Spend", ascending=False)
            ch_df["CTR"] = ch_df.apply(
                lambda r: fmt_pct(r["Clicks"]/r["Impressions"]) if r["Impressions"]>0 else "—", axis=1)
            ch_df["CPA"] = ch_df.apply(
                lambda r: fmt_currency(r["Spend"]/r["Conversions"], 2) if r["Conversions"]>0 else "—", axis=1)
            ch_df["Spend_fmt"] = ch_df["Spend"].apply(fmt_currency)
            ch_df["Impressions_fmt"] = ch_df["Impressions"].apply(fmt_int)
            ch_df["Conversions_fmt"] = ch_df["Conversions"].apply(fmt_int)

            display_ch = ch_df[["Channel","Spend_fmt","Impressions_fmt","Conversions_fmt","CTR","CPA"]].copy()
            display_ch.columns = ["Channel","Spend","Impressions","Conversions","CTR","CPA"]
            st.dataframe(display_ch, use_container_width=True, hide_index=True)

            import plotly.express as px
            fig = px.bar(ch_df, x="Spend", y="Channel", orientation="h",
                         color_discrete_sequence=["#1A2B4A"],
                         title="Spend by Channel")
            fig.update_layout(
                plot_bgcolor="white", paper_bgcolor="white",
                font_family="Avenir, Calibri",
                title_font_size=14, title_font_color="#1A2B4A",
                xaxis_title="", yaxis_title="",
                margin=dict(l=0,r=0,t=40,b=0),
                height=300,
            )
            fig.update_traces(marker_line_width=0)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No 'Channel' column found in this dataset.")

    with tab2:
        st.markdown('<div class="section-header">Platform Performance</div>', unsafe_allow_html=True)
        if "Platform" in dff.columns:
            pl_df = dff.groupby("Platform").agg(
                Spend=("Spend","sum"), Impressions=("Impressions","sum"),
                Clicks=("Clicks","sum"), Conversions=("Conversion","sum"),
            ).reset_index().sort_values("Spend", ascending=False)
            pl_df["CTR"] = pl_df.apply(
                lambda r: fmt_pct(r["Clicks"]/r["Impressions"]) if r["Impressions"]>0 else "—", axis=1)
            pl_df["CPA"] = pl_df.apply(
                lambda r: fmt_currency(r["Spend"]/r["Conversions"], 2) if r["Conversions"]>0 else "—", axis=1)

            c1, c2 = st.columns([1.2, 1])
            with c1:
                display_pl = pl_df.copy()
                display_pl["Spend"] = display_pl["Spend"].apply(fmt_currency)
                display_pl["Impressions"] = display_pl["Impressions"].apply(fmt_int)
                display_pl["Conversions"] = display_pl["Conversions"].apply(fmt_int)
                st.dataframe(display_pl[["Platform","Spend","Impressions","Conversions","CTR","CPA"]],
                             use_container_width=True, hide_index=True)
            with c2:
                import plotly.express as px
                pl_plot = pl_df[pl_df["Conversions"] > 0].copy()
                if pl_plot.empty:
                    pl_plot = pl_df.copy()
                fig2 = px.scatter(pl_plot, x="Spend", y="Conversions", size="Spend",
                                  text="Platform", color_discrete_sequence=["#E81F76"],
                                  title="Spend vs Conversions by Platform")
                fig2.update_layout(
                    plot_bgcolor="white", paper_bgcolor="white",
                    font_family="Avenir, Calibri", height=280,
                    title_font_size=13, margin=dict(l=0,r=0,t=40,b=0),
                )
                fig2.update_traces(textposition="top center", marker_line_width=0)
                st.plotly_chart(fig2, use_container_width=True)
        else:
            st.info("No 'Platform' column found in this dataset.")

    with tab3:
        st.markdown('<div class="section-header">Product Performance</div>', unsafe_allow_html=True)
        prod_col = next((c for c in ["Product","product","LOB","lob","Campaign_Initiative"] if c in dff.columns), None)
        if prod_col:
            pr_df = dff.groupby(prod_col).agg(
                Spend=("Spend","sum"), Conversions=("Conversion","sum"),
                Clicks=("Clicks","sum"),
            ).reset_index().sort_values("Spend", ascending=False).head(10)
            pr_df["CPA"] = pr_df.apply(
                lambda r: r["Spend"]/r["Conversions"] if r["Conversions"]>0 else None, axis=1)
            pr_df["CPA_fmt"] = pr_df["CPA"].apply(lambda v: fmt_currency(v,2) if v else "—")
            pr_df["Spend_fmt"] = pr_df["Spend"].apply(fmt_currency)
            pr_df["Conv_fmt"] = pr_df["Conversions"].apply(fmt_int)

            c1, c2 = st.columns([1, 1.2])
            with c1:
                st.dataframe(
                    pr_df[[prod_col,"Spend_fmt","Conv_fmt","CPA_fmt"]].rename(columns={
                        prod_col: "Segment", "Spend_fmt": "Spend",
                        "Conv_fmt": "Conversions", "CPA_fmt": "CPA"
                    }),
                    use_container_width=True, hide_index=True
                )
            with c2:
                import plotly.express as px
                pr_plot = pr_df.dropna(subset=["CPA"]).head(8)
                if not pr_plot.empty:
                    fig3 = px.bar(pr_plot, x=prod_col, y="CPA",
                                  color_discrete_sequence=["#E81F76"],
                                  title="Cost Per Acquisition by Segment")
                    fig3.update_layout(
                        plot_bgcolor="white", paper_bgcolor="white",
                        font_family="Avenir, Calibri", height=280,
                        xaxis_title="", yaxis_title="CPA ($)",
                        title_font_size=13, margin=dict(l=0,r=0,t=40,b=0),
                    )
                    fig3.update_traces(marker_line_width=0)
                    st.plotly_chart(fig3, use_container_width=True)
        else:
            st.info("No product/segment column detected.")

    with tab4:
        st.markdown('<div class="section-header">Monthly Trends</div>', unsafe_allow_html=True)
        if "Date" in dff.columns:
            monthly = dff.copy()
            monthly["Month"] = monthly["Date"].dt.to_period("M")
            trend = monthly.groupby("Month").agg(
                Spend=("Spend","sum"), Conversions=("Conversion","sum"),
                Clicks=("Clicks","sum"), Impressions=("Impressions","sum"),
            ).reset_index()
            trend["Month_str"] = trend["Month"].astype(str)
            trend["CPA"] = trend.apply(
                lambda r: r["Spend"]/r["Conversions"] if r["Conversions"]>0 else None, axis=1)

            import plotly.graph_objects as go
            fig4 = go.Figure()
            fig4.add_trace(go.Bar(
                x=trend["Month_str"], y=trend["Spend"],
                name="Spend", marker_color="#1A2B4A", yaxis="y1"
            ))
            fig4.add_trace(go.Scatter(
                x=trend["Month_str"], y=trend["Conversions"],
                name="Conversions", line=dict(color="#E81F76", width=3),
                mode="lines+markers", marker=dict(size=8), yaxis="y2"
            ))
            fig4.update_layout(
                plot_bgcolor="white", paper_bgcolor="white",
                font_family="Avenir, Calibri",
                title="Monthly Spend vs. Conversions",
                title_font_size=14, title_font_color="#1A2B4A",
                yaxis=dict(title="Spend ($)", showgrid=True, gridcolor="#F0F0F0"),
                yaxis2=dict(title="Conversions", overlaying="y", side="right"),
                legend=dict(orientation="h", yanchor="bottom", y=1.02),
                margin=dict(l=0,r=0,t=60,b=0),
                height=350,
                barmode="group",
            )
            st.plotly_chart(fig4, use_container_width=True)

            peak_spend_row = trend.loc[trend["Spend"].idxmax()]
            peak_conv_row  = trend.loc[trend["Conversions"].idxmax()]
            m1, m2 = st.columns(2)
            with m1:
                st.markdown(f"""
                <div class="insight-card">
                    <div class="insight-title">Peak Spend Month</div>
                    <div class="insight-text">{peak_spend_row["Month_str"]} — {fmt_currency(peak_spend_row["Spend"])}</div>
                </div>
                """, unsafe_allow_html=True)
            with m2:
                st.markdown(f"""
                <div class="insight-card">
                    <div class="insight-title">Peak Conversion Month</div>
                    <div class="insight-text">{peak_conv_row["Month_str"]} — {fmt_int(peak_conv_row["Conversions"])} conversions</div>
                </div>
                """, unsafe_allow_html=True)
        else:
            st.info("No 'Date' column found for trend analysis.")

    with tab5:
        st.markdown('<div class="section-header">Automated Insights</div>', unsafe_allow_html=True)
        insights = generate_insights(dff, kpis)
        if insights:
            for ins in insights:
                st.markdown(f"""
                <div class="insight-card">
                    <div class="insight-title">{ins["title"]}</div>
                    <div class="insight-text">{ins["text"]}</div>
                </div>
                """, unsafe_allow_html=True)
        else:
            st.info("Load data to generate insights.")

    # ── Generate deck ──────────────────────────────────────────────────────────
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<div class="section-header">Generate Deck</div>', unsafe_allow_html=True)

    gen_col1, gen_col2 = st.columns([2, 1])
    with gen_col1:
        lob_label = selected_lob if selected_lob != "All" else report_title
        st.markdown(f"""
        <div style="background:#F0F4FF;border-left:4px solid #1A2B4A;padding:0.75rem 1rem;border-radius:0 6px 6px 0;font-size:0.9rem;color:#1A2B4A;">
            Ready to generate: <strong>{client_name} — {lob_label}</strong><br>
            <span style="color:#8C8C8C;font-size:0.8rem;">{len(dff):,} rows | {fmt_currency(kpis['spend'])} spend | {fmt_int(kpis['conversions'])} conversions</span>
        </div>
        """, unsafe_allow_html=True)

    with gen_col2:
        if st.button("⚡ Generate Performance Deck", use_container_width=True):
            with st.spinner("Building your deck..."):
                config = {
                    "lob": lob_label,
                    "client": client_name,
                    "date_range": date_range_label,
                }
                buf, err = generate_deck(dff, config)
            if err:
                st.error(f"Error: {err}")
            elif buf:
                st.session_state["deck_buf"] = buf
                st.session_state["deck_name"] = f"{client_name}_{lob_label.replace(' ','_')}_Performance.pptx"
                st.success("Deck generated!")

    if "deck_buf" in st.session_state:
        st.download_button(
            label="⬇ Download PPTX",
            data=st.session_state["deck_buf"],
            file_name=st.session_state.get("deck_name", "performance_deck.pptx"),
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
