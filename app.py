"""
Trade Governance Lab — National & Economic Security in the WTO.

Reads Database_v21.xlsx (sheets: Database, Vocabularies, Issues_Log) and presents five tabs:
Overview · Members · Measures · Framing · Data & method.

Design notes
------------
* Stance is the colour language of the whole dashboard: Apprehension (ochre),
  Defence/Explanation (teal), Proposal/Recommendation (olive), General Statement (slate).
  The same four colours mean the same four things in every chart, so a reader learns the
  key once.
* Measure grouping is done in the app (see MEASURE_GROUPS) rather than in the workbook.
  The full mapping is shown and downloadable on the "Data & method" tab; if a
  "Measure_Group" column is ever added to the workbook it takes precedence automatically.
* Summaries are computed deterministically from the rows in view — no API key, no network
  call, and they cannot state a number the data does not contain.
* Counting: a row is one interaction (one member, one intervention). Measures, security
  sub-domains and governance topics are multi-valued, so charts built on them count
  *mentions* and are labelled as such.
"""

from __future__ import annotations

import re
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.io as pio
import streamlit as st

# ======================================================================================
# Config and theme
# ======================================================================================
st.set_page_config(page_title="Trade Governance Lab", page_icon="🛡️", layout="wide")

SHOW_TIME_CHARTS = False          # seven months of event-driven data will not carry a trend
UNSPECIFIED = "Unspecified Measure"

INK = "#16232E"          # headings, display type
PRIMARY = "#2C5F7C"      # accents, links, metric values
MUTED = "#5F6E78"        # secondary text
RULE = "#DBE2E6"         # hairlines
BAR = "#4A7E9B"          # neutral single-series bars

# Stance carries meaning, so it carries colour.
STANCE_ORDER = ["Apprehension", "Defence/Explanation", "Proposal/Recommendation", "General Statement"]
STANCE_COLORS = {
    "Apprehension": "#C1662F",
    "Defence/Explanation": "#2F7E8C",
    "Proposal/Recommendation": "#5C8A4A",
    "General Statement": "#8B99A6",
}
CAT_PALETTE = ["#2C5F7C", "#C1662F", "#5C8A4A", "#8B6BA8", "#2F7E8C", "#B0894A", "#8B99A6", "#9B5B6B"]
HEAT_SCALE = ["#F5F7F8", "#CBDBE3", "#93B8C9", "#5A8FA8", "#2C5F7C"]

PCONF = {"displayModeBar": False, "responsive": True}

FORUM_SHORT = {
    "General Council": "GC",
    "CTG": "CTG",
    "CMA": "CMA",
    "SCM Committee": "SCM",
    "TBT Committee": "TBT",
    "TRIMS Committee": "TRIMS",
    "Council for TRIPS / WGTTT": "TRIPS/WGTTT",
}
FORUM_FULL = {
    "GC": "General Council",
    "CTG": "Council for Trade in Goods",
    "CMA": "Committee on Market Access",
    "SCM": "Committee on Subsidies and Countervailing Measures",
    "TBT": "Committee on Technical Barriers to Trade",
    "TRIMS": "Committee on Trade-Related Investment Measures",
    "TRIPS/WGTTT": "Council for TRIPS / Working Group on Trade and Transfer of Technology",
}
MEMBER_SHORT = {
    "Bolivarian Republic of Venezuela": "Venezuela",
    "Russian Federation": "Russia",
    "Republic of Korea": "Korea, Rep. of",
    "Gambia (on behalf of the LDC Group)": "Gambia (LDC Group)",
    "Mozambique (on behalf of the African Group)": "Mozambique (African Group)",
}

# ---- Measure grouping rules (ordered; first match wins) -------------------------------
MEASURE_GROUPS: list[tuple[str, list[str]]] = [
    ("Sanctions & coercive measures",
     [r"sanction", r"coercive", r"russia", r"oil price cap", r"secondary tariffs", r"restrictive measures"]),
    ("Export controls & restrictions",
     [r"export control", r"export restriction", r"export permit", r"entity list", r"dual-use",
      r"rare earth", r"critical raw material", r"critical mineral", r"downstream processing"]),
    ("Tariffs & trade remedies",
     [r"tariff", r"section 232", r"section 301", r"surtax", r"trq", r"customs", r"countervailing",
      r"\bcvd\b", r"steel trade measures", r"steel import"]),
    ("Industrial policy & subsidies",
     [r"chips", r"\bact\b", r"subsid", r"overcapacity", r"local content", r"made in", r"golden share",
      r"guardrail", r"decree", r"\bfund\b", r"support for food", r"accelerator", r"net-zero"]),
    ("Digital, data & technology",
     [r"encryption", r"cryptograph", r"cybersecurity", r"e-commerce", r"semiconductor"]),
    ("Green & environmental conditions",
     [r"eudr", r"green protectionist", r"photovoltaic"]),
    ("General / unspecified",
     [r"^unspecified", r"^unilateral measures$", r"^tariff and non-tariff"]),
]

pio.templates["tgl"] = pio.templates["plotly_white"]
pio.templates["tgl"].layout.update(
    colorway=CAT_PALETTE,
    font=dict(family="IBM Plex Sans, Segoe UI, system-ui, sans-serif", size=13, color="#3A464F"),
    margin=dict(l=10, r=18, t=52, b=16),
    title=dict(font=dict(family="Source Serif 4, Georgia, serif", size=16, color=INK), x=0, xanchor="left"),
    xaxis=dict(automargin=True, gridcolor="#EDF1F3", zerolinecolor="#EDF1F3"),
    yaxis=dict(automargin=True, gridcolor="#EDF1F3", zerolinecolor="#EDF1F3"),
    legend=dict(orientation="h", yanchor="top", y=-0.14, xanchor="center", x=0.5, title_text=""),
)
pio.templates.default = "tgl"

st.markdown(
    f"""
    <style>
      @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@500;600&family=IBM+Plex+Sans:wght@400;500;600&family=Source+Serif+4:opsz,wght@8..60,600;8..60,700&display=swap');

      html, body, .stApp, [data-testid="stAppViewContainer"], [data-testid="stSidebar"],
      .stApp p, .stApp li, .stApp label, .stApp span, .stMarkdown, .stCaption, button, input {{
          font-family: 'IBM Plex Sans', Segoe UI, system-ui, sans-serif;
      }}
      h1, h2, h3, h4 {{ font-family: 'Source Serif 4', Georgia, serif; color: {INK}; letter-spacing: -0.01em; }}
      .block-container {{ padding-top: 2.2rem; padding-bottom: 2.5rem; max-width: 1480px; }}

      .masthead {{ border-bottom: 2px solid {INK}; padding-bottom: 10px; margin-bottom: 6px; }}
      .masthead .eyebrow {{ font-family:'IBM Plex Mono', monospace; font-size:.7rem; font-weight:600;
                            letter-spacing:.16em; text-transform:uppercase; color:{PRIMARY}; }}
      .masthead h1 {{ font-size: clamp(1.7rem, 3.4vw, 2.35rem); margin:.15rem 0 .2rem 0; font-weight:700; }}
      .masthead .sub {{ color:{MUTED}; font-size:.98rem; max-width:70ch; margin-bottom:.35rem; }}
      .masthead .meta {{ font-family:'IBM Plex Mono', monospace; font-size:.74rem; color:{MUTED};
                         letter-spacing:.02em; }}

      /* Stance strip — the composition of the current view, always visible */
      .strip {{ display:flex; width:100%; height:12px; border-radius:6px; overflow:hidden;
                margin:14px 0 6px 0; border:1px solid {RULE}; }}
      .strip div {{ height:100%; }}
      .strip-key {{ font-family:'IBM Plex Mono', monospace; font-size:.72rem; color:{MUTED};
                    display:flex; flex-wrap:wrap; gap:14px; margin-bottom:16px; }}
      .strip-key span b {{ font-weight:600; color:{INK}; }}
      .dot {{ display:inline-block; width:9px; height:9px; border-radius:2px; margin-right:5px; }}

      div[data-testid="stMetricValue"] {{ font-family:'IBM Plex Mono', monospace;
                                          color:{PRIMARY}; font-size:1.55rem; }}
      div[data-testid="stMetricLabel"] p {{ font-size:.76rem; letter-spacing:.05em;
                                            text-transform:uppercase; color:{MUTED}; }}

      .stTabs [data-baseweb="tab-list"] {{ gap:2px; flex-wrap:wrap; border-bottom:1px solid {RULE}; }}
      .stTabs [data-baseweb="tab"] {{ font-weight:600; font-size:.92rem; padding:9px 18px;
                                      border-radius:8px 8px 0 0; color:{MUTED}; }}
      .stTabs [aria-selected="true"] {{ background:{PRIMARY}0F; color:{PRIMARY};
                                        border-bottom:3px solid {PRIMARY}; }}

      .read {{ background:#FFFFFF; border:1px solid {RULE}; border-left:4px solid {PRIMARY};
               border-radius:4px; padding:14px 18px; margin:4px 0 20px 0;
               font-size:.95rem; line-height:1.6; color:#2B3740; }}
      .read .tag {{ font-family:'IBM Plex Mono', monospace; font-size:.68rem; font-weight:600;
                    letter-spacing:.14em; text-transform:uppercase; color:{PRIMARY};
                    display:block; margin-bottom:7px; }}
      .read b {{ color:{INK}; font-weight:600; }}
      .note {{ font-size:.82rem; color:{MUTED}; margin:-6px 0 14px 0; }}

      section[data-testid="stSidebar"] {{ background:#F7F9FA; border-right:1px solid {RULE}; }}
      section[data-testid="stSidebar"] h2 {{ font-size:1.05rem; }}

      @media (max-width: 640px) {{
          .block-container {{ padding-left:.7rem; padding-right:.7rem; }}
          div[data-testid="stMetricValue"] {{ font-size:1.15rem; }}
      }}
    </style>
    """,
    unsafe_allow_html=True,
)


# ======================================================================================
# Data
# ======================================================================================
CANDIDATE_FILES = ["Database_v21.xlsx", "Database.xlsx", "WTO_Database.xlsx"]


def find_workbook() -> Path | None:
    here = Path(__file__).parent
    for name in CANDIDATE_FILES:
        p = here / name
        if p.exists():
            return p
    matches = sorted(here.glob("Database*.xlsx"))
    return matches[-1] if matches else None


def group_measure(name: str) -> str:
    """Fold near-duplicate measure labels into families. First matching rule wins."""
    if not isinstance(name, str) or not name.strip():
        return "General / unspecified"
    low = name.lower()
    for group, patterns in MEASURE_GROUPS:
        if any(re.search(p, low) for p in patterns):
            return group
    return "Other measures"


@st.cache_data(show_spinner=False)
def load_data(path_str: str, mtime: float):
    """mtime busts the cache when the workbook is replaced."""
    xl = pd.ExcelFile(path_str)
    df = pd.read_excel(xl, sheet_name="Database")
    df.columns = df.columns.str.strip()

    vocab = pd.read_excel(xl, sheet_name="Vocabularies") if "Vocabularies" in xl.sheet_names else pd.DataFrame()
    issues = pd.read_excel(xl, sheet_name="Issues_Log") if "Issues_Log" in xl.sheet_names else pd.DataFrame()

    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df["Row_ID"] = range(len(df))
    df["Forum"] = df["WTO_Forum"].map(FORUM_SHORT).fillna(df["WTO_Forum"])
    df["Member"] = df["Participant"].map(MEMBER_SHORT).fillna(df["Participant"])
    df["Owner"] = df["Measure_Owner"].map(MEMBER_SHORT).fillna(df["Measure_Owner"])
    for col in ["Stance", "Security_Relevance", "Confidence"]:
        if col in df.columns:
            df[col] = df[col].astype("string").str.strip()

    # Inclusion rules R1..R7 as flags, for the method tab.
    rules = df.get("Inclusion_Rule", pd.Series([""] * len(df))).fillna("").astype(str)
    df["Rule_Codes"] = rules.apply(lambda s: sorted(set(re.findall(r"R\d", s))))

    return df, vocab, issues


def measures_long(d: pd.DataFrame) -> pd.DataFrame:
    """One row per (interaction, measure). Counts here are mentions, not interactions."""
    keep = ["Row_ID", "Member", "Owner", "Forum", "Stance", "Date", "Document_Symbol"]
    parts = []
    for col in ["Measure 1", "Measure 2", "Measure 3"]:
        if col not in d.columns:
            continue
        sub = d[keep].copy()
        sub["Measure"] = d[col].astype("string").str.strip()
        parts.append(sub)
    if not parts:
        return pd.DataFrame(columns=keep + ["Measure", "Measure_Group"])
    out = pd.concat(parts, ignore_index=True).dropna(subset=["Measure"])
    out = out[out["Measure"] != ""]
    if "Measure_Group" in d.columns:                      # workbook column wins if it appears later
        lookup = d.set_index("Row_ID")["Measure_Group"]
        out["Measure_Group"] = out["Row_ID"].map(lookup).fillna(out["Measure"].map(group_measure))
    else:
        out["Measure_Group"] = out["Measure"].map(group_measure)
    return out.drop_duplicates(subset=["Row_ID", "Measure"]).reset_index(drop=True)


def subdomains_long(d: pd.DataFrame) -> pd.DataFrame:
    keep = ["Row_ID", "Member", "Forum", "Stance", "Owner"]
    parts = []
    for col, rank in [("Security_SubDomain_1", "Primary"), ("Security_SubDomain_2", "Secondary")]:
        if col not in d.columns:
            continue
        sub = d[keep].copy()
        sub["SubDomain"] = d[col].astype("string").str.strip()
        sub["Rank"] = rank
        parts.append(sub)
    if not parts:
        return pd.DataFrame(columns=keep + ["SubDomain", "Rank"])
    out = pd.concat(parts, ignore_index=True).dropna(subset=["SubDomain"])
    out = out[out["SubDomain"] != ""].drop_duplicates(subset=["Row_ID", "SubDomain"])
    return out.reset_index(drop=True)


def topics_long(d: pd.DataFrame) -> pd.DataFrame:
    col = "Governance_Dimensions_Topics"
    keep = ["Row_ID", "Member", "Forum", "Stance"]
    if col not in d.columns:
        return pd.DataFrame(columns=keep + ["Dimension", "Topic", "Full_Topic"])
    sub = d[keep + [col]].copy()
    sub[col] = sub[col].astype("string")
    sub = sub.dropna(subset=[col])
    sub["Full_Topic"] = sub[col].str.split("|")
    out = sub.explode("Full_Topic")
    out["Full_Topic"] = out["Full_Topic"].str.strip()
    out = out[out["Full_Topic"].str.len() > 0]
    split = out["Full_Topic"].str.split(":", n=1, expand=True)
    out["Dimension"] = split[0].str.strip()
    out["Topic"] = split[1].str.strip() if split.shape[1] > 1 else split[0].str.strip()
    out = out.drop(columns=[col]).drop_duplicates(subset=["Row_ID", "Full_Topic"])
    return out.reset_index(drop=True)


# ======================================================================================
# Small helpers
# ======================================================================================
def vc(series: pd.Series, top: int | None = None) -> pd.DataFrame:
    out = series.dropna().value_counts()
    if top:
        out = out.head(top)
    out = out.reset_index()
    out.columns = ["label", "count"]
    return out


def pct(part: int, whole: int) -> float:
    return 0.0 if not whole else round(part / whole * 100, 1)


def pcs(part: int, whole: int) -> str:
    """Percentage as display text — 54.5%, but 100% rather than 100.0%."""
    return f"{pct(part, whole):g}%"


def show(fig, key: str, container=None):
    target = container or st
    target.plotly_chart(fig, width="stretch", config=PCONF, key=key)


def int_axis(fig, maxval, axis: str = "x"):
    upd = fig.update_xaxes if axis == "x" else fig.update_yaxes
    if not maxval or maxval <= 1:
        upd(tickformat="d", dtick=1, rangemode="tozero")
    elif maxval <= 10:
        upd(tickformat="d", dtick=1)
    else:
        upd(tickformat="d")
    return fig


def hbar(data: pd.DataFrame, title: str, xlabel: str = "", height: int | None = None):
    data = data.sort_values("count")
    h = height or max(230, 30 * len(data) + 110)
    fig = px.bar(data, x="count", y="label", orientation="h", title=title)
    fig.update_traces(marker_color=BAR, hovertemplate="%{y}: %{x}<extra></extra>")
    fig.update_layout(height=h, yaxis_title=None, xaxis_title=xlabel or None, showlegend=False)
    int_axis(fig, data["count"].max() if len(data) else 0)
    return fig


def stance_hbar(data: pd.DataFrame, ycol: str, title: str, normalise: bool = False,
                height: int | None = None, xlabel: str = ""):
    """Horizontal bars split by stance. `data` needs columns [ycol, 'Stance', 'count']."""
    order = data.groupby(ycol)["count"].sum().sort_values().index.tolist()
    h = height or max(250, 28 * len(order) + 130)
    fig = px.bar(
        data, x="count", y=ycol, color="Stance", orientation="h", title=title,
        category_orders={ycol: order, "Stance": STANCE_ORDER},
        color_discrete_map=STANCE_COLORS,
    )
    fig.update_layout(height=h, barmode="relative", yaxis_title=None,
                      xaxis_title=xlabel or None, margin=dict(t=52, b=70, l=10, r=18))
    if normalise:
        fig.update_layout(barnorm="percent")
        fig.update_xaxes(ticksuffix="%", title=None)
    else:
        int_axis(fig, data.groupby(ycol)["count"].sum().max() if len(data) else 0)
    return fig


def heatmap(matrix: pd.DataFrame, title: str, height: int = 400, xlab: str = "", ylab: str = ""):
    fig = px.imshow(matrix, aspect="auto", text_auto=True,
                    color_continuous_scale=HEAT_SCALE, title=title)
    fig.update_layout(height=height, coloraxis_showscale=False,
                      xaxis_title=xlab or None, yaxis_title=ylab or None,
                      margin=dict(t=52, b=20, l=10, r=18))
    fig.update_xaxes(tickangle=-30, side="bottom")
    fig.update_traces(hovertemplate="%{y} → %{x}: %{z}<extra></extra>")
    return fig


def readout(text: str):
    st.markdown(f"<div class='read'><span class='tag'>Summary of this view</span>{text}</div>",
                unsafe_allow_html=True)


def note(text: str):
    st.markdown(f"<div class='note'>{text}</div>", unsafe_allow_html=True)


def stance_strip(d: pd.DataFrame):
    counts = d["Stance"].value_counts()
    total = int(counts.sum())
    if not total:
        return
    bars, keys = [], []
    for s in STANCE_ORDER:
        n = int(counts.get(s, 0))
        if not n:
            continue
        bars.append(f"<div style='width:{n / total * 100:.3f}%;background:{STANCE_COLORS[s]};'></div>")
        keys.append(f"<span><span class='dot' style='background:{STANCE_COLORS[s]}'></span>"
                    f"{s} <b>{n}</b> · {pct(n, total)}%</span>")
    st.markdown(f"<div class='strip'>{''.join(bars)}</div>"
                f"<div class='strip-key'>{''.join(keys)}</div>", unsafe_allow_html=True)


def top_label(series: pd.Series) -> tuple[str, int]:
    s = series.dropna()
    if s.empty:
        return "—", 0
    counts = s.value_counts()
    return str(counts.index[0]), int(counts.iloc[0])


def joined(items, limit: int = 3) -> str:
    items = [str(i) for i in items][:limit]
    if not items:
        return "—"
    if len(items) == 1:
        return items[0]
    return ", ".join(items[:-1]) + " and " + items[-1]


def small_n(d: pd.DataFrame) -> str:
    return (" <b>Small sample</b> — with fewer than 20 interactions in view, read shares as"
            " illustrative rather than representative." if len(d) < 20 else "")


def concentration_note(d: pd.DataFrame) -> str:
    """Flag when one meeting record dominates the view — G/C/M/153 is a third of the corpus."""
    if d.empty or "Document_Symbol" not in d.columns:
        return ""
    doc, n = top_label(d["Document_Symbol"])
    share = pct(n, len(d))
    if share >= 30 and len(d) > 10:
        return (f" Note that <b>{doc}</b> alone accounts for {share}% of the rows in view,"
                " so member rankings partly reflect who attended that one meeting.")
    return ""


# ======================================================================================
# Load
# ======================================================================================
wb = find_workbook()
if wb is None:
    st.error(
        "**Workbook not found.** Place `Database_v21.xlsx` in the same folder as `app.py` "
        "(the app also accepts `Database.xlsx` or any `Database*.xlsx`), then reload."
    )
    st.stop()

df, vocab, issues = load_data(str(wb), wb.stat().st_mtime)

DOMAIN = df["Domain"].dropna().iloc[0] if "Domain" in df.columns and df["Domain"].notna().any() else "—"
COVER_FROM = df["Date"].min()
COVER_TO = df["Date"].max()

# ======================================================================================
# Filters
# ======================================================================================
st.sidebar.header("Filters")
st.sidebar.caption("Leave a filter empty to include everything. Filters apply to every tab.")

all_subdomains = sorted(subdomains_long(df)["SubDomain"].unique())
all_groups = sorted(measures_long(df)["Measure_Group"].unique())

f_forum = st.sidebar.multiselect("WTO body", sorted(df["Forum"].dropna().unique()))
f_member = st.sidebar.multiselect("Member speaking", sorted(df["Member"].dropna().unique()))
f_stance = st.sidebar.multiselect("Stance", [s for s in STANCE_ORDER if s in set(df["Stance"].dropna())])
f_sub = st.sidebar.multiselect("Security sub-domain", all_subdomains,
                               help="Matches either the primary or the secondary sub-domain.")
f_group = st.sidebar.multiselect("Measure family", all_groups)
f_owner = st.sidebar.multiselect("Measure owner", sorted(df["Owner"].dropna().unique()),
                                 help="The member whose measure is being discussed — not the speaker.")

core_only = st.sidebar.toggle("Core relevance only", value=False,
                              help="Keep only rows flagged Core in Security_Relevance.")

d_from = d_to = None
if pd.notna(COVER_FROM) and pd.notna(COVER_TO):
    picked = st.sidebar.date_input(
        "Date range", value=(COVER_FROM.date(), COVER_TO.date()),
        min_value=COVER_FROM.date(), max_value=COVER_TO.date(),
    )
    # Streamlit returns a one-item tuple while the user is mid-selection.
    if isinstance(picked, (list, tuple)):
        d_from = picked[0] if picked else COVER_FROM.date()
        d_to = picked[1] if len(picked) > 1 else COVER_TO.date()
    else:
        d_from = d_to = picked

filtered = df.copy()
if f_forum:
    filtered = filtered[filtered["Forum"].isin(f_forum)]
if f_member:
    filtered = filtered[filtered["Member"].isin(f_member)]
if f_stance:
    filtered = filtered[filtered["Stance"].isin(f_stance)]
if f_owner:
    filtered = filtered[filtered["Owner"].isin(f_owner)]
if core_only and "Security_Relevance" in filtered.columns:
    filtered = filtered[filtered["Security_Relevance"] == "Core"]
if f_sub:
    ids = set(subdomains_long(df).loc[lambda x: x["SubDomain"].isin(f_sub), "Row_ID"])
    filtered = filtered[filtered["Row_ID"].isin(ids)]
if f_group:
    ids = set(measures_long(df).loc[lambda x: x["Measure_Group"].isin(f_group), "Row_ID"])
    filtered = filtered[filtered["Row_ID"].isin(ids)]
if d_from and d_to:
    filtered = filtered[(filtered["Date"] >= pd.Timestamp(d_from)) & (filtered["Date"] <= pd.Timestamp(d_to))]

st.sidebar.markdown("---")
st.sidebar.metric("Interactions in view", f"{len(filtered)} / {len(df)}")
if st.sidebar.button("Reset filters", width="stretch", key="reset"):
    for k in list(st.session_state.keys()):
        del st.session_state[k]
    st.rerun()

# ======================================================================================
# Masthead
# ======================================================================================
cover = (f"{COVER_FROM:%d %b %Y} – {COVER_TO:%d %b %Y}"
         if COVER_FROM is not pd.NaT else "—")
st.markdown(
    f"""
    <div class="masthead">
      <div class="eyebrow">WTO discussion analytics</div>
      <h1>Trade Governance Lab</h1>
      <div class="sub">How WTO members argue about <b>{DOMAIN.lower()}</b> — which measures they
      raise, whose measures they raise them about, and the grounds they argue on.</div>
      <div class="meta">{len(df)} interactions · {df['Document_Symbol'].nunique()} documents ·
      {df['Participant'].nunique()} members · {df['Forum'].nunique()} WTO bodies · {cover}</div>
    </div>
    """,
    unsafe_allow_html=True,
)

if filtered.empty:
    st.warning("No interactions match the current filters. Use **Reset filters** in the sidebar.")
    st.stop()

stance_strip(filtered)

M = measures_long(filtered)
S = subdomains_long(filtered)
T = topics_long(filtered)

tab_over, tab_mem, tab_meas, tab_frame, tab_data = st.tabs(
    ["Overview", "Members", "Measures", "Framing", "Data & method"]
)

# ======================================================================================
# OVERVIEW
# ======================================================================================
with tab_over:
    named = M[M["Measure"] != UNSPECIFIED]
    cols = st.columns(6)
    tiles = [
        ("Interactions", len(filtered), "One member intervention on one agenda item."),
        ("Members", filtered["Member"].nunique(), "Distinct members speaking."),
        ("Documents", filtered["Document_Symbol"].nunique(), "WTO documents the rows are drawn from."),
        ("WTO bodies", filtered["Forum"].nunique(), "Councils and committees involved."),
        ("Named measures", named["Measure"].nunique(), "Distinct measures identified by name."),
        ("Core relevance", pcs(int((filtered['Security_Relevance'] == 'Core').sum()), len(filtered)),
         "Share of rows where security is the core of the intervention rather than context."),
    ]
    for c, (label, value, help_text) in zip(cols, tiles):
        c.metric(label, value, help=help_text)

    top_stance, n_stance = top_label(filtered["Stance"])
    top_sub, n_sub = top_label(S["SubDomain"])
    top_grp, n_grp = top_label(named["Measure_Group"])
    top_mem, n_mem = top_label(filtered["Member"])
    top_forum, n_forum = top_label(filtered["Forum"])
    owners = filtered.loc[~filtered["Owner"].isin(["Not applicable"]), "Owner"]
    top_owner, n_owner = top_label(owners)

    readout(
        f"This view holds <b>{len(filtered)}</b> interactions by <b>{filtered['Member'].nunique()}</b> "
        f"members across <b>{filtered['Forum'].nunique()}</b> WTO bodies, drawn from "
        f"<b>{filtered['Document_Symbol'].nunique()}</b> documents. The prevailing stance is "
        f"<b>{top_stance}</b> ({pcs(n_stance, len(filtered))} of interactions). Discussion clusters on "
        f"<b>{top_sub}</b> ({n_sub} mentions) and, among named measures, on "
        f"<b>{top_grp}</b> ({n_grp} mentions). <b>{top_mem}</b> speaks most often ({n_mem} interactions), "
        f"while <b>{top_owner}</b> is the member whose measures are discussed most ({n_owner} interactions). "
        f"Most of this happens in <b>{FORUM_FULL.get(top_forum, top_forum)}</b> "
        f"({pcs(n_forum, len(filtered))}).{concentration_note(filtered)}{small_n(filtered)}"
    )

    c1, c2 = st.columns(2)
    sub_stance = (S.groupby(["SubDomain", "Stance"]).size().reset_index(name="count"))
    show(stance_hbar(sub_stance, "SubDomain", "Security sub-domains, by stance", xlabel="mentions"),
         "ov_sub", c1)
    c1.caption("A row can carry a primary and a secondary sub-domain, so mentions exceed interactions.")

    forum_stance = filtered.groupby(["Forum", "Stance"]).size().reset_index(name="count")
    show(stance_hbar(forum_stance, "Forum", "Where the discussion happens, by stance", xlabel="interactions"),
         "ov_forum", c2)
    c2.caption(" · ".join(f"**{k}** {v}" for k, v in FORUM_FULL.items() if k in set(filtered["Forum"])))

    grp_stance = named.groupby(["Measure_Group", "Stance"]).size().reset_index(name="count")
    if len(grp_stance):
        show(stance_hbar(grp_stance, "Measure_Group", "Measure families, by stance", xlabel="mentions"),
             "ov_group")
        note("Families are assembled in the app from the individual measure labels — the full "
             "mapping is on the <b>Data & method</b> tab.")

    if SHOW_TIME_CHARTS:
        monthly = (filtered.assign(Month=filtered["Date"].dt.to_period("M").dt.to_timestamp())
                   .groupby(["Month", "Stance"]).size().reset_index(name="count"))
        fig = px.bar(monthly, x="Month", y="count", color="Stance",
                     color_discrete_map=STANCE_COLORS, title="Interactions over time")
        show(fig, "ov_time")

# ======================================================================================
# MEMBERS
# ======================================================================================
with tab_mem:
    st.subheader("Who speaks, and how")

    n_members = filtered["Member"].nunique()
    if n_members > 6:
        top_n = st.slider("Members shown", 5, min(30, n_members), min(15, n_members), key="mem_topn")
    else:
        top_n = n_members
    normalise = st.toggle("Show stance as a share of each member's interactions", value=False,
                          key="mem_norm",
                          help="Useful when comparing members with very different totals.")

    keep_members = filtered["Member"].value_counts().head(top_n).index
    mem_stance = (filtered[filtered["Member"].isin(keep_members)]
                  .groupby(["Member", "Stance"]).size().reset_index(name="count"))

    counts = filtered["Member"].value_counts()
    apprehensive = (filtered[filtered["Stance"] == "Apprehension"]["Member"].value_counts()
                    .reindex(counts.index).fillna(0))
    defending = (filtered[filtered["Stance"] == "Defence/Explanation"]["Member"].value_counts()
                 .reindex(counts.index).fillna(0))
    vocal = counts[counts >= 3]
    critics = (apprehensive / counts).reindex(vocal.index).sort_values(ascending=False)
    defenders = (defending / counts).reindex(vocal.index).sort_values(ascending=False)

    readout(
        f"<b>{filtered['Member'].nunique()}</b> members take part. "
        f"{joined(counts.head(3).index)} speak most often "
        f"({joined([str(i) for i in counts.head(3).values])} interactions respectively). "
        + (f"Among members with at least three interactions, "
           f"<b>{critics.index[0]}</b> is the most consistently apprehensive "
           f"({pcs(int(apprehensive[critics.index[0]]), int(counts[critics.index[0]]))} of its interventions), "
           f"and <b>{defenders.index[0]}</b> spends the largest share of its interventions defending or "
           f"explaining measures ({pcs(int(defending[defenders.index[0]]), int(counts[defenders.index[0]]))}). "
           if len(vocal) else "")
        + concentration_note(filtered) + small_n(filtered)
    )

    show(stance_hbar(mem_stance, "Member", f"Top {len(keep_members)} members, by stance",
                     normalise=normalise, xlabel="" if normalise else "interactions"),
         "mem_stance")

    st.markdown("#### Who raises measures about whom")
    pairs = filtered[(filtered["Owner"].notna()) & (~filtered["Owner"].isin(["Not applicable"]))]
    if len(pairs) >= 5:
        mat = pd.crosstab(pairs["Member"], pairs["Owner"])
        mat = mat.loc[mat.sum(axis=1).sort_values(ascending=False).index]
        mat = mat[mat.sum().sort_values(ascending=False).index]
        mat = mat.iloc[:14, :10]
        show(heatmap(mat, "Speaker (row) → owner of the measure discussed (column)",
                     height=max(340, 26 * len(mat) + 150),
                     xlab="measure owner", ylab="member speaking"), "mem_matrix")
        self_ref = int((pairs["Member"] == pairs["Owner"]).sum())
        note(f"The diagonal is self-reference: <b>{self_ref}</b> interactions "
             f"({pcs(self_ref, len(pairs))} of attributed rows) are members speaking about their own "
             "measures — typically defence or explanation. Rows where the owner is "
             "<i>Not applicable</i> are excluded here.")
    else:
        note("Too few rows with an attributed measure owner in this view to draw the matrix.")

    st.markdown("#### Sub-domain focus by member")
    sub_mem = S[S["Member"].isin(keep_members[:12])]
    if len(sub_mem) >= 5:
        mat2 = pd.crosstab(sub_mem["Member"], sub_mem["SubDomain"])
        mat2 = mat2.loc[mat2.sum(axis=1).sort_values(ascending=False).index]
        show(heatmap(mat2, "Security sub-domains raised, by member",
                     height=max(320, 26 * len(mat2) + 140), xlab="sub-domain", ylab="member"),
             "mem_sub")
    else:
        note("Not enough sub-domain mentions in this view for a member breakdown.")

# ======================================================================================
# MEASURES
# ======================================================================================
with tab_meas:
    st.subheader("What is being discussed")

    hide_unspec = st.toggle("Hide unspecified measures", value=True, key="meas_hide",
                            help=f"'{UNSPECIFIED}' marks interventions with no named measure. "
                                 "They stay in the interaction count either way.")
    Mv = M[M["Measure"] != UNSPECIFIED] if hide_unspec else M

    if Mv.empty:
        note("No measures match the current filters.")
    else:
        counts = Mv["Measure"].value_counts()
        grp_counts = Mv["Measure_Group"].value_counts()
        unspec_n = int((M["Measure"] == UNSPECIFIED).sum())

        contested = (Mv.assign(is_app=Mv["Stance"] == "Apprehension")
                     .groupby("Measure")
                     .agg(mentions=("Row_ID", "nunique"), apprehension=("is_app", "sum")))
        floor = st.slider("Minimum mentions for the contested ranking", 1, 8, 3, key="meas_floor")
        contested = contested[contested["mentions"] >= floor]
        contested["share"] = (contested["apprehension"] / contested["mentions"] * 100).round(1)
        contested = contested.sort_values(["share", "mentions"], ascending=False)

        readout(
            f"<b>{Mv['Measure'].nunique()}</b> distinct measures appear in "
            f"<b>{len(Mv)}</b> mentions, falling into <b>{Mv['Measure_Group'].nunique()}</b> families. "
            f"The largest family is <b>{grp_counts.index[0]}</b> ({int(grp_counts.iloc[0])} mentions); "
            f"the most-discussed single measure is <b>{counts.index[0]}</b> ({int(counts.iloc[0])} mentions). "
            + (f"Of measures mentioned at least {floor} times, <b>{contested.index[0]}</b> draws the highest "
               f"share of apprehension ({contested['share'].iloc[0]:g}% of its mentions). "
               if len(contested) else "")
            + (f"A further <b>{unspec_n}</b> interactions in this view raise no named measure and are "
               f"{'excluded from' if hide_unspec else 'included in'} these charts."
               if unspec_n else "")
            + small_n(filtered)
        )

        c1, c2 = st.columns([3, 2])
        top_measures = vc(Mv["Measure"], top=14)
        show(hbar(top_measures, "Most-discussed measures", xlabel="mentions"), "meas_top", c1)
        show(hbar(vc(Mv["Measure_Group"]), "Measure families", xlabel="mentions"), "meas_grp", c2)

        meas_stance = (Mv[Mv["Measure"].isin(top_measures["label"])]
                       .groupby(["Measure", "Stance"]).size().reset_index(name="count"))
        show(stance_hbar(meas_stance, "Measure", "Stance towards each measure",
                         normalise=True), "meas_stance")
        note("Shares, so a measure raised twice sits beside one raised twelve times — read it with "
             "the mention counts above.")

        st.markdown("#### Whose measures draw discussion")
        owner_rows = filtered[~filtered["Owner"].isin(["Not applicable"]) & filtered["Owner"].notna()]
        if len(owner_rows):
            own_stance = owner_rows.groupby(["Owner", "Stance"]).size().reset_index(name="count")
            show(stance_hbar(own_stance, "Owner", "Measure owners, by the stance taken towards them",
                             xlabel="interactions"), "meas_owner")

        if len(contested):
            st.markdown("#### Most contested measures")
            disp = contested.reset_index().rename(columns={
                "Measure": "Measure", "mentions": "Mentions",
                "apprehension": "Apprehension", "share": "Apprehension %"})
            st.dataframe(disp, width="stretch", hide_index=True)

# ======================================================================================
# FRAMING
# ======================================================================================
with tab_frame:
    st.subheader("The grounds members argue on")

    if T.empty:
        note("No governance topics recorded for the rows in this view.")
    else:
        dim_counts = T["Dimension"].value_counts()
        topic_counts = T["Topic"].value_counts()
        per_row = T.groupby("Row_ID").size()

        legal = T[T["Dimension"] == "Legal"]["Row_ID"].nunique()
        readout(
            f"Members frame these interventions along <b>{T['Dimension'].nunique()}</b> governance "
            f"dimensions and <b>{T['Topic'].nunique()}</b> topics, averaging "
            f"<b>{per_row.mean():.1f}</b> topics per interaction. "
            f"<b>{dim_counts.index[0]}</b> dominates ({pcs(int(dim_counts.iloc[0]), int(dim_counts.sum()))} "
            f"of topic mentions), and the single most common ground is "
            f"<b>{topic_counts.index[0]}</b> ({int(topic_counts.iloc[0])} mentions). "
            f"<b>{legal}</b> interactions ({pcs(legal, len(filtered))}) argue at least partly on legal "
            f"grounds — WTO consistency, transparency or due process.{small_n(filtered)}"
        )

        c1, c2 = st.columns([2, 3])
        show(hbar(vc(T["Dimension"]), "Governance dimensions", xlabel="mentions"), "fr_dim", c1)
        show(hbar(vc(T["Topic"]), "Governance topics", xlabel="mentions"), "fr_topic", c2)

        topic_stance = T.groupby(["Topic", "Stance"]).size().reset_index(name="count")
        show(stance_hbar(topic_stance, "Topic", "Stance by governance topic", normalise=True),
             "fr_topic_stance")
        note("Topics are capped at three per interaction and ordered by relevance, per the rule "
             "adopted in the issues log.")

        if T["Forum"].nunique() > 1:
            mat = pd.crosstab(T["Topic"], T["Forum"])
            mat = mat.loc[mat.sum(axis=1).sort_values(ascending=False).index]
            show(heatmap(mat, "Which grounds are used in which body",
                         height=max(320, 28 * len(mat) + 140), xlab="WTO body", ylab="topic"),
                 "fr_matrix")

# ======================================================================================
# DATA & METHOD
# ======================================================================================
with tab_data:
    st.subheader("Reference tables and downloads")

    # --- summary reference tables ------------------------------------------------------
    measure_ref = (M.assign(is_app=M["Stance"] == "Apprehension",
                            is_def=M["Stance"] == "Defence/Explanation")
                   .groupby(["Measure_Group", "Measure"])
                   .agg(Mentions=("Row_ID", "nunique"),
                        Members=("Member", "nunique"),
                        Apprehension=("is_app", "sum"),
                        Defence=("is_def", "sum"),
                        Bodies=("Forum", lambda s: ", ".join(sorted(set(s)))),
                        First=("Date", "min"),
                        Last=("Date", "max"))
                   .reset_index()
                   .sort_values(["Mentions", "Measure"], ascending=[False, True]))
    owner_map = (M.groupby("Measure")["Owner"]
                 .agg(lambda s: s.dropna().mode().iloc[0] if s.dropna().size else ""))
    measure_ref.insert(2, "Owner", measure_ref["Measure"].map(owner_map).fillna(""))
    measure_ref["First"] = pd.to_datetime(measure_ref["First"]).dt.strftime("%d %b %Y")
    measure_ref["Last"] = pd.to_datetime(measure_ref["Last"]).dt.strftime("%d %b %Y")

    member_ref = (filtered.assign(one=1)
                  .pivot_table(index="Member", columns="Stance", values="one",
                               aggfunc="sum", fill_value=0)
                  .reindex(columns=STANCE_ORDER, fill_value=0)
                  .assign(Total=lambda x: x.sum(axis=1))
                  .sort_values("Total", ascending=False)
                  .reset_index())
    member_ref["Bodies"] = member_ref["Member"].map(
        filtered.groupby("Member")["Forum"].agg(lambda s: ", ".join(sorted(set(s)))))
    member_ref["Documents"] = member_ref["Member"].map(
        filtered.groupby("Member")["Document_Symbol"].nunique())

    which = st.radio("Reference table", ["Measures", "Members"], horizontal=True, key="ref_pick")
    table = measure_ref if which == "Measures" else member_ref
    st.dataframe(table, width="stretch", hide_index=True, height=340)
    note("Both tables reflect the current filters.")

    d1, d2, d3, d4 = st.columns(4)
    d1.download_button("Measure summary (CSV)", measure_ref.to_csv(index=False),
                       "measure_summary.csv", "text/csv", width="stretch", key="dl_meas")
    d2.download_button("Member summary (CSV)", member_ref.to_csv(index=False),
                       "member_summary.csv", "text/csv", width="stretch", key="dl_mem")
    d3.download_button("Filtered rows (CSV)", filtered.drop(columns=["Rule_Codes"]).to_csv(index=False),
                       "filtered_interactions.csv", "text/csv", width="stretch", key="dl_rows")
    if not vocab.empty:
        d4.download_button("Data dictionary (CSV)", vocab.to_csv(index=False),
                           "data_dictionary.csv", "text/csv", width="stretch", key="dl_dict")

    st.markdown("---")

    # --- coding quality ----------------------------------------------------------------
    st.markdown("#### Coding quality of the rows in view")
    q1, q2, q3 = st.columns(3)
    if "Confidence" in filtered.columns:
        conf = vc(filtered["Confidence"])
        fig = px.pie(conf, names="label", values="count", hole=.55, title="Coder confidence",
                     color="label",
                     color_discrete_map={"High": "#5C8A4A", "Medium": "#B0894A", "Low": "#C1662F"})
        fig.update_layout(height=290)
        show(fig, "q_conf", q1)
    if "Security_Relevance" in filtered.columns:
        show(hbar(vc(filtered["Security_Relevance"]), "Security relevance", xlabel="interactions",
                  height=290), "q_rel", q2)
    rules = filtered.explode("Rule_Codes")["Rule_Codes"].dropna()
    if len(rules):
        show(hbar(vc(rules), "Inclusion rules triggered", xlabel="interactions", height=290),
             "q_rule", q3)
        q3.caption("A row can satisfy several selection rules; R-codes follow Selection_Criteria.docx.")

    # --- vocabularies -------------------------------------------------------------------
    if not vocab.empty:
        with st.expander("Controlled vocabulary (permitted values by field)"):
            st.dataframe(vocab, width="stretch", hide_index=True)

    # --- measure grouping transparency ---------------------------------------------------
    with st.expander("How measures are grouped into families"):
        mapping = (measures_long(df)[["Measure", "Measure_Group"]]
                   .drop_duplicates().sort_values(["Measure_Group", "Measure"]))
        st.dataframe(mapping, width="stretch", hide_index=True, height=300)
        st.download_button("Measure family mapping (CSV)", mapping.to_csv(index=False),
                           "measure_family_mapping.csv", "text/csv", key="dl_map")
        st.caption(
            "Grouping runs in the app on keyword rules, applied in order, so the first matching "
            "family wins. To fix any assignment permanently, add a `Measure_Group` column to the "
            "Database sheet — the app will use it in preference to these rules."
        )

    # --- issues log ------------------------------------------------------------------------
    if not issues.empty:
        with st.expander(f"Issues log ({len(issues)} entries)"):
            sev = st.multiselect("Severity", sorted(issues["Severity"].dropna().unique()),
                                 key="iss_sev")
            view = issues[issues["Severity"].isin(sev)] if sev else issues
            st.dataframe(view, width="stretch", hide_index=True, height=340)

    # --- row explorer ---------------------------------------------------------------------
    st.markdown("#### Interaction records")
    search = st.text_input("Search summaries, measures and members", key="explore_q",
                           placeholder="e.g. rare earths, semiconductors, transparency")
    cols_show = [c for c in ["Date", "Document_Symbol", "Forum", "Member", "Stance",
                             "Measure 1", "Measure 2", "Owner", "Security_SubDomain_1",
                             "Interaction_Summary"] if c in filtered.columns]
    rows = filtered[cols_show].copy()
    if search:
        mask = rows.apply(lambda r: search.lower() in " ".join(map(str, r.values)).lower(), axis=1)
        rows = rows[mask]
    rows = rows.sort_values("Date", ascending=False)
    rows["Date"] = pd.to_datetime(rows["Date"]).dt.strftime("%d %b %Y")
    st.caption(f"{len(rows)} record{'s' if len(rows) != 1 else ''} shown.")
    st.dataframe(rows, width="stretch", hide_index=True, height=420)

st.markdown(
    f"<div class='note' style='margin-top:28px;border-top:1px solid {RULE};padding-top:10px;'>"
    f"Source: {wb.name} · Domain in scope: {DOMAIN} · Summaries are generated from the rows in view, "
    "not from a language model.</div>",
    unsafe_allow_html=True,
)
