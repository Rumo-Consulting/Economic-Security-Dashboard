"""
Trade Governance Lab — National & Economic Security in the WTO.

Simplified build. Design rules held throughout:
  * One tab answers one question. One chart answers one sub-question, and its title IS
    that question in plain English.
  * Four filters on the left, all optional. Everything else lives behind "More filters".
  * No sliders or toggles inside the tabs — sensible defaults are chosen for the reader.
  * Stance is the colour language: Apprehension ochre, Defence teal, Proposal olive,
    General Statement slate. Same meaning in every chart.
  * Summaries are computed from the rows in view — no API key, no network call, and they
    cannot state a number the data does not contain.

Reads Database_v21.xlsx (sheets: Database, Vocabularies, Issues_Log).
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

UNSPECIFIED = "Unspecified Measure"
TOP_N = 12                      # fixed everywhere, so no "how many to show" control is needed
CONTEST_FLOOR = 3               # minimum mentions before a measure enters the concern ranking

# The period the dataset was compiled over. This is the search window, which is wider than the
# first and last dates that happen to appear in the rows — edit it when the coverage extends.
PERIOD_FROM = "01 Jan 2026"
PERIOD_TO = "15 Aug 2026"

INK = "#16232E"
PRIMARY = "#2C5F7C"
MUTED = "#5F6E78"
RULE = "#DBE2E6"
BAR = "#4A7E9B"
CONCERN = "#C1662F"

STANCE_ORDER = ["Apprehension", "Defence/Explanation", "Proposal/Recommendation", "General Statement"]
STANCE_COLORS = {
    "Apprehension": "#C1662F",
    "Defence/Explanation": "#2F7E8C",
    "Proposal/Recommendation": "#5C8A4A",
    "General Statement": "#8B99A6",
}
STANCE_PLAIN = {
    "Apprehension": "raising a concern",
    "Defence/Explanation": "defending or explaining a measure",
    "Proposal/Recommendation": "proposing something",
    "General Statement": "making a general statement",
}
HEAT_SCALE = ["#F5F7F8", "#CBDBE3", "#93B8C9", "#5A8FA8", "#2C5F7C"]
PCONF = {"displayModeBar": False, "responsive": True}

FORUM_SHORT = {
    "General Council": "General Council",
    "CTG": "Council for Trade in Goods",
    "CMA": "Committee on Market Access",
    "SCM Committee": "Subsidies Committee",
    "TBT Committee": "TBT Committee",
    "TRIMS Committee": "TRIMS Committee",
    "Council for TRIPS / WGTTT": "TRIPS Council",
}
MEMBER_SHORT = {
    "Bolivarian Republic of Venezuela": "Venezuela",
    "Russian Federation": "Russia",
    "Republic of Korea": "Korea, Rep. of",
    "Gambia (on behalf of the LDC Group)": "Gambia (LDC Group)",
    "Mozambique (on behalf of the African Group)": "Mozambique (African Group)",
}

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
    colorway=[PRIMARY, CONCERN, "#5C8A4A", "#8B6BA8", "#2F7E8C", "#B0894A"],
    font=dict(family="IBM Plex Sans, Segoe UI, system-ui, sans-serif", size=13, color="#3A464F"),
    margin=dict(l=10, r=18, t=56, b=16),
    title=dict(font=dict(family="Source Serif 4, Georgia, serif", size=17, color=INK), x=0, xanchor="left"),
    xaxis=dict(automargin=True, gridcolor="#EDF1F3", zerolinecolor="#EDF1F3"),
    yaxis=dict(automargin=True, gridcolor="#EDF1F3", zerolinecolor="#EDF1F3"),
    legend=dict(orientation="h", yanchor="top", y=-0.13, xanchor="center", x=0.5, title_text=""),
)
pio.templates.default = "tgl"

st.markdown(
    f"""
    <style>
      @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@500;600&family=IBM+Plex+Sans:wght@400;500;600&family=Source+Serif+4:opsz,wght@8..60,600;8..60,700&display=swap');

      html, body, .stApp, [data-testid="stAppViewContainer"], [data-testid="stSidebar"],
      .stApp p, .stApp li, .stApp label, .stMarkdown {{
          font-family: 'IBM Plex Sans', Segoe UI, system-ui, sans-serif;
      }}
      /* Streamlit draws expander arrows and dropdown chevrons as Material ligatures.
         Never let the body font touch them or they render as the literal word. */
      [data-testid="stIconMaterial"], .material-symbols-rounded, [class*="material-symbols"],
      [data-testid="stExpanderToggleIcon"] {{
          font-family: 'Material Symbols Rounded', 'Material Icons' !important;
      }}
      h1, h2, h3, h4 {{ font-family: 'Source Serif 4', Georgia, serif; color: {INK};
                        letter-spacing: -0.01em; }}
      .block-container {{ padding-top: 3.2rem; padding-bottom: 3rem; max-width: 1280px; }}

      .masthead {{ border-bottom: 2px solid {INK}; padding: 6px 0 14px 0; margin-bottom: 4px; }}
      .masthead .eyebrow {{ font-family:'IBM Plex Mono', monospace; font-size:.7rem; font-weight:600;
                            letter-spacing:.16em; text-transform:uppercase; color:{PRIMARY};
                            line-height:1.9; }}
      .masthead h1 {{ font-size: clamp(1.7rem, 3.4vw, 2.3rem); margin:.1rem 0 .3rem 0;
                      font-weight:700; line-height:1.25; }}
      .masthead .sub {{ color:{MUTED}; font-size:1rem; max-width:none; line-height:1.5; }}
      .masthead .period {{ font-family:'IBM Plex Mono', monospace; font-size:.78rem; color:{MUTED};
                           margin-top:8px; }}

      .strip {{ display:flex; width:100%; height:14px; border-radius:7px; overflow:hidden;
                margin:18px 0 8px 0; border:1px solid {RULE}; }}
      .strip div {{ height:100%; }}
      .strip-key {{ font-size:.8rem; color:{MUTED}; display:flex; flex-wrap:wrap; gap:16px;
                    margin-bottom:6px; }}
      .strip-key b {{ font-weight:600; color:{INK}; }}
      .dot {{ display:inline-block; width:10px; height:10px; border-radius:2px; margin-right:6px; }}

      div[data-testid="stMetricValue"] {{ font-family:'IBM Plex Mono', monospace;
                                          color:{PRIMARY}; font-size:1.7rem; }}
      div[data-testid="stMetricLabel"] p {{ font-size:.78rem; letter-spacing:.04em;
                                            text-transform:uppercase; color:{MUTED}; }}

      .stTabs [data-baseweb="tab-list"] {{ gap:2px; flex-wrap:wrap; border-bottom:1px solid {RULE};
                                           margin-bottom:8px; }}
      .stTabs [data-baseweb="tab"] {{ font-weight:600; font-size:.95rem; padding:10px 20px;
                                      border-radius:8px 8px 0 0; color:{MUTED}; }}
      .stTabs [aria-selected="true"] {{ background:{PRIMARY}0F; color:{PRIMARY};
                                        border-bottom:3px solid {PRIMARY}; }}

      .read {{ background:#FFFFFF; border:1px solid {RULE}; border-left:4px solid {PRIMARY};
               border-radius:4px; padding:15px 19px; margin:6px 0 26px 0;
               font-size:.97rem; line-height:1.65; color:#2B3740; }}
      .read .tag {{ font-family:'IBM Plex Mono', monospace; font-size:.68rem; font-weight:600;
                    letter-spacing:.14em; text-transform:uppercase; color:{PRIMARY};
                    display:block; margin-bottom:8px; }}
      .read b {{ color:{INK}; font-weight:600; }}
      .note {{ font-size:.85rem; color:{MUTED}; margin:4px 0 30px 0; }}
      .lead {{ font-size:.92rem; color:{MUTED}; margin:2px 0 16px 0; }}

      section[data-testid="stSidebar"] {{ background:#F7F9FA; border-right:1px solid {RULE}; }}

      @media (max-width: 640px) {{
          .block-container {{ padding-left:.7rem; padding-right:.7rem; }}
          div[data-testid="stMetricValue"] {{ font-size:1.2rem; }}
      }}
    </style>
    """,
    unsafe_allow_html=True,
)


# ======================================================================================
# Data
# ======================================================================================
def find_workbook() -> Path | None:
    here = Path(__file__).parent
    for name in ["Database_v21.xlsx", "Database.xlsx", "WTO_Database.xlsx"]:
        if (here / name).exists():
            return here / name
    matches = sorted(here.glob("Database*.xlsx"))
    return matches[-1] if matches else None


def group_measure(name: str) -> str:
    if not isinstance(name, str) or not name.strip():
        return "General / unspecified"
    low = name.lower()
    for group, patterns in MEASURE_GROUPS:
        if any(re.search(p, low) for p in patterns):
            return group
    return "Other measures"


@st.cache_data(show_spinner=False)
def load_data(path_str: str, mtime: float):
    xl = pd.ExcelFile(path_str)
    df = pd.read_excel(xl, sheet_name="Database")
    df.columns = df.columns.str.strip()
    vocab = pd.read_excel(xl, sheet_name="Vocabularies") if "Vocabularies" in xl.sheet_names else pd.DataFrame()
    issues = pd.read_excel(xl, sheet_name="Issues_Log") if "Issues_Log" in xl.sheet_names else pd.DataFrame()

    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df["Row_ID"] = range(len(df))
    df["Body"] = df["WTO_Forum"].map(FORUM_SHORT).fillna(df["WTO_Forum"])
    df["Member"] = df["Participant"].map(MEMBER_SHORT).fillna(df["Participant"])
    df["Owner"] = df["Measure_Owner"].map(MEMBER_SHORT).fillna(df["Measure_Owner"])
    for col in ["Stance", "Security_Relevance", "Confidence"]:
        if col in df.columns:
            df[col] = df[col].astype("string").str.strip()
    return df, vocab, issues


def measures_long(d: pd.DataFrame) -> pd.DataFrame:
    keep = ["Row_ID", "Member", "Owner", "Body", "Stance", "Date", "Document_Symbol"]
    parts = []
    for col in ["Measure 1", "Measure 2", "Measure 3"]:
        if col in d.columns:
            sub = d[keep].copy()
            sub["Measure"] = d[col].astype("string").str.strip()
            parts.append(sub)
    if not parts:
        return pd.DataFrame(columns=keep + ["Measure", "Measure_Group"])
    out = pd.concat(parts, ignore_index=True).dropna(subset=["Measure"])
    out = out[out["Measure"] != ""]
    if "Measure_Group" in d.columns:
        lookup = d.set_index("Row_ID")["Measure_Group"]
        out["Measure_Group"] = out["Row_ID"].map(lookup).fillna(out["Measure"].map(group_measure))
    else:
        out["Measure_Group"] = out["Measure"].map(group_measure)
    return out.drop_duplicates(subset=["Row_ID", "Measure"]).reset_index(drop=True)


def areas_long(d: pd.DataFrame) -> pd.DataFrame:
    keep = ["Row_ID", "Member", "Body", "Stance"]
    parts = []
    for col in ["Security_SubDomain_1", "Security_SubDomain_2"]:
        if col in d.columns:
            sub = d[keep].copy()
            sub["Area"] = d[col].astype("string").str.strip()
            parts.append(sub)
    if not parts:
        return pd.DataFrame(columns=keep + ["Area"])
    out = pd.concat(parts, ignore_index=True).dropna(subset=["Area"])
    out = out[out["Area"] != ""].drop_duplicates(subset=["Row_ID", "Area"])
    return out.reset_index(drop=True)


def topics_long(d: pd.DataFrame) -> pd.DataFrame:
    col = "Governance_Dimensions_Topics"
    keep = ["Row_ID", "Member", "Body", "Stance"]
    if col not in d.columns:
        return pd.DataFrame(columns=keep + ["Dimension", "Topic"])
    sub = d[keep + [col]].copy()
    sub[col] = sub[col].astype("string")
    sub = sub.dropna(subset=[col])
    sub["Full"] = sub[col].str.split("|")
    out = sub.explode("Full")
    out["Full"] = out["Full"].str.strip()
    out = out[out["Full"].str.len() > 0]
    split = out["Full"].str.split(":", n=1, expand=True)
    out["Dimension"] = split[0].str.strip()
    out["Topic"] = split[1].str.strip() if split.shape[1] > 1 else split[0].str.strip()
    out = out.drop(columns=[col]).drop_duplicates(subset=["Row_ID", "Full"])
    return out.reset_index(drop=True)


# ======================================================================================
# Helpers
# ======================================================================================
def vc(series: pd.Series, top: int | None = None) -> pd.DataFrame:
    out = series.dropna().value_counts()
    if top:
        out = out.head(top)
    out = out.reset_index()
    out.columns = ["label", "count"]
    return out


def pct(part, whole) -> float:
    return 0.0 if not whole else round(part / whole * 100, 1)


def pcs(part, whole) -> str:
    return f"{pct(part, whole):g}%"


def show(fig, key, container=None):
    (container or st).plotly_chart(fig, width="stretch", config=PCONF, key=key)


def int_axis(fig, maxval):
    if not maxval or maxval <= 10:
        fig.update_xaxes(tickformat="d", dtick=1, rangemode="tozero")
    else:
        fig.update_xaxes(tickformat="d")
    return fig


def hbar(data: pd.DataFrame, title: str, unit: str, color=BAR, height=None):
    data = data.sort_values("count")
    h = height or max(250, 30 * len(data) + 130)
    fig = px.bar(data, x="count", y="label", orientation="h", title=title)
    fig.update_traces(marker_color=color, hovertemplate="%{y}<br>%{x} " + unit + "<extra></extra>")
    fig.update_layout(height=h, yaxis_title=None, xaxis_title=unit, showlegend=False,
                      margin=dict(t=56, b=56, l=10, r=18))
    return int_axis(fig, data["count"].max() if len(data) else 0)


def stance_hbar(data: pd.DataFrame, ycol: str, title: str, unit: str, height=None):
    order = data.groupby(ycol)["count"].sum().sort_values().index.tolist()
    h = height or max(280, 28 * len(order) + 170)
    fig = px.bar(data, x="count", y=ycol, color="Stance", orientation="h", title=title,
                 category_orders={ycol: order, "Stance": STANCE_ORDER},
                 color_discrete_map=STANCE_COLORS)
    # The legend sits under the plot, so the x-axis carries no title — it would collide.
    fig.update_layout(
        height=h, barmode="relative", yaxis_title=None, xaxis_title=None,
        margin=dict(t=56, b=96, l=10, r=18),
        legend=dict(orientation="h", yanchor="top", y=-0.16, xanchor="center", x=0.5,
                    title_text="", font=dict(size=12)),
    )
    fig.update_traces(hovertemplate="%{y}<br>%{x} " + unit + "<extra>%{fullData.name}</extra>")
    return int_axis(fig, data.groupby(ycol)["count"].sum().max() if len(data) else 0)


def readout(text: str):
    st.markdown(f"<div class='read'><span class='tag'>What this tab shows</span>{text}</div>",
                unsafe_allow_html=True)


def records_panel(key: str, data: pd.DataFrame):
    """The underlying rows, collapsed, at the foot of every tab."""
    with st.expander(f"See the {len(data)} records behind this view"):
        search = st.text_input("Search", key=f"q_{key}",
                               placeholder="Try: rare earths, semiconductors, transparency…")
        cols = [c for c in ["Date", "Document_Symbol", "Body", "Member", "Stance", "Measure 1",
                            "Owner", "Security_SubDomain_1", "Interaction_Summary"] if c in data.columns]
        rows = data[cols].copy()
        if search:
            mask = rows.apply(lambda r: search.lower() in " ".join(map(str, r.values)).lower(), axis=1)
            rows = rows[mask]
        rows = rows.sort_values("Date", ascending=False)
        rows["Date"] = pd.to_datetime(rows["Date"]).dt.strftime("%d %b %Y")
        rows = rows.rename(columns={"Document_Symbol": "Document", "Measure 1": "Measure",
                                    "Security_SubDomain_1": "Security area",
                                    "Interaction_Summary": "What was said"})
        if search:
            st.caption(f"{len(rows)} of {len(data)} records match “{search}”.")
        st.dataframe(rows, width="stretch", hide_index=True, height=380)
        st.download_button("Download these records (CSV)", data.to_csv(index=False),
                           "interactions.csv", "text/csv", key=f"dl_records_{key}")


def lead(text: str):
    st.markdown(f"<div class='lead'>{text}</div>", unsafe_allow_html=True)


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
                    f"{s} <b>{n}</b> ({pcs(n, total)})</span>")
    st.markdown(f"<div class='strip'>{''.join(bars)}</div>"
                f"<div class='strip-key'>{''.join(keys)}</div>", unsafe_allow_html=True)


def top_label(series: pd.Series):
    s = series.dropna()
    if s.empty:
        return "—", 0
    counts = s.value_counts()
    return str(counts.index[0]), int(counts.iloc[0])


def joined(items, limit=3) -> str:
    items = [str(i) for i in items][:limit]
    if not items:
        return "—"
    if len(items) == 1:
        return items[0]
    return ", ".join(items[:-1]) + " and " + items[-1]


def small_n(d: pd.DataFrame) -> str:
    return (" <b>Only a handful of records are in view</b>, so treat these shares as illustrative."
            if len(d) < 20 else "")


def concentration_note(d: pd.DataFrame) -> str:
    if d.empty or "Document_Symbol" not in d.columns:
        return ""
    doc, n = top_label(d["Document_Symbol"])
    share = pct(n, len(d))
    if share >= 30 and len(d) > 10:
        return (f" One meeting record, <b>{doc}</b>, supplies {share}% of these records, so the "
                "member rankings partly reflect who happened to attend it.")
    return ""


# ======================================================================================
# Load
# ======================================================================================
wb = find_workbook()
if wb is None:
    st.error("**Workbook not found.** Put `Database_v21.xlsx` in the same folder as `app.py` and reload.")
    st.stop()

df, vocab, issues = load_data(str(wb), wb.stat().st_mtime)
DOMAIN = df["Domain"].dropna().iloc[0] if "Domain" in df.columns and df["Domain"].notna().any() else "—"
COVER_FROM, COVER_TO = df["Date"].min(), df["Date"].max()

# ======================================================================================
# Filters — four in plain sight, the rest tucked away
# ======================================================================================
st.sidebar.header("Filters")
st.sidebar.caption("Leave a filter empty to include everything. Filters apply to every tab.")

FILTER_KEYS = ["f_body", "f_member", "f_stance", "f_area", "f_group", "f_owner", "f_core", "f_dates"]

f_body = st.sidebar.multiselect("WTO body", sorted(df["Body"].dropna().unique()), key="f_body")
f_member = st.sidebar.multiselect("Member speaking", sorted(df["Member"].dropna().unique()), key="f_member")
f_stance = st.sidebar.multiselect("How they engaged",
                                  [s for s in STANCE_ORDER if s in set(df["Stance"].dropna())],
                                  key="f_stance")
f_area = st.sidebar.multiselect("Security area", sorted(areas_long(df)["Area"].unique()), key="f_area")

with st.sidebar.expander("More filters"):
    f_group = st.multiselect("Measure family", sorted(measures_long(df)["Measure_Group"].unique()),
                             key="f_group")
    f_owner = st.multiselect("Measure owner", sorted(df["Owner"].dropna().unique()), key="f_owner",
                             help="The member whose measure is under discussion — not the speaker.")
    core_only = st.toggle("Only records where security is the core issue", value=False, key="f_core")
    picked = st.date_input("Date range", value=(COVER_FROM.date(), COVER_TO.date()),
                           min_value=COVER_FROM.date(), max_value=COVER_TO.date(), key="f_dates")
    if isinstance(picked, (list, tuple)):
        d_from = picked[0] if picked else COVER_FROM.date()
        d_to = picked[1] if len(picked) > 1 else COVER_TO.date()
    else:
        d_from = d_to = picked

filtered = df.copy()
if f_body:
    filtered = filtered[filtered["Body"].isin(f_body)]
if f_member:
    filtered = filtered[filtered["Member"].isin(f_member)]
if f_stance:
    filtered = filtered[filtered["Stance"].isin(f_stance)]
if f_owner:
    filtered = filtered[filtered["Owner"].isin(f_owner)]
if core_only and "Security_Relevance" in filtered.columns:
    filtered = filtered[filtered["Security_Relevance"] == "Core"]
if f_area:
    ids = set(areas_long(df).loc[lambda x: x["Area"].isin(f_area), "Row_ID"])
    filtered = filtered[filtered["Row_ID"].isin(ids)]
if f_group:
    ids = set(measures_long(df).loc[lambda x: x["Measure_Group"].isin(f_group), "Row_ID"])
    filtered = filtered[filtered["Row_ID"].isin(ids)]
filtered = filtered[(filtered["Date"] >= pd.Timestamp(d_from)) & (filtered["Date"] <= pd.Timestamp(d_to))]

st.sidebar.markdown("---")
st.sidebar.metric("Records in view", f"{len(filtered)} / {len(df)}")
if st.sidebar.button("Reset filters", width="stretch", key="reset"):
    # Deleting the widget keys (rather than clearing all of session_state) is what actually
    # returns a multiselect to empty — Streamlit re-creates it at its default on the rerun.
    for k in FILTER_KEYS:
        st.session_state.pop(k, None)
    st.rerun()

# ======================================================================================
# Masthead
# ======================================================================================
st.markdown(
    f"""
    <div class="masthead">
      <div class="eyebrow">WTO discussion analytics</div>
      <h1>Trade Governance Lab</h1>
      <div class="sub">How WTO members raise, defend and contest {DOMAIN.lower()} measures in the
      organisation's formal meetings.</div>
      <div class="period">Data period: {PERIOD_FROM} to {PERIOD_TO}</div>
    </div>
    """,
    unsafe_allow_html=True,
)

if filtered.empty:
    st.warning("Nothing matches these filters. Use **Reset filters** on the left.")
    st.stop()

M = measures_long(filtered)
A = areas_long(filtered)
T = topics_long(filtered)
NAMED = M[M["Measure"] != UNSPECIFIED]

tab_over, tab_mem, tab_meas, tab_args, tab_data = st.tabs(
    ["Overview", "Members", "Measures", "Arguments", "Data & method"]
)

# ======================================================================================
# OVERVIEW — what is in the data
# ======================================================================================
with tab_over:
    with st.expander("New here? How to read this dashboard"):
        st.markdown(
            """
- **One record = one intervention** — a single member speaking once on a single agenda item.
- **Each tab answers one question.** *Overview* — what is in the data. *Members* — who speaks.
  *Measures* — what is discussed. *Arguments* — the grounds they argue on. *Data & method* — how
  the dataset was built.
- **The four colours never change their meaning:** ochre is a concern being raised, teal is a
  member defending or explaining a measure, olive is a proposal, grey is a general statement.
- **Filters are on the left** and apply everywhere. Every tab opens with a summary written from
  whatever is currently in view, and closes with the records behind it.
            """
        )

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Interventions", len(filtered))
    c2.metric("Members", filtered["Member"].nunique())
    c3.metric("WTO bodies", filtered["Body"].nunique())
    c4.metric("Documents", filtered["Document_Symbol"].nunique())

    top_stance, n_stance = top_label(filtered["Stance"])
    top_area, n_area = top_label(A["Area"])
    top_body, n_body = top_label(filtered["Body"])
    readout(
        f"There are <b>{len(filtered)}</b> interventions here from <b>{filtered['Member'].nunique()}</b> "
        f"members. Most often a member is <b>{STANCE_PLAIN.get(top_stance, top_stance.lower())}</b> "
        f"({pcs(n_stance, len(filtered))} of interventions). The security area that comes up most is "
        f"<b>{top_area}</b>, and most of the talking happens in the <b>{top_body}</b> "
        f"({pcs(n_body, len(filtered))}).{concentration_note(filtered)}{small_n(filtered)}"
    )

    st.markdown("##### How members engaged")
    stance_strip(filtered)
    note("These four colours mean the same thing in every chart across the dashboard.")

    area_stance = A.groupby(["Area", "Stance"]).size().reset_index(name="count")
    show(stance_hbar(area_stance, "Area", "Which security areas come up most?", "interventions"),
         "ov_area")
    note("An intervention can touch two areas, so the bars add up to more than the record count.")

    show(hbar(vc(filtered["Body"]), "Where does the discussion happen?", "interventions"), "ov_body")

    records_panel("over", filtered)

# ======================================================================================
# MEMBERS — who speaks
# ======================================================================================
with tab_mem:
    lead("Who takes the floor, how they engage, and whose measures they are talking about.")

    counts = filtered["Member"].value_counts()
    apprehensive = (filtered[filtered["Stance"] == "Apprehension"]["Member"]
                    .value_counts().reindex(counts.index).fillna(0))
    defending = (filtered[filtered["Stance"] == "Defence/Explanation"]["Member"]
                 .value_counts().reindex(counts.index).fillna(0))
    vocal = counts[counts >= 3]

    line = ""
    if len(vocal):
        critics = (apprehensive / counts).reindex(vocal.index).sort_values(ascending=False)
        defenders = (defending / counts).reindex(vocal.index).sort_values(ascending=False)
        line = (f"Among those speaking at least three times, <b>{critics.index[0]}</b> raises concerns "
                f"most consistently ({pcs(apprehensive[critics.index[0]], counts[critics.index[0]])} of "
                f"its interventions), while <b>{defenders.index[0]}</b> spends the most time defending "
                f"or explaining measures "
                f"({pcs(defending[defenders.index[0]], counts[defenders.index[0]])}). ")

    readout(
        f"<b>{filtered['Member'].nunique()}</b> members speak in this view. "
        f"{joined(counts.head(3).index)} speak most often "
        f"({joined([str(i) for i in counts.head(3).values])} interventions respectively). "
        + line + concentration_note(filtered) + small_n(filtered)
    )

    keep = counts.head(TOP_N).index
    mem_stance = (filtered[filtered["Member"].isin(keep)]
                  .groupby(["Member", "Stance"]).size().reset_index(name="count"))
    show(stance_hbar(mem_stance, "Member", "Who speaks most, and how?", "interventions"), "mem_stance")

    pairs = filtered[filtered["Owner"].notna() & ~filtered["Owner"].isin(["Not applicable"])]
    cross = pairs[pairs["Member"] != pairs["Owner"]]
    if len(cross) >= 5:
        pc = (cross.groupby(["Member", "Owner"]).size().reset_index(name="count")
              .sort_values("count", ascending=False).head(TOP_N))
        pc["label"] = pc["Member"] + "  →  " + pc["Owner"]
        show(hbar(pc[["label", "count"]], "Who talks about whose measures?", "interventions"),
             "mem_pairs")
        self_ref = int((pairs["Member"] == pairs["Owner"]).sum())
        note(f"Read each bar as “the first member spoke about a measure belonging to the second”. "
             f"A further <b>{self_ref}</b> interventions are members talking about their own measures, "
             f"which are left out of this chart.")
    else:
        note("Too few records here name whose measure is being discussed to show the pairings.")

    records_panel("mem", filtered)

# ======================================================================================
# MEASURES — what is discussed
# ======================================================================================
with tab_meas:
    lead("The specific policy measures members raise, and how much heat each one draws.")

    if NAMED.empty:
        st.info("No named measures in this view. Widen the filters on the left.")
    else:
        m_counts = NAMED["Measure"].value_counts()
        g_counts = NAMED["Measure_Group"].value_counts()
        unspec = int((M["Measure"] == UNSPECIFIED).sum())

        contested = (NAMED.assign(app=NAMED["Stance"] == "Apprehension")
                     .groupby("Measure").agg(mentions=("Row_ID", "nunique"), concerns=("app", "sum")))
        contested = contested[contested["mentions"] >= CONTEST_FLOOR]
        contested["share"] = (contested["concerns"] / contested["mentions"] * 100).round(1)
        contested = contested.sort_values(["share", "mentions"], ascending=False)

        readout(
            f"<b>{NAMED['Measure'].nunique()}</b> named measures are discussed here, across "
            f"<b>{NAMED['Row_ID'].nunique()}</b> interventions. The one raised most often is "
            f"<b>{m_counts.index[0]}</b> ({int(m_counts.iloc[0])} interventions), and the broadest "
            f"theme is <b>{g_counts.index[0]}</b>. "
            + (f"Of the measures raised at least {CONTEST_FLOOR} times, <b>{contested.index[0]}</b> "
               f"attracts concern most consistently — {contested['share'].iloc[0]:g}% of the "
               f"interventions about it are concerns. " if len(contested) else "")
            + (f"A further <b>{unspec}</b> interventions in this view discuss security in general "
               "without naming a measure; they are not counted below." if unspec else "")
        )

        show(stance_hbar(NAMED[NAMED["Measure"].isin(m_counts.head(TOP_N).index)]
                         .groupby(["Measure", "Stance"]).size().reset_index(name="count"),
                         "Measure", "Which measures are raised most often?", "interventions"),
             "meas_top")

        if len(contested):
            cdf = contested.reset_index().head(TOP_N)
            cdf = cdf.rename(columns={"Measure": "label", "share": "count"})
            fig = hbar(cdf[["label", "count"]], "Which measures draw the most concern?", "% of mentions",
                       color=CONCERN)
            fig.update_xaxes(ticksuffix="%", dtick=25, tickformat=None, range=[0, 105])
            fig.update_traces(hovertemplate="%{y}<br>%{x}% of mentions are concerns<extra></extra>")
            show(fig, "meas_contested")
            note(f"Measures raised fewer than {CONTEST_FLOOR} times are left out, so a single "
                 "critical remark cannot top the chart.")

        show(hbar(vc(NAMED["Measure_Group"]), "What kinds of measure are these?", "interventions"),
             "meas_group")
        note("Families are grouped automatically from the measure names — the full mapping is on the "
             "<b>Records</b> tab.")

        owner_rows = filtered[filtered["Owner"].notna() & ~filtered["Owner"].isin(["Not applicable"])]
        if len(owner_rows) >= 5:
            show(stance_hbar(owner_rows.groupby(["Owner", "Stance"]).size().reset_index(name="count"),
                             "Owner", "Whose measures are being discussed?", "interventions"),
                 "meas_owner")

        st.markdown("#### Measure summary table")
        ref = (M.assign(app=M["Stance"] == "Apprehension",
                        dfd=M["Stance"] == "Defence/Explanation")
               .groupby(["Measure_Group", "Measure"])
               .agg(Interventions=("Row_ID", "nunique"), Members=("Member", "nunique"),
                    Concerns=("app", "sum"), Defences=("dfd", "sum"),
                    Bodies=("Body", lambda s: ", ".join(sorted(set(s)))),
                    First=("Date", "min"), Last=("Date", "max"))
               .reset_index().sort_values("Interventions", ascending=False))
        ref.insert(2, "Owner", ref["Measure"].map(
            M.groupby("Measure")["Owner"].agg(lambda s: s.dropna().mode().iloc[0] if s.dropna().size else "")))
        ref["First"] = pd.to_datetime(ref["First"]).dt.strftime("%d %b %Y")
        ref["Last"] = pd.to_datetime(ref["Last"]).dt.strftime("%d %b %Y")
        st.dataframe(ref, width="stretch", hide_index=True, height=300)
        st.download_button("Download this table (CSV)", ref.to_csv(index=False),
                           "measure_summary.csv", "text/csv", key="dl_meas")

    records_panel("meas", filtered)

# ======================================================================================
# ARGUMENTS — the grounds members argue on
# ======================================================================================
with tab_args:
    lead("Beyond the measure itself, what kind of argument does a member make?")

    if T.empty:
        st.info("No argument grounds recorded for these records.")
    else:
        t_counts = T["Topic"].value_counts()
        legal = T[T["Dimension"] == "Legal"]["Row_ID"].nunique()
        readout(
            f"The most common ground of argument is <b>{t_counts.index[0]}</b> "
            f"({int(t_counts.iloc[0])} interventions). "
            f"<b>{legal}</b> interventions ({pcs(legal, len(filtered))}) argue at least partly in legal "
            f"terms — whether a measure is WTO-consistent, transparent or procedurally "
            f"fair.{small_n(filtered)}"
        )

        show(stance_hbar(T.groupby(["Topic", "Stance"]).size().reset_index(name="count"),
                         "Topic", "What grounds do members argue on?", "interventions"), "arg_topic")
        note("A member usually argues on more than one ground at once, so an intervention can "
             "appear in several bars.")

        show(hbar(vc(T["Dimension"]), "Grouped into broad dimensions", "interventions", height=320),
             "arg_dim")

    records_panel("args", filtered)

# ======================================================================================
# DATA & METHOD — how the dataset was built
# ======================================================================================
with tab_data:
    lead("Where the records come from, what the terms mean, and how much to trust them.")

    if not vocab.empty:
        with st.expander("What the terms mean"):
            st.dataframe(vocab.rename(columns={"Permitted value": "Value"}),
                         width="stretch", hide_index=True)
            st.download_button("Download the data dictionary (CSV)", vocab.to_csv(index=False),
                               "data_dictionary.csv", "text/csv", key="dl_dict")

    with st.expander("How measures are grouped into families"):
        mapping = (measures_long(df)[["Measure", "Measure_Group"]]
                   .drop_duplicates().sort_values(["Measure_Group", "Measure"])
                   .rename(columns={"Measure_Group": "Family"}))
        st.dataframe(mapping, width="stretch", hide_index=True, height=280)
        st.download_button("Download the mapping (CSV)", mapping.to_csv(index=False),
                           "measure_families.csv", "text/csv", key="dl_map")
        st.caption("Grouping runs in the app on keyword rules. To fix an assignment permanently, add "
                   "a `Measure_Group` column to the Database sheet — the app will use it instead.")

    with st.expander("How confident is the coding?"):
        q1, q2 = st.columns(2)
        if "Confidence" in filtered.columns:
            conf = vc(filtered["Confidence"])
            fig = px.pie(conf, names="label", values="count", hole=.55, title="Coder confidence",
                         color="label",
                         color_discrete_map={"High": "#5C8A4A", "Medium": "#B0894A", "Low": CONCERN})
            fig.update_layout(height=300, margin=dict(t=56, b=20, l=10, r=10))
            show(fig, "q_conf", q1)
        if "Security_Relevance" in filtered.columns:
            core = int((filtered["Security_Relevance"] == "Core").sum())
            q2.metric("Security is the core issue", pcs(core, len(filtered)),
                      help="The rest touch on security as context rather than as the point of the "
                           "intervention. Filter to Core-only under More filters.")
            q2.caption(f"{core} of {len(filtered)} records in view.")

    if not issues.empty:
        with st.expander(f"Known issues in the dataset ({len(issues)})"):
            st.dataframe(issues, width="stretch", hide_index=True, height=300)

    records_panel("data", filtered)

st.markdown(
    f"<div class='note' style='margin-top:26px;border-top:1px solid {RULE};padding-top:10px;'>"
    f"Source: {wb.name} · Scope: {DOMAIN} · Summaries are written from the records in view, not by a "
    "language model.</div>",
    unsafe_allow_html=True,
)
