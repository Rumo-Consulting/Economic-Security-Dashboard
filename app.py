"""
Trade Governance Lab
WTO National & Economic Security Dashboard

Period covered: 1 Jan 2026 to 15 Aug 2026
Last updated: 17 Aug 2026

Data file expected in the same directory:
    WTO_Database(5).xlsx

Optional AI summaries:
    Set OPENAI_API_KEY in Streamlit secrets or environment.
    Optional OPENAI_MODEL defaults to gpt-5-mini.
    If no API key is available, the dashboard uses a deterministic
    dataset-derived summary. It never invents facts beyond the data.
"""

import io
import os
import html
from typing import Optional

import pandas as pd
import plotly.express as px
import plotly.io as pio
import streamlit as st


# =============================================================================
# PAGE / THEME
# =============================================================================

st.set_page_config(
    page_title="Trade Governance Lab | Economic Security",
    page_icon="🌐",
    layout="wide",
    initial_sidebar_state="expanded",
)

PRIMARY = "#245C73"
PRIMARY_DARK = "#173F50"
ACCENT = "#5F8FA3"
TEXT = "#263238"
MUTED = "#66757D"
GRID = "#E7ECEF"

PALETTE = [
    "#5F8FA3",
    "#D39B58",
    "#78A88A",
    "#9182B5",
    "#5FA59B",
    "#B97E9F",
    "#8499B0",
    "#C5A77D",
]

HEAT_SCALE = [
    "#F4F7F9",
    "#D7E3E8",
    "#AFC7D1",
    "#7FA5B4",
    "#4E7D91",
    "#2E5D72",
]

NO_MEASURE = "No Specific Measure"
NO_OWNER = "Not applicable"

SUMMARY_COLUMNS = [
    "Date",
    "WTO Body",
    "Document",
    "Member",
    "Agenda Item",
    "Reference Paragraph",
    "Stance",
    "Security Relevance",
    "Security Sub-domain",
    "Measure",
    "Measure Owner",
    "Governance Dimension",
    "Governance Topic",
    "Interaction Summary",
]

pio.templates["tgl"] = pio.templates["plotly_white"]
pio.templates["tgl"].layout.update(
    colorway=PALETTE,
    font=dict(
        family="Inter, Segoe UI, system-ui, sans-serif",
        size=13,
        color=TEXT,
    ),
    paper_bgcolor="white",
    plot_bgcolor="white",
    margin=dict(l=12, r=18, t=58, b=18),
    title=dict(
        font=dict(size=16, color=PRIMARY_DARK),
        x=0,
        xanchor="left",
    ),
    xaxis=dict(
        automargin=True,
        gridcolor=GRID,
        zeroline=False,
    ),
    yaxis=dict(
        automargin=True,
        gridcolor=GRID,
        zeroline=False,
    ),
    legend=dict(
        orientation="h",
        yanchor="top",
        y=-0.18,
        xanchor="center",
        x=0.5,
        title_text="",
    ),
)
pio.templates.default = "tgl"

st.markdown(
    f"""
    <style>
      .block-container {{
          max-width: 1500px;
          padding-top: 1.4rem;
          padding-bottom: 2.5rem;
      }}

      h1, h2, h3, h4 {{
          color: {PRIMARY_DARK};
      }}

      .app-title {{
          text-align: center;
          color: {PRIMARY};
          font-weight: 750;
          font-size: clamp(1.65rem, 3.3vw, 2.35rem);
          line-height: 1.25;
          margin: 0;
          padding: 4px 4px 2px 4px;
      }}

      .app-sub {{
          text-align: center;
          color: {MUTED};
          font-size: clamp(.92rem, 1.8vw, 1.06rem);
          margin: 2px auto 13px auto;
          max-width: 850px;
      }}

      .coverage-row {{
          display: flex;
          justify-content: center;
          gap: 10px;
          flex-wrap: wrap;
          margin: 0 auto 17px auto;
      }}

      .coverage-chip {{
          background: #F5F8F9;
          border: 1px solid #DCE6EA;
          border-radius: 9px;
          padding: 6px 12px;
          color: #4E5D63;
          font-size: .82rem;
      }}

      .ai-box {{
          background: linear-gradient(135deg, #F5F9FA, #FBFCFC);
          border: 1px solid #C9DCE3;
          border-left: 5px solid {PRIMARY};
          border-radius: 11px;
          padding: 13px 17px;
          margin: 5px 0 19px 0;
          font-size: .95rem;
          line-height: 1.55;
      }}

      .ai-tag {{
          display: inline-block;
          font-size: .68rem;
          font-weight: 750;
          letter-spacing: .07em;
          text-transform: uppercase;
          color: {PRIMARY};
          margin-bottom: 5px;
      }}

      .section-note {{
          color: {MUTED};
          font-size: .88rem;
          margin-top: -7px;
          margin-bottom: 8px;
      }}

      .source-heading {{
          color: {PRIMARY_DARK};
          font-weight: 700;
          font-size: 1.05rem;
          margin-top: 24px;
          margin-bottom: 2px;
      }}

      .source-note {{
          color: {MUTED};
          font-size: .84rem;
          margin-bottom: 8px;
      }}

      div[data-testid="stMetricValue"] {{
          color: {PRIMARY};
          font-size: 1.55rem;
      }}

      .stTabs [data-baseweb="tab-list"] {{
          gap: 3px;
          flex-wrap: wrap;
          justify-content: center;
      }}

      .stTabs [data-baseweb="tab"] {{
          font-weight: 650;
          padding: 8px 14px;
          border-radius: 9px 9px 0 0;
      }}

      .stTabs [aria-selected="true"] {{
          background: #245C7312;
          color: {PRIMARY};
          border-bottom: 3px solid {PRIMARY};
      }}

      @media (max-width: 700px) {{
          .block-container {{
              padding-left: .65rem;
              padding-right: .65rem;
          }}

          div[data-testid="stMetricValue"] {{
              font-size: 1.2rem;
          }}
      }}
    </style>
    """,
    unsafe_allow_html=True,
)


# =============================================================================
# DATA
# =============================================================================

@st.cache_data
def load_data():
    candidates = [
        "WTO_Database(5).xlsx",
        "WTO_Database.xlsx",
    ]
    path = next((p for p in candidates if os.path.exists(p)), None)

    if not path:
        st.error(
            "The dashboard could not find the Excel database. "
            "Place WTO_Database(5).xlsx in the same folder as app.py."
        )
        st.stop()

    data = pd.read_excel(path, sheet_name="Database")
    data.columns = data.columns.astype(str).str.strip()

    data["Date"] = pd.to_datetime(data["Date"], errors="coerce")
    data["WTO Body"] = data["WTO_Forum"].fillna("Not specified")
    data["Member"] = data["Participant"].fillna("Not specified")

    # Security sub-domains are stored across two columns.
    # Each interaction is counted once per distinct sub-domain.
    data["Security Sub-domain"] = data.apply(
        lambda r: distinct_join(
            [r.get("Security_SubDomain_1"), r.get("Security_SubDomain_2")]
        ),
        axis=1,
    )

    # Governance dimensions are stored across three columns.
    data["Governance Dimension"] = data.apply(
        lambda r: distinct_join(
            [
                r.get("Governance Dimension 1"),
                r.get("Governance Dimension 2"),
                r.get("Governance Dimension 3"),
            ]
        ),
        axis=1,
    )

    data["Governance Topic"] = data["Governance_Dimensions_Topics"].fillna("")

    # Measures are stored across three columns.
    data["Measure"] = data.apply(
        lambda r: distinct_join(
            [r.get("Measure 1"), r.get("Measure 2"), r.get("Measure 3")]
        ),
        axis=1,
    )

    # Display-only fields.
    data["Document"] = data["Document_Symbol"].fillna("")
    data["Agenda Item"] = data["Agenda_Item"].fillna("")
    data["Reference Paragraph"] = data["Reference_Paragraph"].fillna("")
    data["Stance"] = data["Stance"].fillna("Not specified")
    data["Security Relevance"] = data["Security_Relevance"].fillna("Not specified")
    data["Measure Owner"] = data["Measure_Owner"].fillna(NO_OWNER)
    data["Interaction Summary"] = data["Interaction_Summary"].fillna("")

    data["Month"] = data["Date"].dt.to_period("M").dt.to_timestamp()

    return data


def distinct_join(values):
    out = []
    for value in values:
        if pd.isna(value):
            continue
        value = str(value).strip()
        if not value or value.lower() in {"nan", "none"}:
            continue
        if value not in out:
            out.append(value)
    return " | ".join(out)


df = load_data()


# =============================================================================
# HELPERS
# =============================================================================

def count_unique_from_pipe(series):
    values = []
    for item in series.dropna():
        for part in str(item).split(" | "):
            part = part.strip()
            if part:
                values.append(part)
    return pd.Series(values).value_counts() if values else pd.Series(dtype="int64")


def explode_pipe(data, column, output_name):
    rows = []
    for _, row in data.iterrows():
        value = row.get(column, "")
        if pd.isna(value) or not str(value).strip():
            continue
        for part in str(value).split(" | "):
            part = part.strip()
            if part:
                rows.append(part)
    if not rows:
        return pd.DataFrame(columns=[output_name, "count"])
    return (
        pd.Series(rows, name=output_name)
        .value_counts()
        .rename_axis(output_name)
        .reset_index(name="count")
    )


def top_counts(series, top=12):
    s = series.dropna().astype(str)
    s = s[s.str.strip() != ""]
    out = s.value_counts().head(top).rename_axis("label").reset_index(name="count")
    return out


def pct(part, whole):
    return round((part / whole) * 100, 1) if whole else 0


def plural(n, singular, plural_form=None):
    if n == 1:
        return f"{n} {singular}"
    return f"{n} {plural_form or singular + 's'}"


def metric_strip(data):
    cols = st.columns(5)
    measures = data[data["Measure"] != NO_MEASURE]["Measure"].pipe(
        count_unique_from_pipe
    )

    vals = [
        ("Interactions", len(data)),
        ("Members", data["Member"].nunique()),
        ("WTO bodies", data["WTO Body"].nunique()),
        ("Named measures", len(measures)),
        ("Security themes", len(count_unique_from_pipe(data["Security Sub-domain"]))),
    ]

    for col, (label, value) in zip(cols, vals):
        col.metric(label, value)


def make_hbar(data, title, height=None, color=ACCENT, top=None):
    if data is None or data.empty:
        return None

    d = data.copy()
    if "count" not in d.columns:
        d.columns = ["label", "count"]

    d = d.sort_values("count", ascending=True)
    if top:
        d = d.tail(top)

    computed_height = height or max(270, 42 * len(d) + 95)

    fig = px.bar(
        d,
        x="count",
        y="label",
        orientation="h",
        title=title,
    )
    fig.update_traces(marker_color=color)
    fig.update_layout(
        height=computed_height,
        showlegend=False,
        xaxis_title=None,
        yaxis_title=None,
        margin=dict(l=12, r=18, t=58, b=18),
    )
    fig.update_xaxes(tickformat="d", dtick=1 if d["count"].max() <= 12 else None)
    return fig


def make_grouped_hbar(data, ycol, color_col, title, top=15, height=None):
    if data.empty:
        return None

    d = data.copy()
    totals = d.groupby(ycol)["count"].sum().sort_values()
    keep = totals.tail(top).index
    d = d[d[ycol].isin(keep)]

    computed_height = height or max(310, 40 * len(keep) + 120)

    fig = px.bar(
        d,
        x="count",
        y=ycol,
        color=color_col,
        orientation="h",
        barmode="stack",
        title=title,
        color_discrete_sequence=PALETTE,
    )
    fig.update_layout(
        height=computed_height,
        xaxis_title=None,
        yaxis_title=None,
        legend_title_text="",
        margin=dict(l=12, r=18, t=58, b=92),
        legend=dict(
            orientation="h",
            yanchor="top",
            y=-0.18,
            xanchor="center",
            x=0.5,
        ),
    )
    fig.update_xaxes(tickformat="d")
    return fig


def make_heatmap(matrix, title, height=390):
    if matrix is None or matrix.empty:
        return None

    fig = px.imshow(
        matrix,
        aspect="auto",
        text_auto=True,
        color_continuous_scale=HEAT_SCALE,
        title=title,
    )
    fig.update_layout(
        height=height,
        xaxis_title=None,
        yaxis_title=None,
        coloraxis_showscale=False,
        margin=dict(l=12, r=18, t=58, b=75),
    )
    fig.update_xaxes(tickangle=-20)
    return fig


def render_fig(fig, key):
    if fig is not None:
        st.plotly_chart(
            fig,
            use_container_width=True,
            config={"displayModeBar": False, "responsive": True},
            key=key,
        )


def deterministic_summary(data, context="current view"):
    if data.empty:
        return "No records match the current filters."

    n = len(data)
    members = data["Member"].nunique()
    bodies = data["WTO Body"].nunique()

    theme_counts = count_unique_from_pipe(data["Security Sub-domain"])
    stance_counts = data["Stance"].value_counts()
    relevance_counts = data["Security Relevance"].value_counts()

    text = (
        f"The {context} contains {n} interactions involving "
        f"{members} Members across {bodies} WTO bodies. "
    )

    if not theme_counts.empty:
        lead = theme_counts.index[0]
        text += (
            f"{lead} is the most frequently identified security sub-domain "
            f"({theme_counts.iloc[0]} tagged interactions). "
        )

    if not stance_counts.empty:
        lead = stance_counts.index[0]
        text += (
            f"{lead} is the most common stance "
            f"({stance_counts.iloc[0]} interactions, "
            f"{pct(stance_counts.iloc[0], n)}%). "
        )

    if not relevance_counts.empty:
        lead = relevance_counts.index[0]
        text += (
            f"{lead} accounts for {relevance_counts.iloc[0]} interactions "
            f"({pct(relevance_counts.iloc[0], n)}%)."
        )

    return text


def get_openai_key():
    try:
        if "OPENAI_API_KEY" in st.secrets:
            return st.secrets["OPENAI_API_KEY"]
    except Exception:
        pass
    return os.getenv("OPENAI_API_KEY")


def ai_summary(data, context="current view"):
    fallback = deterministic_summary(data, context)

    api_key = get_openai_key()
    if not api_key:
        label = "AI Summary · dataset-derived"
        text = fallback
    else:
        try:
            from openai import OpenAI

            model = os.getenv("OPENAI_MODEL", "gpt-5-mini")

            theme_counts = count_unique_from_pipe(data["Security Sub-domain"]).head(10)
            governance_counts = count_unique_from_pipe(data["Governance Dimension"]).head(10)
            stance_counts = data["Stance"].value_counts().head(10)
            body_counts = data["WTO Body"].value_counts().head(10)
            measure_counts = count_unique_from_pipe(
                data.loc[data["Measure"] != NO_MEASURE, "Measure"]
            ).head(10)

            payload = {
                "interactions": len(data),
                "members": int(data["Member"].nunique()),
                "wto_bodies": int(data["WTO Body"].nunique()),
                "security_relevance": data["Security Relevance"].value_counts().to_dict(),
                "security_subdomains": theme_counts.to_dict(),
                "stances": stance_counts.to_dict(),
                "governance_dimensions": governance_counts.to_dict(),
                "wto_bodies_breakdown": body_counts.to_dict(),
                "named_measures": measure_counts.to_dict(),
            }

            client = OpenAI(api_key=api_key)
            response = client.responses.create(
                model=model,
                input=(
                    "Write a concise, neutral research dashboard summary based ONLY on "
                    "the supplied dataset statistics. Do not infer causation, motives, "
                    "importance, or facts outside the statistics. Do not use superlatives "
                    "unless directly supported by counts. Write 2 to 4 sentences. "
                    "Mention the most important patterns visible in this filtered view. "
                    f"Context: {context}. Statistics: {payload}"
                ),
            )
            text = response.output_text.strip()
            label = "AI Summary"
        except Exception:
            text = fallback
            label = "AI Summary · dataset-derived"

    st.markdown(
        f"""
        <div class="ai-box">
          <div class="ai-tag">🤖 {html.escape(label)}</div><br>
          {html.escape(text)}
        </div>
        """,
        unsafe_allow_html=True,
    )


def source_table(data, key, title="Sources for this view"):
    st.markdown(f'<div class="source-heading">{title}</div>', unsafe_allow_html=True)
    st.markdown(
        f'<div class="source-note">{len(data)} records match the current filters. '
        "This table is generated from the same records used by the charts above.</div>",
        unsafe_allow_html=True,
    )

    table = prepare_summary_table(data)

    st.dataframe(
        table,
        use_container_width=True,
        hide_index=True,
        height=min(520, max(240, 45 + 36 * min(len(table), 12))),
    )

    c1, c2 = st.columns(2)

    csv_bytes = table.to_csv(index=False).encode("utf-8-sig")
    c1.download_button(
        "⬇️ Download reference table (CSV)",
        data=csv_bytes,
        file_name="trade_governance_lab_filtered_reference.csv",
        mime="text/csv",
        use_container_width=True,
        key=f"{key}_csv",
    )

    xlsx_buffer = io.BytesIO()
    with pd.ExcelWriter(xlsx_buffer, engine="openpyxl") as writer:
        table.to_excel(writer, index=False, sheet_name="Reference Table")
    xlsx_buffer.seek(0)

    c2.download_button(
        "⬇️ Download reference table (Excel)",
        data=xlsx_buffer.getvalue(),
        file_name="trade_governance_lab_filtered_reference.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
        use_container_width=True,
        key=f"{key}_xlsx",
    )


def prepare_summary_table(data):
    if data.empty:
        return pd.DataFrame(columns=SUMMARY_COLUMNS)

    rows = []

    for _, r in data.iterrows():
        # Use the first listed security theme, measure and governance dimension
        # for a compact one-row-per-interaction reference table.
        rows.append(
            {
                "Date": r["Date"].strftime("%d %b %Y") if pd.notna(r["Date"]) else "",
                "WTO Body": r["WTO Body"],
                "Document": r["Document"],
                "Member": r["Member"],
                "Agenda Item": r["Agenda Item"],
                "Reference Paragraph": r["Reference Paragraph"],
                "Stance": r["Stance"],
                "Security Relevance": r["Security Relevance"],
                "Security Sub-domain": r["Security Sub-domain"],
                "Measure": r["Measure"],
                "Measure Owner": r["Measure Owner"],
                "Governance Dimension": r["Governance Dimension"],
                "Governance Topic": r["Governance Topic"],
                "Interaction Summary": r["Interaction Summary"],
            }
        )

    return pd.DataFrame(rows, columns=SUMMARY_COLUMNS)


def filter_options(data, column):
    return sorted(
        [x for x in data[column].dropna().astype(str).unique() if x.strip()]
    )


# =============================================================================
# HEADER
# =============================================================================

st.markdown(
    '<div class="app-title">Trade Governance Lab</div>',
    unsafe_allow_html=True,
)
st.markdown(
    '<div class="app-sub">WTO discussions on national and economic security</div>',
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="coverage-row">
      <div class="coverage-chip"><b>Period covered:</b> 1 Jan 2026 to 15 Aug 2026</div>
      <div class="coverage-chip"><b>Last updated:</b> 17 Aug 2026</div>
    </div>
    """,
    unsafe_allow_html=True,
)


# =============================================================================
# GLOBAL FILTERS
# =============================================================================

st.sidebar.markdown("## 🔎 Filters")
st.sidebar.caption(
    "These filters apply across all tabs. Leave a filter empty to include everything."
)

min_date = df["Date"].min()
max_date = df["Date"].max()

date_value = st.sidebar.date_input(
    "Date range",
    value=(min_date.date(), max_date.date()),
    min_value=min_date.date(),
    max_value=pd.Timestamp("2026-08-15").date(),
)

if isinstance(date_value, tuple) and len(date_value) == 2:
    start_date, end_date = pd.Timestamp(date_value[0]), pd.Timestamp(date_value[1])
else:
    start_date = end_date = pd.Timestamp(date_value)

body_filter = st.sidebar.multiselect(
    "WTO Body",
    filter_options(df, "WTO Body"),
)

member_filter = st.sidebar.multiselect(
    "Member",
    filter_options(df, "Member"),
)

theme_filter = st.sidebar.multiselect(
    "Security sub-domain",
    [
        "Strategic Autonomy",
        "Supply Chains",
        "Sanctions",
        "Export Controls",
        "Security Tariffs",
    ],
)

stance_filter = st.sidebar.multiselect(
    "Stance",
    filter_options(df, "Stance"),
)

relevance_filter = st.sidebar.multiselect(
    "Security relevance",
    ["Core", "Contextual"],
)

with st.sidebar.expander("More filters"):
    owner_filter = st.multiselect(
        "Measure owner",
        filter_options(df, "Measure Owner"),
        key="global_owner_filter",
    )
    gov_filter = st.multiselect(
        "Governance dimension",
        [
            "Political Economy",
            "Legal",
            "Economic",
            "Development",
            "Institutional",
        ],
        key="global_gov_filter",
    )
    confidence_filter = st.multiselect(
        "Confidence",
        filter_options(df, "Confidence"),
        key="global_conf_filter",
    )

filtered = df[
    (df["Date"] >= start_date)
    & (df["Date"] <= end_date + pd.Timedelta(days=1) - pd.Timedelta(seconds=1))
].copy()

if body_filter:
    filtered = filtered[filtered["WTO Body"].isin(body_filter)]

if member_filter:
    filtered = filtered[filtered["Member"].isin(member_filter)]

if theme_filter:
    filtered = filtered[
        filtered["Security Sub-domain"].apply(
            lambda x: any(t in str(x).split(" | ") for t in theme_filter)
        )
    ]

if stance_filter:
    filtered = filtered[filtered["Stance"].isin(stance_filter)]

if relevance_filter:
    filtered = filtered[filtered["Security Relevance"].isin(relevance_filter)]

if owner_filter:
    filtered = filtered[filtered["Measure Owner"].isin(owner_filter)]

if gov_filter:
    filtered = filtered[
        filtered["Governance Dimension"].apply(
            lambda x: any(g in str(x).split(" | ") for g in gov_filter)
        )
    ]

if confidence_filter:
    filtered = filtered[filtered["Confidence"].isin(confidence_filter)]

st.sidebar.markdown("---")
st.sidebar.metric("Current view", f"{len(filtered)} / {len(df)} interactions")

if st.sidebar.button("Reset filters", use_container_width=True):
    st.rerun()

if filtered.empty:
    st.warning(
        "No records match the current filters. Adjust the filters or use "
        "**Reset filters** in the sidebar."
    )
    st.stop()


# =============================================================================
# TABS
# =============================================================================

tab_overview, tab_measures, tab_governance, tab_participation, tab_evidence = st.tabs(
    [
        "📊 Overview",
        "📑 Security Measures",
        "🏛️ Governance",
        "👥 Members & WTO Bodies",
        "🔎 Explore the Evidence",
    ]
)


# =============================================================================
# 1. OVERVIEW
# =============================================================================

with tab_overview:
    metric_strip(filtered)
    ai_summary(filtered, "the current Overview view")

    st.markdown("### Security discussions over time")
    monthly = (
        filtered.groupby("Month")
        .size()
        .reset_index(name="Interactions")
    )

    if not monthly.empty:
        fig = px.bar(
            monthly,
            x="Month",
            y="Interactions",
            title="Interactions by month",
        )
        fig.update_traces(marker_color=ACCENT)
        fig.update_layout(
            height=330,
            xaxis_title=None,
            yaxis_title=None,
            showlegend=False,
            margin=dict(l=12, r=18, t=58, b=25),
        )
        fig.update_xaxes(
            tickformat="%b %Y",
            dtick="M1",
        )
        fig.update_yaxes(tickformat="d")
        render_fig(fig, "overview_monthly")

    c1, c2 = st.columns(2)

    with c1:
        theme_counts = count_unique_from_pipe(
            filtered["Security Sub-domain"]
        ).head(10).rename_axis("label").reset_index(name="count")
        render_fig(
            make_hbar(
                theme_counts,
                "Security sub-domains",
                color=ACCENT,
            ),
            "overview_themes",
        )

    with c2:
        stance_counts = top_counts(filtered["Stance"], 8)
        render_fig(
            make_hbar(
                stance_counts,
                "How Members engage",
                color="#7D8FA3",
            ),
            "overview_stance",
        )

    relevance = filtered["Security Relevance"].value_counts().reset_index()
    relevance.columns = ["label", "count"]

    fig = px.bar(
        relevance,
        x="label",
        y="count",
        title="Core versus Contextual security relevance",
        color="label",
        color_discrete_sequence=[PRIMARY, "#B7C8CF"],
    )
    fig.update_layout(
        height=300,
        xaxis_title=None,
        yaxis_title=None,
        showlegend=False,
        margin=dict(l=12, r=18, t=58, b=18),
    )
    fig.update_yaxes(tickformat="d")
    render_fig(fig, "overview_relevance")

    source_table(filtered, "overview_sources")


# =============================================================================
# 2. SECURITY MEASURES
# =============================================================================

with tab_measures:
    named = filtered[filtered["Measure"] != NO_MEASURE].copy()

    metric_cols = st.columns(4)
    metric_cols[0].metric("Named measures", len(count_unique_from_pipe(named["Measure"])))
    metric_cols[1].metric("Measure owners", named["Measure Owner"].nunique())
    metric_cols[2].metric("Members engaged", named["Member"].nunique())
    metric_cols[3].metric("Interactions", len(named))

    ai_summary(filtered, "the Security Measures view")

    if named.empty:
        st.info("There are no named measures in the current filtered view.")
    else:
        c1, c2 = st.columns(2)

        with c1:
            measures = count_unique_from_pipe(named["Measure"]).head(12)
            measures = measures.rename_axis("label").reset_index(name="count")
            render_fig(
                make_hbar(
                    measures,
                    "Most discussed named measures",
                    color=ACCENT,
                ),
                "measures_top",
            )

        with c2:
            owners = named["Measure Owner"].value_counts().head(12)
            owners = owners.rename_axis("label").reset_index(name="count")
            owners = owners[owners["label"] != NO_OWNER]
            render_fig(
                make_hbar(
                    owners,
                    "Measures by owner",
                    color="#D39B58",
                ),
                "measures_owners",
            )

        measure_theme = []
        for _, row in named.iterrows():
            measures_for_row = str(row["Measure"]).split(" | ")
            themes_for_row = str(row["Security Sub-domain"]).split(" | ")
            for measure in measures_for_row:
                if measure and measure != NO_MEASURE:
                    for theme in themes_for_row:
                        if theme:
                            measure_theme.append((measure, theme))

        if measure_theme:
            mt = (
                pd.DataFrame(measure_theme, columns=["Measure", "Security theme"])
                .value_counts()
                .reset_index(name="count")
            )
            render_fig(
                make_grouped_hbar(
                    mt,
                    "Measure",
                    "Security theme",
                    "Measures by security sub-domain",
                    top=14,
                    height=460,
                ),
                "measures_theme",
            )

        st.markdown("### Focus on a measure")
        measure_options = sorted(count_unique_from_pipe(named["Measure"]).index.tolist())

        selected_measure = st.selectbox(
            "Select a named measure",
            ["All named measures"] + measure_options,
            key="measure_focus",
        )

        if selected_measure != "All named measures":
            mdata = named[
                named["Measure"].apply(
                    lambda x: selected_measure in str(x).split(" | ")
                )
            ]
        else:
            mdata = named

        mc = st.columns(4)
        mc[0].metric("Interactions", len(mdata))
        mc[1].metric("Members", mdata["Member"].nunique())
        mc[2].metric("WTO bodies", mdata["WTO Body"].nunique())
        owner_mode = mdata["Measure Owner"].mode()
        mc[3].metric(
            "Main owner",
            owner_mode.iloc[0] if not owner_mode.empty else "Not specified",
        )

        c3, c4 = st.columns(2)
        with c3:
            people = top_counts(mdata["Member"], 12)
            render_fig(make_hbar(people, "Members discussing the measure"), "measure_members")

        with c4:
            themes = count_unique_from_pipe(mdata["Security Sub-domain"])
            themes = themes.rename_axis("label").reset_index(name="count")
            render_fig(make_hbar(themes, "Security framing"), "measure_themes")

        source_table(
            mdata,
            "measures_sources",
            title="Sources for the selected measure view",
        )


# =============================================================================
# 3. GOVERNANCE
# =============================================================================

with tab_governance:
    ai_summary(filtered, "the Governance view")

    dims = count_unique_from_pipe(
        filtered["Governance Dimension"]
    ).rename_axis("label").reset_index(name="count")

    c1, c2 = st.columns(2)

    with c1:
        render_fig(
            make_hbar(
                dims,
                "Governance dimensions",
                color=ACCENT,
            ),
            "gov_dimensions",
        )

    with c2:
        # Governance topics are already stored as text. Extract topic labels
        # conservatively using the database's semicolon-separated structure.
        topic_rows = []
        for value in filtered["Governance Topic"].dropna():
            for topic in str(value).split(";"):
                topic = topic.strip()
                if topic:
                    topic_rows.append(topic)

        topics = (
            pd.Series(topic_rows)
            .value_counts()
            .head(12)
            .rename_axis("label")
            .reset_index(name="count")
            if topic_rows
            else pd.DataFrame(columns=["label", "count"])
        )

        render_fig(
            make_hbar(
                topics,
                "Governance topics",
                color="#7D8FA3",
            ),
            "gov_topics",
        )

    # Security sub-domain × governance dimension.
    pairs = []
    for _, row in filtered.iterrows():
        themes = str(row["Security Sub-domain"]).split(" | ")
        dimensions = str(row["Governance Dimension"]).split(" | ")
        for theme in themes:
            if theme:
                for dim in dimensions:
                    if dim:
                        pairs.append((theme, dim))

    if pairs:
        cross = pd.DataFrame(
            pairs, columns=["Security sub-domain", "Governance dimension"]
        )
        matrix = pd.crosstab(
            cross["Security sub-domain"],
            cross["Governance dimension"],
        )
        render_fig(
            make_heatmap(
                matrix,
                "Security sub-domain × governance dimension",
                height=390,
            ),
            "gov_heatmap",
        )

    # WTO body × governance dimension.
    body_dim = []
    for _, row in filtered.iterrows():
        bodies = [str(row["WTO Body"])]
        dimensions = str(row["Governance Dimension"]).split(" | ")
        for body in bodies:
            for dim in dimensions:
                if dim:
                    body_dim.append((body, dim))

    if body_dim:
        bd = pd.DataFrame(
            body_dim, columns=["WTO body", "Governance dimension"]
        )
        matrix2 = pd.crosstab(
            bd["WTO body"],
            bd["Governance dimension"],
        )
        render_fig(
            make_heatmap(
                matrix2,
                "WTO body × governance dimension",
                height=410,
            ),
            "gov_body_heatmap",
        )

    source_table(filtered, "governance_sources")


# =============================================================================
# 4. MEMBERS & WTO BODIES
# =============================================================================

with tab_participation:
    ai_summary(filtered, "the Members and WTO Bodies view")

    mode = st.radio(
        "Analyse",
        ["Members", "WTO Bodies"],
        horizontal=True,
        key="participation_mode",
    )

    if mode == "Members":
        members = top_counts(filtered["Member"], 15)
        render_fig(
            make_hbar(
                members,
                "Most active Members",
                color=ACCENT,
            ),
            "members_active",
        )

        # Top-member profile.
        member_options = filter_options(filtered, "Member")
        selected_member = st.selectbox(
            "Member profile",
            ["All Members"] + member_options,
            key="member_profile",
        )

        pdata = (
            filtered
            if selected_member == "All Members"
            else filtered[filtered["Member"] == selected_member]
        )

        pc = st.columns(4)
        pc[0].metric("Interactions", len(pdata))
        pc[1].metric("WTO bodies", pdata["WTO Body"].nunique())
        pc[2].metric("Security themes", len(count_unique_from_pipe(pdata["Security Sub-domain"])))
        pc[3].metric("Named measures", len(count_unique_from_pipe(
            pdata.loc[pdata["Measure"] != NO_MEASURE, "Measure"]
        )))

        c1, c2 = st.columns(2)
        with c1:
            themes = count_unique_from_pipe(pdata["Security Sub-domain"])
            themes = themes.rename_axis("label").reset_index(name="count")
            render_fig(make_hbar(themes, "Security themes engaged with"), "member_themes")

        with c2:
            stances = top_counts(pdata["Stance"], 8)
            render_fig(make_hbar(stances, "How the Member engages"), "member_stance")

        c3, c4 = st.columns(2)
        with c3:
            dims = count_unique_from_pipe(pdata["Governance Dimension"])
            dims = dims.rename_axis("label").reset_index(name="count")
            render_fig(make_hbar(dims, "Governance dimensions"), "member_dims")

        with c4:
            bodies = top_counts(pdata["WTO Body"], 10)
            render_fig(make_hbar(bodies, "WTO bodies where active"), "member_bodies")

        source_table(
            pdata,
            "member_sources",
            title="Sources for the selected Member view",
        )

    else:
        bodies = top_counts(filtered["WTO Body"], 12)
        render_fig(
            make_hbar(
                bodies,
                "Interactions by WTO body",
                color=ACCENT,
            ),
            "bodies_active",
        )

        body_options = filter_options(filtered, "WTO Body")
        selected_body = st.selectbox(
            "WTO body profile",
            ["All WTO bodies"] + body_options,
            key="body_profile",
        )

        bdata = (
            filtered
            if selected_body == "All WTO bodies"
            else filtered[filtered["WTO Body"] == selected_body]
        )

        bc = st.columns(4)
        bc[0].metric("Interactions", len(bdata))
        bc[1].metric("Members", bdata["Member"].nunique())
        bc[2].metric("Security themes", len(count_unique_from_pipe(bdata["Security Sub-domain"])))
        bc[3].metric("Named measures", len(count_unique_from_pipe(
            bdata.loc[bdata["Measure"] != NO_MEASURE, "Measure"]
        )))

        c1, c2 = st.columns(2)
        with c1:
            people = top_counts(bdata["Member"], 12)
            render_fig(make_hbar(people, "Most active Members"), "body_members")

        with c2:
            themes = count_unique_from_pipe(bdata["Security Sub-domain"])
            themes = themes.rename_axis("label").reset_index(name="count")
            render_fig(make_hbar(themes, "Security themes"), "body_themes")

        c3, c4 = st.columns(2)
        with c3:
            stances = top_counts(bdata["Stance"], 8)
            render_fig(make_hbar(stances, "Stances"), "body_stance")

        with c4:
            dims = count_unique_from_pipe(bdata["Governance Dimension"])
            dims = dims.rename_axis("label").reset_index(name="count")
            render_fig(make_hbar(dims, "Governance dimensions"), "body_dims")

        source_table(
            bdata,
            "body_sources",
            title="Sources for the selected WTO body view",
        )


# =============================================================================
# 5. EXPLORE THE EVIDENCE
# =============================================================================

with tab_evidence:
    ai_summary(filtered, "the Explore the Evidence view")

    st.markdown("### Search the current filtered records")
    search = st.text_input(
        "Search",
        placeholder="Search Member, document, measure, topic, summary, WTO body, etc.",
        key="evidence_search",
    )

    evidence = filtered.copy()

    if search:
        mask = evidence.astype(str).apply(
            lambda col: col.str.contains(search, case=False, na=False)
        )
        evidence = evidence[mask.any(axis=1)]

    st.caption(f"Showing {len(evidence)} matching records.")

    # Full researcher-facing table. Internal coding fields are retained here.
    evidence_cols = [
        "Date",
        "Document",
        "Title",
        "WTO Body",
        "Member",
        "Agenda Item",
        "Reference Paragraph",
        "Stance",
        "Domain",
        "Measure 1",
        "Measure 2",
        "Measure 3",
        "Measure Owner",
        "Security_SubDomain_1",
        "Security_SubDomain_2",
        "Governance Dimension 1",
        "Governance Dimension 2",
        "Governance Dimension 3",
        "Governance_Dimensions_Topics",
        "Interaction_Summary",
        "Confidence",
        "Security Relevance",
        "Inclusion_Rule",
        "Review_Status",
        "Review_Notes",
    ]

    evidence_cols = [c for c in evidence_cols if c in evidence.columns]

    st.dataframe(
        evidence[evidence_cols].sort_values("Date"),
        use_container_width=True,
        hide_index=True,
        height=600,
    )

    # Full filtered research dataset download.
    full_csv = evidence[evidence_cols].to_csv(index=False).encode("utf-8-sig")

    st.download_button(
        "⬇️ Download current filtered research dataset (CSV)",
        data=full_csv,
        file_name="trade_governance_lab_filtered_research_dataset.csv",
        mime="text/csv",
        use_container_width=True,
        key="evidence_full_csv",
    )

    source_table(
        evidence,
        "evidence_sources",
        title="Downloadable summary reference table",
    )


# =============================================================================
# FOOTER
# =============================================================================

st.markdown("---")
st.caption(
    "Trade Governance Lab. The dashboard presents coded WTO discussion records "
    "from the underlying research database. Counts and summaries update with the "
    "current filters."
)
