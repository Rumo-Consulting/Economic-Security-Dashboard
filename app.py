"""
Economic Security Dashboard — Trade Law Observatory.

National & economic security in the WTO.

Design rules held throughout:
  * One tab answers one question. One chart answers one sub-question, and its title IS
    that question in plain English.
  * Four filters on the left, all optional. Everything else lives behind "More filters".
  * No sliders or toggles inside the tabs — sensible defaults are chosen for the reader.
  * Stance is the colour language: concern ochre, defence teal, proposal olive,
    general statement slate. Same meaning in every chart.
  * Summaries are computed from the rows in view — no API key, no network call, and they
    cannot state a number the data does not contain.

THE WORKBOOK LEADS, THE CODE FOLLOWS
  * Sheet names and column headers are matched loosely. "Governance Topic 2",
    "Gov Topic 2" and "governance_topic_2" all fill the same role. COLMAP records exactly
    which sheet column filled which role.
  * Stance values, their order and their colours are read off the Vocabularies sheet.
    Add, rename or reorder a stance there and the whole dashboard follows.
  * Governance dimensions, and which dimension each topic belongs to, come from the
    Vocabularies sheet too.
  * Governance coding uses the paired columns Governance Dimension 1/2/3 +
    Governance Topic 1/2/3. The older single "Governance_Dimensions_Topics" column is
    still understood if the pairs are absent.

Files expected in the same folder as this script:
    WTO_Database.xlsx      the data
    TLO_Logo_web.png       the masthead logo (optional — the app runs without it)
"""

from __future__ import annotations

import base64
import re
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.io as pio
import streamlit as st

# ======================================================================================
# Config and theme
# ======================================================================================
st.set_page_config(page_title="Economic Security Dashboard", page_icon="🛡️", layout="wide")

TOP_N = 12                      # fixed everywhere, so no "how many to show" control is needed
CONTEST_FLOOR = 3               # minimum mentions before a measure enters the concern ranking

# The period the dataset was compiled over. This is the search window, which is wider than the
# first and last dates that happen to appear in the rows — edit it when the coverage extends.
PERIOD_FROM = "01 Jan 2026"
PERIOD_TO = "15 Aug 2026"
LAST_UPDATED = "03 Sep 2026"    # edit this line whenever the workbook is refreshed

INK = "#16232E"
PRIMARY = "#2C5F7C"
MUTED = "#5F6E78"
RULE = "#DBE2E6"
BAR = "#4A7E9B"
CONCERN = "#C1662F"
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


# ======================================================================================
# Matching workbook labels to the roles the dashboard needs
# --------------------------------------------------------------------------------------
# Nothing below compares a workbook label with "==" against a hard-coded string. Sheet
# names, column headers and controlled values are all matched loosely, so renaming a
# heading or a permitted value in the workbook does not silently empty a chart.
# ======================================================================================
def norm(text) -> str:
    """Lower-case, strip punctuation and underscores, collapse spaces."""
    return re.sub(r"[^a-z0-9]+", " ", str(text).lower()).strip()


def tokens(text) -> set[str]:
    return set(norm(text).split())


SHEET_ALIASES: dict[str, list[str]] = {
    "Database": ["database", "data", "records", "interventions", "main", "coding"],
    "Vocabularies": ["vocabularies", "vocabulary", "vocab", "dictionary", "data dictionary",
                     "terms", "codebook"],
    "Issues_Log": ["issues log", "issues", "issue log", "known issues", "log", "review log"],
}

# canonical role | other spellings | words every candidate header must contain | essential?
COLUMN_SPEC: list[tuple[str, list[str], set[str], bool]] = [
    ("Document_Symbol", ["document", "doc symbol", "symbol", "document ref", "source document"],
     {"document"}, True),
    ("Date", ["meeting date", "date of meeting"], {"date"}, True),
    ("Title", ["document title"], {"title"}, False),
    ("WTO_Forum", ["forum", "wto body", "body", "committee", "council", "venue"], {"forum"}, True),
    ("Participant", ["member", "speaker", "delegation", "member speaking", "intervening member"],
     {"participant"}, True),
    ("Agenda_Item", ["agenda", "item", "agenda number"], {"agenda"}, False),
    ("Reference_Paragraph", ["reference para", "paragraph", "para", "reference", "cite"],
     {"reference"}, False),
    ("Stance", ["engagement", "position", "type of intervention", "intervention type"],
     {"stance"}, True),
    ("Domain", ["scope", "policy domain"], {"domain"}, False),
    ("Measure 1", ["measure", "measure one", "primary measure", "measure a"], {"measure", "1"}, False),
    ("Measure 2", ["measure two", "secondary measure", "measure b"], {"measure", "2"}, False),
    ("Measure 3", ["measure three", "measure c"], {"measure", "3"}, False),
    ("Measure_Owner", ["owner", "measure owner", "owner of measure", "responding member"],
     {"measure", "owner"}, False),
    ("Security_SubDomain_1", ["security subdomain 1", "security sub domain 1", "sub domain 1",
                              "security area 1", "subdomain 1", "security domain 1"],
     {"security", "1"}, False),
    ("Security_SubDomain_2", ["security subdomain 2", "security sub domain 2", "sub domain 2",
                              "security area 2", "subdomain 2", "security domain 2"],
     {"security", "2"}, False),
    ("Governance Dimension 1", ["gov dimension 1", "governance dimension one", "dimension 1"],
     {"dimension", "1"}, False),
    ("Governance Dimension 2", ["gov dimension 2", "governance dimension two", "dimension 2"],
     {"dimension", "2"}, False),
    ("Governance Dimension 3", ["gov dimension 3", "governance dimension three", "dimension 3"],
     {"dimension", "3"}, False),
    ("Governance Topic 1", ["gov topic 1", "governance topic one", "topic 1"], {"topic", "1"}, False),
    ("Governance Topic 2", ["gov topic 2", "governance topic two", "topic 2"], {"topic", "2"}, False),
    ("Governance Topic 3", ["gov topic 3", "governance topic three", "topic 3"], {"topic", "3"}, False),
    ("Governance_Dimensions_Topics", ["governance dimensions topics", "governance topics",
                                      "governance dimension topics"], set(), False),
    ("Interaction_Summary", ["summary", "what was said", "interaction", "intervention summary"],
     {"summary"}, False),
    ("Confidence", ["coder confidence", "confidence level"], {"confidence"}, False),
    ("Security_Relevance", ["relevance", "security relevance", "centrality"], {"relevance"}, False),
    ("Inclusion_Rule", ["inclusion", "rule"], {"inclusion"}, False),
    ("Review_Status", ["review status", "status"], {"review", "status"}, False),
    ("Review_Notes", ["review notes", "notes", "coder notes"], {"review", "notes"}, False),
]

VOCAB_FIELD_ALIASES: dict[str, list[str]] = {
    "Stance": ["stance", "engagement", "position"],
    "Security_Relevance": ["security relevance", "relevance"],
    "Confidence": ["confidence"],
    "Dimension": ["governance dimension", "dimension"],
    "Topic": ["governance topic", "governance dimensions topics", "topic"],
    "Security_SubDomain": ["security subdomain", "security sub domain", "subdomain", "security area"],
}

# How a stance NAME is read, whatever it is called. First rule that matches wins, so the
# order of this list is the order of precedence.
STANCE_ROLE_RULES: list[tuple[str, list[str]]] = [
    ("concern", ["apprehension", "objection", "criticism", "concern", "complaint", "opposition"]),
    ("defence", ["defence", "defense", "explanation", "justification", "rebuttal", "reply",
                 "response", "support"]),
    ("proposal", ["proposal", "recommendation", "suggestion", "request"]),
    ("general", ["general", "statement", "information", "sharing", "factual", "update", "note"]),
]
# Shades within a role, so two kinds of concern both stay recognisably ochre.
ROLE_SHADES: dict[str, list[str]] = {
    "concern": ["#C1662F", "#9C4F24"],
    "defence": ["#2F7E8C", "#245F6A"],
    "proposal": ["#5C8A4A", "#456A37"],
    "general": ["#8B99A6", "#AEB8C2"],
}
ROLE_PLAIN: dict[str, str] = {
    "concern": "raising a concern",
    "defence": "defending or explaining a measure",
    "proposal": "proposing something",
    "general": "making a general statement",
}
SPARE_COLORS = ["#8B6BA8", "#B0894A", "#4A7E9B", "#7A8B57", "#A8646E"]

# Fallbacks, used only where the Vocabularies sheet is silent.
GOV_DIMENSIONS_DEFAULT = ["Legal", "Economic", "Development", "Institutional", "Political Economy"]
TOPIC_PARENT_DEFAULT: dict[str, str] = {
    "WTO Consistency & Legal Interpretation": "Legal",
    "Transparency & Due Process": "Legal",
    "Market Access & Trade Effects": "Economic",
    "Competitiveness": "Economic",
    "Supply Chain & Economic Resilience": "Economic",
    "Supply Chain Impact": "Economic",
    "Policy Space": "Development",
    "Developmental & Distributional Effects": "Development",
    "International Cooperation": "Institutional",
    "Cooperation & Coordination": "Institutional",
    "Institutional Capacity & Coordination": "Institutional",
    "Economic Security & Strategic Autonomy": "Political Economy",
}

# Spelling repairs applied before anything is counted. Left side = what is in the
# workbook, right side = what to count it as. Empty, because the corrections now live in
# the workbook itself, which is where they belong.
TOPIC_FIXES: dict[str, str] = {
    # "Old label as typed": "Label to count it as",
}

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
      .block-container {{ padding-top: 2.4rem; padding-bottom: 3rem; max-width: 1280px; }}

      .masthead {{ border-bottom: 2px solid {INK}; padding: 6px 0 14px 0; margin-bottom: 4px; }}
      /* The logo is sized in viewport units so it stays legible on a phone and never
         outgrows the headline on a desktop screen. */
      .masthead .logo {{ display:block; width:100%; max-width:min(400px, 66vw); height:auto;
                         margin:0 0 16px 0; }}
      .masthead .eyebrow {{ font-family:'IBM Plex Mono', monospace; font-size:.7rem; font-weight:600;
                            letter-spacing:.16em; text-transform:uppercase; color:{PRIMARY};
                            line-height:1.9; }}
      .masthead h1 {{ font-size: clamp(1.7rem, 3.4vw, 2.3rem); margin:.1rem 0 .3rem 0;
                      font-weight:700; line-height:1.25; }}
      .masthead .sub {{ color:{MUTED}; font-size:1rem; max-width:none; line-height:1.5; }}
      .masthead .period {{ font-family:'IBM Plex Mono', monospace; font-size:.78rem; color:{MUTED};
                           margin-top:8px; }}
      .masthead .updated {{ font-family:'IBM Plex Mono', monospace; font-size:.78rem; color:{MUTED};
                            margin-top:2px; }}

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
      footer {{ display:none !important; }}

      @media (max-width: 640px) {{
          .block-container {{ padding-left:.7rem; padding-right:.7rem; }}
          div[data-testid="stMetricValue"] {{ font-size:1.2rem; }}
          .masthead .logo {{ max-width:80vw; margin-bottom:12px; }}
          .masthead {{ padding-top:2px; }}
      }}
    </style>
    """,
    unsafe_allow_html=True,
)


# ======================================================================================
# Files
# ======================================================================================
def find_workbook() -> Path | None:
    here = Path(__file__).parent
    for name in ["WTO_Database.xlsx", "Database_v21.xlsx", "Database.xlsx"]:
        if (here / name).exists():
            return here / name
    matches = sorted(p for p in here.glob("*.xlsx") if not p.name.startswith("~$"))
    return matches[-1] if matches else None


def find_logo() -> Path | None:
    here = Path(__file__).parent
    for name in ["TLO_Logo_web.png", "TLO_Logo.png", "logo.png"]:
        if (here / name).exists():
            return here / name
    return None


@st.cache_data(show_spinner=False)
def logo_uri(path_str: str, mtime: float) -> str:
    data = base64.b64encode(Path(path_str).read_bytes()).decode()
    return f"data:image/png;base64,{data}"


# ======================================================================================
# Loading, with loose matching of sheets and headers
# ======================================================================================
def resolve_sheet(sheet_names: list[str], role: str) -> str | None:
    aliases = {norm(a) for a in SHEET_ALIASES.get(role, [])} | {norm(role)}
    for name in sheet_names:                       # exact, once punctuation is ignored
        if norm(name) in aliases:
            return name
    for name in sheet_names:                       # loose, so "Database v22" still lands
        n = norm(name)
        if any(a and (n.startswith(a) or a in n) for a in aliases):
            return name
    return None


def resolve_columns(headers: list[str]) -> tuple[dict[str, str], list[str]]:
    """Work out which sheet header plays which role.

    Returns {header: role} plus any essential role that could not be filled.
    """
    mapping: dict[str, str] = {}
    taken: set[str] = set()
    filled: set[str] = set()

    def claim(header: str, role: str):
        mapping[header] = role
        taken.add(header)
        filled.add(role)

    # Pass 1 — the header IS the role name, give or take case and punctuation.
    for role, _aliases, _req, _ess in COLUMN_SPEC:
        if role in filled:
            continue
        for h in headers:
            if h not in taken and norm(h) == norm(role):
                claim(h, role)
                break

    # Pass 2 — the header is a known alternative spelling.
    for role, aliases, _req, _ess in COLUMN_SPEC:
        if role in filled:
            continue
        wanted = {norm(a) for a in aliases}
        for h in headers:
            if h not in taken and norm(h) in wanted:
                claim(h, role)
                break

    # Pass 3 — the header simply contains the words that identify the role, so
    # "Gov Topic 2" and "governance_topic_2" both land on Governance Topic 2.
    for role, _aliases, required, _ess in COLUMN_SPEC:
        if role in filled or not required:
            continue
        for h in headers:
            if h not in taken and required <= tokens(h):
                claim(h, role)
                break

    missing = [role for role, _a, _r, essential in COLUMN_SPEC if essential and role not in filled]
    return mapping, missing


@st.cache_data(show_spinner=False)
def load_data(path_str: str, mtime: float):
    xl = pd.ExcelFile(path_str)
    names = list(xl.sheet_names)

    db_sheet = resolve_sheet(names, "Database") or names[0]
    voc_sheet = resolve_sheet(names, "Vocabularies")
    iss_sheet = resolve_sheet(names, "Issues_Log")
    sheets = {"Database": db_sheet, "Vocabularies": voc_sheet, "Issues_Log": iss_sheet}

    df = pd.read_excel(xl, sheet_name=db_sheet)
    df.columns = [str(c).strip() for c in df.columns]
    vocab = pd.read_excel(xl, sheet_name=voc_sheet) if voc_sheet else pd.DataFrame()
    issues = pd.read_excel(xl, sheet_name=iss_sheet) if iss_sheet else pd.DataFrame()
    if not vocab.empty:
        vocab.columns = [str(c).strip() for c in vocab.columns]

    colmap, missing = resolve_columns(list(df.columns))
    unused = [c for c in df.columns if c not in colmap]
    df = df.rename(columns=colmap)
    if missing:
        return df, vocab, issues, colmap, unused, missing, sheets

    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df["Row_ID"] = range(len(df))
    df["Body"] = df["WTO_Forum"].map(FORUM_SHORT).fillna(df["WTO_Forum"])
    df["Member"] = df["Participant"].map(MEMBER_SHORT).fillna(df["Participant"])
    if "Measure_Owner" in df.columns:
        df["Owner"] = df["Measure_Owner"].map(MEMBER_SHORT).fillna(df["Measure_Owner"])
    else:
        df["Owner"] = pd.NA
    for col in ["Stance", "Security_Relevance", "Confidence"]:
        if col in df.columns:
            df[col] = df[col].astype("string").str.strip()
    return df, vocab, issues, colmap, unused, missing, sheets


# ======================================================================================
# Reading the controlled vocabularies out of the workbook
# ======================================================================================
def vocab_values(vocab: pd.DataFrame, role: str) -> list[str]:
    """Permitted values for a field, in the order the sheet lists them."""
    if vocab.empty:
        return []
    fcol = next((c for c in vocab.columns if "field" in norm(c)), None)
    vcol = next((c for c in vocab.columns if "value" in norm(c) or "permitted" in norm(c)), None)
    if fcol is None or vcol is None:
        return []
    aliases = [norm(a) for a in VOCAB_FIELD_ALIASES.get(role, [norm(role)])]
    hit = vocab[vocab[fcol].astype(str).map(lambda f: any(a in norm(f) for a in aliases))]
    out, seen = [], set()
    for v in hit[vcol].dropna().astype(str):
        v = v.strip()
        if v and v not in seen:
            seen.add(v)
            out.append(v)
    return out


def vocab_pairs(vocab: pd.DataFrame) -> list[tuple[str, str]]:
    """The permitted "Dimension: Topic" pairs, in sheet order."""
    pairs = []
    for value in vocab_values(vocab, "Topic"):
        if ":" in value:
            dim, topic = value.split(":", 1)
            pairs.append((dim.strip(), topic.strip()))
    return pairs


def stance_role(name: str) -> str | None:
    n = norm(name)
    for role, words in STANCE_ROLE_RULES:
        if any(w in n for w in words):
            return role
    return None


def build_stance_config(vocab: pd.DataFrame, observed: pd.Series):
    """Stance order, colours and plain-English phrasing, taken from the workbook.

    The Vocabularies sheet sets the order. Anything appearing in the data but not listed
    there is appended rather than dropped, so a new stance still shows up.
    """
    listed = vocab_values(vocab, "Stance")
    seen = observed.dropna().astype(str).str.strip()
    present = set(seen)
    order = [v for v in listed if v in present] + [v for v in seen.value_counts().index
                                                   if v not in listed]
    if not order:
        order = listed or list(dict.fromkeys(seen))

    roles, colors, plain = {}, {}, {}
    role_count: dict[str, int] = {}
    spare = list(SPARE_COLORS)
    for name in order:
        role = stance_role(name)
        roles[name] = role
        if role:
            i = role_count.get(role, 0)
            role_count[role] = i + 1
            shades = ROLE_SHADES[role]
            colors[name] = shades[i] if i < len(shades) else (spare.pop(0) if spare else shades[-1])
            plain[name] = ROLE_PLAIN[role]
        else:
            colors[name] = spare.pop(0) if spare else "#8B99A6"
            plain[name] = f"making a {name.lower()} intervention"

    concerns = [n for n in order if roles[n] == "concern"]
    defences = [n for n in order if roles[n] == "defence"]
    # If nothing reads as a concern, fall back to the most common stance so the concern
    # charts show something rather than silently nothing.
    if not concerns and order:
        concerns = [order[0]]
    return order, colors, plain, roles, concerns, defences


def pick_value(values, patterns: list[str], default: str | None = None) -> str | None:
    """The first value that reads like the thing we are looking for."""
    for v in values:
        if any(re.search(p, norm(v)) for p in patterns):
            return v
    return default


# ======================================================================================
# Reshaping
# ======================================================================================
def group_measure(name: str) -> str:
    if not isinstance(name, str) or not name.strip():
        return "General / unspecified"
    low = name.lower()
    for group, patterns in MEASURE_GROUPS:
        if any(re.search(p, low) for p in patterns):
            return group
    return "Other measures"


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
    """One row per governance ground: Row_ID, Slot (1-3), Dimension, Topic, Full.

    Reads the paired columns Governance Dimension 1/2/3 + Governance Topic 1/2/3, and
    falls back to the older single column if the pairs are absent.
    """
    keep = ["Row_ID", "Member", "Body", "Stance"]
    empty = pd.DataFrame(columns=keep + ["Slot", "Dimension", "Topic", "Full"])

    paired = [i for i in (1, 2, 3)
              if f"Governance Topic {i}" in d.columns or f"Governance Dimension {i}" in d.columns]

    parts = []
    if paired:
        for i in paired:
            sub = d[keep].copy()
            sub["Slot"] = i
            dcol, tcol = f"Governance Dimension {i}", f"Governance Topic {i}"
            sub["Dimension"] = d[dcol].astype("string").str.strip() if dcol in d.columns else pd.NA
            sub["Topic"] = d[tcol].astype("string").str.strip() if tcol in d.columns else pd.NA
            parts.append(sub)
    elif "Governance_Dimensions_Topics" in d.columns:
        # Legacy layout: "Legal: Transparency & Due Process | Economic: Competitiveness"
        legacy = d[keep + ["Governance_Dimensions_Topics"]].copy()
        legacy["Full"] = legacy["Governance_Dimensions_Topics"].astype("string").str.split("|")
        legacy = legacy.explode("Full").dropna(subset=["Full"])
        legacy["Full"] = legacy["Full"].str.strip()
        legacy = legacy[legacy["Full"].str.len() > 0]
        split = legacy["Full"].str.split(":", n=1, expand=True)
        legacy["Dimension"] = split[0].str.strip()
        legacy["Topic"] = split[1].str.strip() if split.shape[1] > 1 else split[0].str.strip()
        legacy["Slot"] = legacy.groupby("Row_ID").cumcount() + 1
        parts.append(legacy[keep + ["Slot", "Dimension", "Topic"]])
    else:
        return empty

    out = pd.concat(parts, ignore_index=True)
    for col in ["Dimension", "Topic"]:
        out[col] = out[col].astype("string").str.strip().replace({"": pd.NA})
        if TOPIC_FIXES:
            out[col] = out[col].replace(TOPIC_FIXES)

    # A topic name typed into the Dimension cell, with the Topic cell left blank, is put
    # back where it belongs rather than silently dropped.
    stray = out["Topic"].isna() & out["Dimension"].isin(list(TOPIC_PARENT))
    out.loc[stray, "Topic"] = out.loc[stray, "Dimension"]
    out.loc[stray, "Dimension"] = pd.NA

    out = out.dropna(subset=["Topic"])

    # A blank or unrecognised dimension is filled from the topic's parent, so the
    # dimension chart never loses a ground it should have counted.
    parent = out["Topic"].map(TOPIC_PARENT)
    needs = out["Dimension"].isna() | ~out["Dimension"].isin(GOV_DIMENSIONS)
    out.loc[needs, "Dimension"] = parent[needs].fillna(out.loc[needs, "Dimension"])
    out["Dimension"] = out["Dimension"].fillna("Unclassified")

    out["Full"] = out["Dimension"].astype(str) + ": " + out["Topic"].astype(str)
    out = out.sort_values(["Row_ID", "Slot"]).drop_duplicates(subset=["Row_ID", "Full"])
    return out.reset_index(drop=True)


def governance_label(d: pd.DataFrame) -> pd.Series:
    """One readable string per record: the grounds in the order they were coded."""
    grounds = topics_long(d)
    if grounds.empty:
        return pd.Series("", index=d.index, dtype="object")
    joined_ = (grounds.sort_values(["Row_ID", "Slot"])
               .groupby("Row_ID")["Full"].agg(lambda s: "; ".join(s)))
    return d["Row_ID"].map(joined_).fillna("")


def combine_cols(d: pd.DataFrame, cols: list[str], sep: str = "; ") -> pd.Series:
    """Squash 'Measure 1/2/3' style columns into one readable cell."""
    present = [c for c in cols if c in d.columns]
    if not present:
        return pd.Series("", index=d.index, dtype="object")
    frame = d[present].astype("string")
    return frame.apply(
        lambda r: sep.join(v.strip() for v in r.dropna() if str(v).strip()), axis=1
    )


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


# The record table shown at the foot of every tab. Ten columns, in reading order:
# when, where, who, what measure, whose measure, which security area, on what grounds,
# and what was actually said.
RECORD_COLUMNS: list[tuple[str, str]] = [
    ("Date", "Date"),
    ("Document_Symbol", "Document"),
    ("Body", "Body"),
    ("Reference_Paragraph", "Reference para"),
    ("Member", "Participant"),
    ("_Measures", "Measure"),
    ("Owner", "Measure owner"),
    ("_Areas", "Security sub-domain"),
    ("_Governance", "Governance dimension & topics"),
    ("Interaction_Summary", "Interaction summary"),
]


def record_table(data: pd.DataFrame) -> pd.DataFrame:
    """The filtered rows, cut down to the ten columns a reader actually needs."""
    rows = data.copy()
    rows["_Measures"] = combine_cols(rows, ["Measure 1", "Measure 2", "Measure 3"])
    rows["_Areas"] = combine_cols(rows, ["Security_SubDomain_1", "Security_SubDomain_2"], sep=" · ")
    rows["_Governance"] = governance_label(rows)

    cols = [(src, label) for src, label in RECORD_COLUMNS if src in rows.columns]
    out = rows[[src for src, _ in cols]].copy()
    out.columns = [label for _, label in cols]
    if "Date" in out.columns:
        out = out.sort_values("Date", ascending=False)
        out["Date"] = pd.to_datetime(out["Date"]).dt.strftime("%d %b %Y")
    return out


def records_panel(key: str, data: pd.DataFrame):
    """The underlying rows, collapsed, at the foot of every tab."""
    with st.expander(f"See the {len(data)} records behind this view"):
        search = st.text_input("Search", key=f"q_{key}",
                               placeholder="Try: rare earths, semiconductors, transparency…")
        rows = record_table(data)
        if search:
            mask = rows.apply(lambda r: search.lower() in " ".join(map(str, r.values)).lower(), axis=1)
            rows = rows[mask]
            st.caption(f"{len(rows)} of {len(data)} records match “{search}”.")
        st.dataframe(rows, width="stretch", hide_index=True, height=380)
        st.download_button("Download these records (CSV)", rows.to_csv(index=False),
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
        color = STANCE_COLORS.get(s, "#8B99A6")
        bars.append(f"<div style='width:{n / total * 100:.3f}%;background:{color};'></div>")
        keys.append(f"<span><span class='dot' style='background:{color}'></span>"
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
    st.error("**Workbook not found.** Put `WTO_Database.xlsx` in the same folder as `app.py` and reload.")
    st.stop()

df, vocab, issues, COLMAP, UNUSED_COLS, MISSING_COLS, SHEETS = load_data(str(wb), wb.stat().st_mtime)

if MISSING_COLS:
    st.error(
        "**The workbook is missing columns the dashboard cannot do without: "
        + ", ".join(f"`{m}`" for m in MISSING_COLS) + ".**\n\n"
        "Headings are matched loosely, so a rename is normally fine — but these roles have to stay "
        "recognisable. Headings found in the sheet: " + ", ".join(f"`{c}`" for c in df.columns)
    )
    st.stop()

# --- controlled values, read from the workbook rather than hard-coded ------------------
STANCE_ORDER, STANCE_COLORS, STANCE_PLAIN, STANCE_ROLES, CONCERN_STANCES, DEFENCE_STANCES = \
    build_stance_config(vocab, df["Stance"])

GOV_DIMENSIONS = list(GOV_DIMENSIONS_DEFAULT)
TOPIC_PARENT = dict(TOPIC_PARENT_DEFAULT)
for _dim, _topic in vocab_pairs(vocab):            # the sheet wins wherever it speaks
    TOPIC_PARENT[_topic] = _dim
    if _dim not in GOV_DIMENSIONS:
        GOV_DIMENSIONS.append(_dim)
for _dim in vocab_values(vocab, "Dimension"):
    if ":" not in _dim and _dim not in GOV_DIMENSIONS:
        GOV_DIMENSIONS.append(_dim)

_measure_cols = [df[c].dropna().astype(str) for c in ["Measure 1", "Measure 2", "Measure 3"]
                 if c in df.columns]
_measure_values = (pd.unique(pd.concat(_measure_cols)) if _measure_cols else [])
UNSPECIFIED = pick_value(_measure_values, [r"^unspecified", r"^not specified", r"^none"],
                         "Unspecified Measure")
NOT_APPLICABLE = pick_value(df["Owner"].dropna().astype(str).unique(),
                            [r"^not applicable", r"^n a$", r"^na$", r"^none$"], "Not applicable")
CORE_VALUE = pick_value(vocab_values(vocab, "Security_Relevance")
                        or (list(df["Security_Relevance"].dropna().astype(str).unique())
                            if "Security_Relevance" in df.columns else []),
                        [r"core", r"central", r"primary"], "Core")
LEGAL_DIM = pick_value(GOV_DIMENSIONS, [r"legal", r"\blaw\b", r"juridical"])

CONF_COLORS: dict[str, str] = {}
for _v in (vocab_values(vocab, "Confidence")
           or (list(df["Confidence"].dropna().astype(str).unique())
               if "Confidence" in df.columns else [])):
    _n = norm(_v)
    CONF_COLORS[_v] = ("#5C8A4A" if "high" in _n else
                       "#B0894A" if ("medium" in _n or "moderate" in _n) else
                       CONCERN if "low" in _n else PRIMARY)

DOMAIN = df["Domain"].dropna().iloc[0] if "Domain" in df.columns and df["Domain"].notna().any() else "—"
COVER_FROM, COVER_TO = df["Date"].min(), df["Date"].max()
GOV_ALL = topics_long(df)

logo_path = find_logo()
LOGO_URI = logo_uri(str(logo_path), logo_path.stat().st_mtime) if logo_path else None
if LOGO_URI:
    # Also pins the logo in the app header, which is where it stays visible on a phone
    # once the page is scrolled. Delete these six lines if you want it in one place only.
    try:
        st.logo(str(logo_path), size="large")
    except TypeError:
        st.logo(str(logo_path))
    except Exception:
        pass

# ======================================================================================
# Filters — four in plain sight, the rest tucked away
# ======================================================================================
st.sidebar.header("Filters")
st.sidebar.caption("Leave a filter empty to include everything. Filters apply to every tab.")

FILTER_KEYS = ["f_body", "f_member", "f_stance", "f_area", "f_group", "f_owner",
               "f_dim", "f_topic", "f_core", "f_dates"]

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
    f_dim = st.multiselect("Governance dimension",
                           sorted(GOV_ALL["Dimension"].unique()) if not GOV_ALL.empty else [],
                           key="f_dim")
    f_topic = st.multiselect("Governance topic",
                             sorted(GOV_ALL["Topic"].unique()) if not GOV_ALL.empty else [],
                             key="f_topic",
                             help="Any of the three grounds coded for a record, not just the first.")
    core_only = st.toggle(f"Only records where security is the {str(CORE_VALUE).lower()} issue",
                          value=False, key="f_core")
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
    filtered = filtered[filtered["Security_Relevance"] == CORE_VALUE]
if f_area:
    ids = set(areas_long(df).loc[lambda x: x["Area"].isin(f_area), "Row_ID"])
    filtered = filtered[filtered["Row_ID"].isin(ids)]
if f_group:
    ids = set(measures_long(df).loc[lambda x: x["Measure_Group"].isin(f_group), "Row_ID"])
    filtered = filtered[filtered["Row_ID"].isin(ids)]
if f_dim and not GOV_ALL.empty:
    filtered = filtered[filtered["Row_ID"].isin(set(GOV_ALL.loc[GOV_ALL["Dimension"].isin(f_dim), "Row_ID"]))]
if f_topic and not GOV_ALL.empty:
    filtered = filtered[filtered["Row_ID"].isin(set(GOV_ALL.loc[GOV_ALL["Topic"].isin(f_topic), "Row_ID"]))]
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
logo_html = f'<img class="logo" src="{LOGO_URI}" alt="Trade Law Observatory">' if LOGO_URI else ""
st.markdown(
    f"""
    <div class="masthead">
      {logo_html}
      <div class="eyebrow">Trade Law Observatory</div>
      <h1>Economic Security Dashboard</h1>
      <div class="sub">How WTO members raise, defend and contest {DOMAIN.lower()} measures in the
      organisation's formal meetings.</div>
      <div class="period">Data period: {PERIOD_FROM} to {PERIOD_TO}</div>
      <div class="updated">Last updated: {LAST_UPDATED}</div>
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

tab_over, tab_mem, tab_meas, tab_args = st.tabs(
    ["Overview", "Members", "Measures", "Arguments"]
)

# ======================================================================================
# OVERVIEW — what is in the data
# ======================================================================================
with tab_over:
    with st.expander("New here? How to read this dashboard"):
        colour_key = ", ".join(f"**{s}** is {STANCE_PLAIN.get(s, s.lower())}"
                               for s in STANCE_ORDER[:4])
        st.markdown(
            f"""
- **One record = one intervention** — a single member speaking once on a single agenda item.
- **Each tab answers one question.** *Overview* — what is in the data. *Members* — who speaks.
  *Measures* — what is discussed. *Arguments* — the grounds they argue on.
- **The colours never change their meaning:** {colour_key}.
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
    note("These colours mean the same thing in every chart across the dashboard.")

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
    apprehensive = (filtered[filtered["Stance"].isin(CONCERN_STANCES)]["Member"]
                    .value_counts().reindex(counts.index).fillna(0))
    defending = (filtered[filtered["Stance"].isin(DEFENCE_STANCES)]["Member"]
                 .value_counts().reindex(counts.index).fillna(0))
    vocal = counts[counts >= 3]

    line = ""
    if len(vocal) and apprehensive.sum():
        critics = (apprehensive / counts).reindex(vocal.index).sort_values(ascending=False)
        line = (f"Among those speaking at least three times, <b>{critics.index[0]}</b> raises concerns "
                f"most consistently ({pcs(apprehensive[critics.index[0]], counts[critics.index[0]])} of "
                f"its interventions)")
        if defending.sum():
            defenders = (defending / counts).reindex(vocal.index).sort_values(ascending=False)
            line += (f", while <b>{defenders.index[0]}</b> spends the most time defending or explaining "
                     f"measures ({pcs(defending[defenders.index[0]], counts[defenders.index[0]])})")
        line += ". "

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

    pairs = filtered[filtered["Owner"].notna() & (filtered["Owner"] != NOT_APPLICABLE)]
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

        contested = (NAMED.assign(app=NAMED["Stance"].isin(CONCERN_STANCES))
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

        owner_rows = filtered[filtered["Owner"].notna() & (filtered["Owner"] != NOT_APPLICABLE)]
        if len(owner_rows) >= 5:
            show(stance_hbar(owner_rows.groupby(["Owner", "Stance"]).size().reset_index(name="count"),
                             "Owner", "Whose measures are being discussed?", "interventions"),
                 "meas_owner")

        st.markdown("#### Measure summary table")
        ref = (M.assign(app=M["Stance"].isin(CONCERN_STANCES),
                        dfd=M["Stance"].isin(DEFENCE_STANCES))
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
        avg_grounds = round(len(T) / T["Row_ID"].nunique(), 1)

        legal_line = ""
        if LEGAL_DIM and (T["Dimension"] == LEGAL_DIM).any():
            legal = T[T["Dimension"] == LEGAL_DIM]["Row_ID"].nunique()
            legal_line = (f"<b>{legal}</b> interventions ({pcs(legal, len(filtered))}) argue at least "
                          f"partly in <b>{LEGAL_DIM.lower()}</b> terms — whether a measure is "
                          "WTO-consistent, transparent or procedurally fair. ")

        readout(
            f"The most common ground of argument is <b>{t_counts.index[0]}</b> "
            f"({int(t_counts.iloc[0])} interventions). " + legal_line
            + f"Each intervention is coded on <b>{avg_grounds}</b> grounds on average, up to a "
              f"maximum of three.{small_n(filtered)}"
        )

        show(stance_hbar(T.groupby(["Topic", "Stance"]).size().reset_index(name="count"),
                         "Topic", "What grounds do members argue on?", "interventions"), "arg_topic")
        note("All three coded grounds count here. A member usually argues on more than one at once, "
             "so an intervention can appear in several bars.")

    records_panel("args", filtered)
