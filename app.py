"""
Team 15 – The Distressed Property Scouts
Streamlit MVP — Final version
Primary Task  : Flag counties where 2BR FMR is >= 25% below (state, metro) median
Phase 2A      : Neighbor-adjusted distress score (literal adjacent counties)
Phase 2B      : Recovery Signal Score (multi-year CAGR + trend consistency)
AI            : Separate County Brief tab with base/enhanced modes
"""

import json
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import requests
import streamlit as st
from google import genai
from google.genai import types
from scipy import stats

st.set_page_config(
    page_title="The Distressed Property Scouts",
    page_icon="🏚️",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
<style>
  .stApp { background-color: #0D0D1A; }
  [data-testid="stSidebar"] {
    background: linear-gradient(180deg, #1A0A2E 0%, #0D0D1A 100%);
  }
  [data-testid="stSidebar"] * { color: #E8E8F0 !important; }
  .kpi { background: #1A1A2E; border-radius: 10px; padding: 16px 20px;
         border-left: 4px solid #FF4500; margin-bottom: 10px; }
  .kpi-val { font-size: 2rem; font-weight: 800; color: #FF4500; }
  .kpi-lbl { font-size: 0.8rem; color: #9090A8; margin-top: 2px; }
  .sec-hdr { font-size: 1rem; font-weight: 700; color: #FF4500;
             text-transform: uppercase; letter-spacing: 1px;
             border-bottom: 1px solid #2A1A3E; padding-bottom: 5px;
             margin: 18px 0 12px 0; }
  p, li, span { color: #D8D8E8 !important; }
  h1, h2, h3 { color: #FF4500 !important; }
</style>
""",
    unsafe_allow_html=True,
)

ALL_STATES = [
    "AL", "AK", "AZ", "AR", "CA", "CO", "CT", "DE", "FL", "GA", "HI", "ID", "IL", "IN", "IA",
    "KS", "KY", "LA", "ME", "MD", "MA", "MI", "MN", "MS", "MO", "MT", "NE", "NV", "NH", "NJ",
    "NM", "NY", "NC", "ND", "OH", "OK", "OR", "PA", "RI", "SC", "SD", "TN", "TX", "UT", "VT",
    "VA", "WA", "WV", "WI", "WY", "DC",
]

DARK = dict(paper_bgcolor="#0D0D1A", plot_bgcolor="#14142A", font_color="#D8D8E8")
COUNTY_GEOJSON_URL = "https://raw.githubusercontent.com/plotly/datasets/master/geojson-counties-fips.json"
DATA_DIR = Path("data")
FY26_PATH = DATA_DIR / "FY26_FMRs.xlsx"
HIST_PATHS = {
    2021: DATA_DIR / "FY2021_FMRs.xlsx",
    2022: DATA_DIR / "FY2022_FMRs.xlsx",
    2023: DATA_DIR / "FY2023_FMRs.xlsx",
    2024: DATA_DIR / "FY2024_FMRs.xlsx",
    2025: DATA_DIR / "FY2025_FMRs.xlsx",
}

BRIEF_SCHEMA = {
    "type": "object",
    "properties": {
        "headline": {"type": "string"},
        "investment_thesis": {"type": "string"},
        "strengths": {"type": "array", "items": {"type": "string"}},
        "risks": {"type": "array", "items": {"type": "string"}},
        "recommendation": {
            "type": "string",
            "enum": ["Investigate", "Watch", "Pass"],
        },
        "confidence_note": {"type": "string"},
    },
    "required": [
        "headline",
        "investment_thesis",
        "strengths",
        "risks",
        "recommendation",
        "confidence_note",
    ],
}


@st.cache_data(show_spinner=False)
def load_fy26_from_path(path: str) -> pd.DataFrame:
    try:
        with open(path, "rb") as f:
            raw_bytes = f.read()
        df = _safe_read_xlsx(raw_bytes)
        df.columns = df.columns.str.strip().str.lower()
        return df
    except Exception as e:
        st.error(f"Could not load FY26 file: {e}")
        return pd.DataFrame()


def _safe_read_xlsx(raw_bytes: bytes) -> pd.DataFrame:
    """Read xlsx bytes, stripping corrupted docProps metadata if needed."""
    import io as _io
    import zipfile

    try:
        return pd.read_excel(_io.BytesIO(raw_bytes), engine="openpyxl")
    except Exception:
        zin = zipfile.ZipFile(_io.BytesIO(raw_bytes))
        buf = _io.BytesIO()
        zout = zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED)
        for item in zin.namelist():
            if "docProps" not in item:
                zout.writestr(item, zin.read(item))
        zout.close()
        buf.seek(0)
        return pd.read_excel(buf, engine="openpyxl")


@st.cache_data(ttl=3600, show_spinner=False)
def parse_historical_file(_file_bytes: bytes, year: int) -> pd.DataFrame:
    """Parse a single uploaded historical FMR xlsx (passed as bytes) and tag with year."""
    try:
        df = _safe_read_xlsx(_file_bytes)
        df.columns = df.columns.str.strip().str.lower()

        if "stusps" not in df.columns:
            for candidate in ["state_alpha"]:
                if candidate in df.columns:
                    df = df.rename(columns={candidate: "stusps"})
                    break

        if "countyname" not in df.columns:
            for candidate in ["areaname", "county_town_name"]:
                if candidate in df.columns:
                    df = df.rename(columns={candidate: "countyname"})
                    break

        df["fmr_2"] = pd.to_numeric(df.get("fmr_2", np.nan), errors="coerce")
        pop_col = next((c for c in df.columns if c.startswith("pop")), None)
        df["pop"] = pd.to_numeric(df[pop_col], errors="coerce").fillna(0) if pop_col else 0

        if "metro" not in df.columns:
            df["metro"] = 0

        df["year"] = year
        df = df.dropna(subset=["stusps", "countyname", "fmr_2"])
        return df[["stusps", "countyname", "fmr_2", "pop", "metro", "year"]]
    except Exception as e:
        st.warning(f"Could not parse FY{year} file: {e}")
        return pd.DataFrame()


@st.cache_data(ttl=3600, show_spinner=False)
def load_historical_from_disk(paths: dict[int, Path]) -> pd.DataFrame:
    frames = []
    for year, path in paths.items():
        if Path(path).exists():
            with open(path, "rb") as f:
                file_bytes = f.read()
            tmp = parse_historical_file(file_bytes, year)
            if not tmp.empty:
                frames.append(tmp)

    if frames:
        return pd.concat(frames, ignore_index=True)

    return pd.DataFrame()


def preprocess(df: pd.DataFrame, gap_thresh: float, pop_thresh: int) -> pd.DataFrame:
    df = df.copy()
    df.columns = df.columns.str.strip().str.lower()

    for c in list(df.columns):
        if c == "state_alpha":
            df = df.rename(columns={c: "stusps"})
        if c == "areaname":
            df = df.rename(columns={c: "countyname"})

    df["fmr_2"] = pd.to_numeric(df.get("fmr_2", np.nan), errors="coerce")

    pop_col_used = "pop2023" if "pop2023" in df.columns else "pop2020"
    df["pop2023"] = pd.to_numeric(df.get(pop_col_used, 0), errors="coerce").fillna(0)
    df["_pop_col_used"] = pop_col_used

    if "metro" not in df.columns:
        df["metro"] = 0

    fips_col = next((c for c in df.columns if "fips" in c), None)
    if fips_col:
        df["fips"] = df[fips_col].astype(str).str.zfill(10).str[:5]

    df = df.dropna(subset=["stusps", "countyname", "fmr_2"])
    group_med = df.groupby(["stusps", "metro"])["fmr_2"].median().rename("group_med")
    df = df.join(group_med, on=["stusps", "metro"])
    df["gap_pct"] = (df["fmr_2"] - df["group_med"]) / df["group_med"]
    df["gap_percent"] = (df["gap_pct"] * 100).round(1)
    df["is_outlier"] = (df["gap_pct"] <= gap_thresh) & (df["pop2023"] >= pop_thresh)
    return df


def add_recovery_score(base: pd.DataFrame, multi: pd.DataFrame) -> pd.DataFrame:
    if multi.empty:
        for col in [
            "cagr",
            "trend_r",
            "score_underval",
            "score_momentum",
            "score_consist",
            "recovery_score",
        ]:
            base[col] = np.nan
        base["opportunity_tier"] = "N/A"
        return base

    pivot = multi.pivot_table(
        index=["stusps", "countyname"],
        columns="year",
        values="fmr_2",
        aggfunc="first",
    ).reset_index()
    year_cols = sorted([c for c in pivot.columns if isinstance(c, int)])

    def momentum(row):
        vals = [(yr, row[yr]) for yr in year_cols if pd.notna(row.get(yr))]
        if len(vals) < 2:
            return pd.Series({"cagr": np.nan, "trend_r": np.nan})

        yrs = np.array([v[0] for v in vals])
        rents = np.array([v[1] for v in vals])
        n = yrs[-1] - yrs[0]
        cagr = ((rents[-1] / rents[0]) ** (1 / n) - 1) * 100 if (rents[0] > 0 and n > 0) else np.nan
        _, _, r, _, _ = stats.linregress(yrs, rents) if len(vals) >= 3 else (0, 0, 0, 0, 0)

        return pd.Series(
            {
                "cagr": round(float(cagr), 2) if pd.notna(cagr) else np.nan,
                "trend_r": round(abs(float(r)), 3),
            }
        )

    mom = pivot.apply(momentum, axis=1)
    pivot = pd.concat([pivot[["stusps", "countyname"]], mom], axis=1)

    df = base.merge(pivot, on=["stusps", "countyname"], how="left")
    df["score_underval"] = np.clip(40 * (-df["gap_pct"] / 0.5), 0, 40).round(1)
    cagr_95 = df["cagr"].quantile(0.95) if df["cagr"].notna().any() else 1
    df["score_momentum"] = np.clip(
        35 * (df["cagr"].clip(lower=0) / max(float(cagr_95), 0.01)),
        0,
        35,
    ).round(1)
    df["score_consist"] = (df["trend_r"].fillna(0) * 25).round(1)
    df["recovery_score"] = (
        df["score_underval"] + df["score_momentum"] + df["score_consist"]
    ).round(1)
    df["opportunity_tier"] = pd.cut(
        df["recovery_score"],
        bins=[0, 30, 50, 70, 100],
        labels=["Weak", "Moderate", "Strong", "Prime"],
        include_lowest=True,
    ).astype(str)
    df.loc[df["recovery_score"].isna(), "opportunity_tier"] = "N/A"
    return df


@st.cache_data(ttl=7 * 24 * 3600, show_spinner=False)
def load_county_geojson() -> dict:
    response = requests.get(COUNTY_GEOJSON_URL, timeout=120)
    response.raise_for_status()
    return response.json()


@st.cache_data(ttl=7 * 24 * 3600, show_spinner=False)
def build_literal_neighbor_lookup() -> dict[str, list[str]]:
    from shapely.geometry import shape
    from shapely.strtree import STRtree

    geojson = load_county_geojson()
    fips_codes: list[str] = []
    geometries = []

    for feature in geojson.get("features", []):
        fips = str(feature.get("id", "")).zfill(5)
        if not fips or not feature.get("geometry"):
            continue
        geom = shape(feature["geometry"])
        if not geom.is_valid:
            geom = geom.buffer(0)
        if geom.is_empty:
            continue
        fips_codes.append(fips)
        geometries.append(geom)

    tree = STRtree(geometries)
    neighbors = {fips: set() for fips in fips_codes}

    for idx, geom in enumerate(geometries):
        candidate_idx = tree.query(geom)
        for other_idx in candidate_idx:
            other_idx = int(other_idx)
            if other_idx <= idx:
                continue
            other_geom = geometries[other_idx]
            if geom.touches(other_geom):
                left = fips_codes[idx]
                right = fips_codes[other_idx]
                neighbors[left].add(right)
                neighbors[right].add(left)

    return {fips: sorted(list(adjacent)) for fips, adjacent in neighbors.items()}


def add_literal_neighbor_metrics(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    if "fips" not in df.columns:
        df["neighbor_count"] = 0
        df["neighbor_median"] = np.nan
        df["neighbor_gap_pct"] = np.nan
        df["dual_flag"] = False
        return df

    lookup = build_literal_neighbor_lookup()
    rent_map = (
        df.dropna(subset=["fips"])
        .drop_duplicates(subset=["fips"])
        .set_index("fips")["fmr_2"]
        .to_dict()
    )

    neighbor_counts = []
    neighbor_medians = []

    for fips in df["fips"].astype(str):
        adjacent_fips = lookup.get(fips, [])
        neighbor_rents = [rent_map[n] for n in adjacent_fips if n in rent_map]
        neighbor_counts.append(len(neighbor_rents))
        neighbor_medians.append(float(np.median(neighbor_rents)) if neighbor_rents else np.nan)

    df["neighbor_count"] = neighbor_counts
    df["neighbor_median"] = neighbor_medians
    df["neighbor_gap_pct"] = ((df["fmr_2"] - df["neighbor_median"]) / df["neighbor_median"] * 100).round(1)
    df["dual_flag"] = df["is_outlier"] & df["neighbor_gap_pct"].notna() & (df["neighbor_gap_pct"] < -15)
    return df


def get_secret(name: str, default=None):
    try:
        return st.secrets[name]
    except Exception:
        return default


def build_county_brief_prompt(row: pd.Series) -> str:
    schema_text = json.dumps(BRIEF_SCHEMA, indent=2)

    optional_metrics = []
    if pd.notna(row.get("neighbor_gap_pct", np.nan)):
        optional_metrics.append(f"Literal Neighbor Gap: {row.get('neighbor_gap_pct')}")
    if pd.notna(row.get("neighbor_count", np.nan)):
        optional_metrics.append(f"Adjacent County Count: {row.get('neighbor_count')}")
    if pd.notna(row.get("cagr", np.nan)):
        optional_metrics.append(f"CAGR: {row.get('cagr')}")
    if pd.notna(row.get("recovery_score", np.nan)):
        optional_metrics.append(f"Recovery Score: {row.get('recovery_score')}")
    if row.get("opportunity_tier", "N/A") != "N/A":
        optional_metrics.append(f"Opportunity Tier: {row.get('opportunity_tier')}")

    optional_text = "\n".join(optional_metrics) if optional_metrics else "No additional metrics available."

    return f"""
You are a real estate investment analyst.

Create a short, client-ready county investment brief using ONLY the structured data below.
Do not invent external facts.
Do not mention any missing data.
Do not use markdown fences.
If recovery metrics are available, include them.
If recovery metrics are not available, write the brief using only the base market signals.
Be concise, grounded, and professional.

Base Data:
County: {row.get('countyname', 'N/A')}
State: {row.get('stusps', 'N/A')}
2BR FMR: {row.get('fmr_2', 'N/A')}
Metro Median: {row.get('group_med', 'N/A')}
Gap Percent: {row.get('gap_percent', 'N/A')}
Population: {row.get('pop2023', 'N/A')}

Additional Available Metrics:
{optional_text}

Return ONLY valid JSON matching this schema exactly:
{schema_text}
"""


def _parse_model_json(raw_text: str) -> dict:
    text = (raw_text or "").strip()
    if not text:
        raise ValueError("Model returned an empty response.")

    if text.startswith("```"):
        text = text.strip("`").strip()
        if text.lower().startswith("json"):
            text = text[4:].strip()

    try:
        return json.loads(text)
    except json.JSONDecodeError:
        start = text.find("{")
        end = text.rfind("}")
        if start != -1 and end != -1 and end > start:
            candidate = text[start : end + 1]
            try:
                return json.loads(candidate)
            except json.JSONDecodeError:
                pass

    raise ValueError(f"Model did not return valid JSON. Raw response: {text[:600]}")


def generate_county_brief_ollama(row: pd.Series) -> dict:
    url = get_secret("ollama_url", "http://localhost:11434/api/generate")
    model = get_secret("ollama_model", "qwen3:4b")
    prompt = build_county_brief_prompt(row)

    response = requests.post(
        url,
        json={
            "model": model,
            "prompt": prompt,
            "format": BRIEF_SCHEMA,
            "stream": False,
            "think": False,
            "options": {"temperature": 0.1},
        },
        timeout=120,
    )
    response.raise_for_status()
    payload = response.json()

    raw_text = (payload.get("response") or "").strip()
    if not raw_text:
        raise ValueError(
            f"Ollama returned empty content. Thinking trace: {(payload.get('thinking') or '')[:300]}"
        )

    brief = _parse_model_json(raw_text)
    for key in BRIEF_SCHEMA["required"]:
        if key not in brief:
            raise ValueError(f"Model response is missing required key: {key}")
    return brief


def generate_county_brief_gemini(row: pd.Series) -> dict:
    api_key = get_secret("gemini_api_key", "")
    model = get_secret("gemini_model", "gemini-2.5-flash")

    if not api_key:
        raise ValueError("Missing gemini_api_key in Streamlit secrets.")

    client = genai.Client(api_key=api_key)
    prompt = build_county_brief_prompt(row)

    response = client.models.generate_content(
        model=model,
        contents=prompt,
        config=types.GenerateContentConfig(
            temperature=0.1,
            response_mime_type="application/json",
            response_json_schema=BRIEF_SCHEMA,
        ),
    )

    brief = _parse_model_json(response.text)
    for key in BRIEF_SCHEMA["required"]:
        if key not in brief:
            raise ValueError(f"Model response is missing required key: {key}")
    return brief


def generate_county_brief(row: pd.Series) -> dict:
    backend = str(get_secret("ai_backend", "ollama")).strip().lower()
    if backend == "ollama":
        return generate_county_brief_ollama(row)
    if backend == "gemini":
        return generate_county_brief_gemini(row)
    raise ValueError(f"Unsupported ai_backend: {backend}")


# ── Sidebar ────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 🏚️ Distressed Property Scouts")
    st.caption("Team 15  ·  Visual Analytics")
    st.markdown("---")

    st.markdown("**Data Source**")
    st.caption("Using preloaded HUD FMR files from the app folder.")

    gap_thresh_pct = st.slider("Gap% outlier threshold", -50, -10, -25, 5)
    pop_thresh_k = st.number_input("Min population (thousands)", value=50, step=10)

    state_filter_list = st.multiselect(
        "Filter by state(s)",
        options=ALL_STATES,
        default=[],
        placeholder="All states (no filter)",
    )

    st.markdown("---")
    show_recovery = st.toggle("📈 Phase 2B: Recovery Score (multi-year)", value=True)

    if FY26_PATH.exists():
        st.caption(f"✅ FY26 file loaded: {FY26_PATH.name}")
    else:
        st.caption("❌ FY26 file missing")

    hist_count = sum(1 for p in HIST_PATHS.values() if p.exists())
    st.caption(f"✅ Historical files found: {hist_count}/5")

    st.markdown("---")
    st.caption("Data: HUD Office of Policy Dev & Research")
    st.caption(f"App loaded: {datetime.now().strftime('%Y-%m-%d %H:%M')}")


# ── Header ─────────────────────────────────────────────────────────
st.markdown("# 🏚️ THE DISTRESSED PROPERTY SCOUTS")
st.caption("Flagging undervalued rental markets ripe for renovation and upsizing.")
st.markdown("---")

with st.expander("📋 Problem Statement & Client Context", expanded=False):
    st.markdown("""
    Our client is a **real estate investor** searching for counties where rents look abnormally low and may signal undervalued markets.  
    The business challenge is that not every low-rent county is a good opportunity — some are weak markets, while others may be mispriced relative to surrounding areas.  
    Using **HUD Fair Market Rent (FMR)** data, this app flags counties with unusually low 2-bedroom rents and helps separate **distressed opportunities** from **value traps**.  
    The goal is to support faster, more informed market screening before deeper investment analysis.
    """)

if not FY26_PATH.exists():
    st.error("Missing preloaded file: data/FY26_FMRs.xlsx")
    st.stop()

with st.spinner("Loading data…"):
    raw = load_fy26_from_path(str(FY26_PATH))
if raw.empty:
    st.stop()

st.caption(f"📅 Data source: **{FY26_PATH.name}** · Loaded {datetime.now().strftime('%Y-%m-%d %H:%M')}")

with st.spinner("Scoring counties…"):
    full_df = preprocess(raw, gap_thresh_pct / 100, int(pop_thresh_k * 1000))

pop_col = full_df["_pop_col_used"].iloc[0] if "_pop_col_used" in full_df.columns else "pop2023"
if pop_col == "pop2020":
    st.info(
        "ℹ️ Population data: using **pop2020** (pop2023 not found in this file). Population filter may be slightly off."
    )
full_df = full_df.drop(columns=["_pop_col_used"], errors="ignore")

if show_recovery:
    with st.spinner("Loading historical files…"):
        multi = load_historical_from_disk(HIST_PATHS)

    years_loaded = sorted(multi["year"].unique().tolist()) if not multi.empty else []
    if not multi.empty:
        st.caption(f"📅 Historical data loaded: **{years_loaded}**")
        full_df = add_recovery_score(full_df, multi)
    else:
        st.info("Historical files not found in /data. Recovery Score is unavailable.")
        for col in ["cagr", "trend_r", "score_underval", "score_momentum", "score_consist", "recovery_score"]:
            full_df[col] = np.nan
        full_df["opportunity_tier"] = "N/A"
else:
    for col in ["cagr", "trend_r", "score_underval", "score_momentum", "score_consist", "recovery_score"]:
        full_df[col] = np.nan
    full_df["opportunity_tier"] = "N/A"

if "fips" in full_df.columns:
    with st.spinner("Building literal county-neighbor graph…"):
        full_df = add_literal_neighbor_metrics(full_df)
else:
    full_df["neighbor_count"] = 0
    full_df["neighbor_median"] = np.nan
    full_df["neighbor_gap_pct"] = np.nan
    full_df["dual_flag"] = False

if state_filter_list:
    df = full_df[full_df["stusps"].isin(state_filter_list)].copy()
    if df.empty:
        st.warning("No data for selected states.")
        st.stop()
else:
    df = full_df.copy()

outliers = df[df["is_outlier"]].copy()
has_recovery_metrics = outliers["recovery_score"].notna().any()

# ── KPI cards ──────────────────────────────────────────────────────
c1, c2, c3, c4 = st.columns(4)
kpi_data = [
    (f"{len(df):,}", "Total Counties Loaded"),
    (f"{len(outliers)}", "Outlier Counties Flagged"),
    (f"${df['fmr_2'].mean():,.0f}", "Overall Avg 2BR FMR"),
    (f"{outliers['gap_percent'].min():.1f}%" if not outliers.empty else "—", "Deepest Discount vs Metro"),
]
for col, (val, lbl) in zip([c1, c2, c3, c4], kpi_data):
    with col:
        st.markdown(
            f'<div class="kpi"><div class="kpi-val">{val}</div><div class="kpi-lbl">{lbl}</div></div>',
            unsafe_allow_html=True,
        )

outlier_pct = len(outliers) / len(df) * 100 if len(df) > 0 else 0
st.caption(
    f"ℹ️ Current threshold ({gap_thresh_pct}%) flags **{len(outliers)} counties** ({outlier_pct:.1f}% of loaded data) with population ≥ {int(pop_thresh_k)}k."
)

st.markdown("---")

# ── Tabs ───────────────────────────────────────────────────────────
t1, t2, t3, t4, t5 = st.tabs([
    "📊 Primary — Outlier Detection",
    "🗺️ Geographic View",
    "📍 Phase 2A — Neighbor Gap",
    "🔥 Phase 2B — Recovery Score",
    "🤖 AI County Brief",
])

with t1:
    st.markdown('<div class="sec-hdr">2BR FMR by County — Sorted Ascending</div>', unsafe_allow_html=True)
    st.caption(
        "🔴 Red = flagged outlier  ·  🟣 Purple = normal  ·  Orange dashed = overall average  ·  Ascending sort puts outliers left."
    )

    ds = df.sort_values("fmr_2").reset_index(drop=True)
    fig = go.Figure(
        go.Bar(
            x=ds["countyname"] + ", " + ds["stusps"],
            y=ds["fmr_2"],
            marker_color=ds["is_outlier"].map({True: "#FF4500", False: "#7B68EE"}),
            hovertemplate="<b>%{x}</b><br>2BR FMR: $%{y:,.0f}<extra></extra>",
        )
    )
    fig.add_hline(
        y=df["fmr_2"].mean(),
        line_dash="dash",
        line_color="#FF8C00",
        line_width=2,
        annotation_text=f"Avg ${df['fmr_2'].mean():,.0f}",
        annotation_position="top right",
    )
    fig.update_layout(
        **DARK,
        height=420,
        showlegend=False,
        xaxis=dict(showticklabels=(len(ds) < 100), tickangle=-45, gridcolor="#1E1E3A"),
        yaxis=dict(title="2BR FMR ($)", gridcolor="#1E1E3A"),
        margin=dict(l=60, r=20, t=20, b=80),
    )
    st.plotly_chart(fig, use_container_width=True)

    fig_h = px.histogram(
        df,
        x="fmr_2",
        nbins=50,
        color="is_outlier",
        color_discrete_map={True: "#FF4500", False: "#7B68EE"},
        labels={"fmr_2": "2BR FMR ($)", "is_outlier": "Outlier"},
        template="plotly_dark",
    )
    fig_h.update_layout(
        **DARK,
        height=280,
        margin=dict(l=60, r=20, t=10, b=40),
        xaxis_title="2BR FMR ($)",
        yaxis_title="Count",
    )
    st.plotly_chart(fig_h, use_container_width=True)

    st.markdown('<div class="sec-hdr">Flagged Outlier Counties</div>', unsafe_allow_html=True)
    if outliers.empty:
        st.info("No outliers at current threshold. Try loosening the slider.")
    else:
        show_cols = ["stusps", "countyname", "fmr_2", "group_med", "gap_percent", "pop2023", "neighbor_gap_pct"]
        if has_recovery_metrics:
            show_cols += ["cagr", "recovery_score", "opportunity_tier"]
        tbl = outliers[show_cols].sort_values("gap_percent").copy()
        rename_map = {
            "stusps": "State",
            "countyname": "County",
            "fmr_2": "2BR FMR ($)",
            "group_med": "Metro Median ($)",
            "gap_percent": "Gap (%)",
            "pop2023": "Population",
            "neighbor_gap_pct": "Literal Neighbor Gap (%)",
            "cagr": "CAGR (%/yr)",
            "recovery_score": "Recovery Score",
            "opportunity_tier": "Tier",
        }
        tbl = tbl.rename(columns={k: v for k, v in rename_map.items() if k in tbl.columns})
        fmt = {
            "2BR FMR ($)": "${:,.0f}",
            "Metro Median ($)": "${:,.0f}",
            "Gap (%)": "{:+.1f}%",
            "Population": "{:,.0f}",
            "Literal Neighbor Gap (%)": "{:+.1f}%",
            "CAGR (%/yr)": "{:+.2f}%",
            "Recovery Score": "{:.1f}",
        }
        st.dataframe(tbl.style.format(fmt), use_container_width=True, height=380)

with t2:
    st.markdown('<div class="sec-hdr">Geographic Rent Heatmap</div>', unsafe_allow_html=True)
    st.caption(
        "County-level 2BR FMR. Dark purple = low rent · Gold = mid · White = high. Geographic clustering reveals regional patterns invisible in the table view."
    )

    if "fips" not in df.columns:
        st.info("FIPS codes not found in dataset — map requires county FIPS column.")
    else:
        map_mode = st.radio("Show on map:", ["All counties", "Outliers only"], horizontal=True, index=0)
        map_df = df if map_mode == "All counties" else outliers

        if map_df.empty:
            st.info("No counties to display with current filter.")
        else:
            fig_map = px.choropleth(
                map_df,
                geojson=COUNTY_GEOJSON_URL,
                locations="fips",
                color="fmr_2",
                hover_name="countyname",
                hover_data={"fmr_2": ":$,.0f", "gap_percent": ":.1f", "is_outlier": True, "stusps": True},
                range_color=[df["fmr_2"].quantile(0.05), df["fmr_2"].quantile(0.85)],
                color_continuous_scale=[
                    [0.00, "#1B102B"],
                    [0.20, "#3B1F6E"],
                    [0.40, "#2563EB"],
                    [0.60, "#06B6D4"],
                    [0.80, "#FACC15"],
                    [1.00, "#F97316"],
                ],
                scope="usa",
                labels={"fmr_2": "2BR FMR ($)"},
            )
            if state_filter_list:
                fig_map.update_geos(fitbounds="locations", visible=False)
            fig_map.update_layout(**DARK, geo_bgcolor="#0D0D1A", height=500, margin=dict(l=0, r=0, t=10, b=0))
            st.plotly_chart(fig_map, use_container_width=True)

with t3:
    st.markdown('<div class="sec-hdr">Phase 2A — Neighbor-Adjusted Distress Score</div>', unsafe_allow_html=True)
    st.info(
        "**Literal-neighbor view:** This version compares each county to its **adjacent counties that share a boundary**. "
        "A county that is cheap vs both its metro peers and its touching neighbors is a stronger local anomaly than a county that is merely cheap in a weak region."
    )

    if "fips" not in df.columns:
        st.warning("FIPS codes required for literal neighbor scoring. Not in this data slice.")
    else:
        dual = df[df["dual_flag"]].sort_values("neighbor_gap_pct")

        st.markdown(
            f"**{len(dual)} counties are cheap vs both metro peers AND literal neighbors** — these are the strongest local anomaly candidates."
        )

        ca, cb = st.columns([3, 2])
        with ca:
            fig_ng = go.Figure()
            norm_n = df[df["gap_pct"] > -0.10]
            fig_ng.add_trace(
                go.Scatter(
                    x=norm_n["gap_percent"],
                    y=norm_n["neighbor_gap_pct"],
                    mode="markers",
                    marker=dict(color="#3A3A5A", size=4, opacity=0.35),
                    name="Normal",
                    hoverinfo="skip",
                )
            )

            cand_n = df[df["gap_pct"] <= -0.10]
            fig_ng.add_trace(
                go.Scatter(
                    x=cand_n[~cand_n["dual_flag"]]["gap_percent"],
                    y=cand_n[~cand_n["dual_flag"]]["neighbor_gap_pct"],
                    mode="markers",
                    marker=dict(color="#7B68EE", size=7, opacity=0.7),
                    name="Metro outlier only",
                    hovertemplate="<b>%{customdata[0]}</b><br>Adjacent counties used: %{customdata[1]}<extra></extra>",
                    customdata=np.column_stack([
                        (cand_n[~cand_n["dual_flag"]]["countyname"] + ", " + cand_n[~cand_n["dual_flag"]]["stusps"]).values,
                        cand_n[~cand_n["dual_flag"]]["neighbor_count"].astype(str).values,
                    ]),
                )
            )
            fig_ng.add_trace(
                go.Scatter(
                    x=dual["gap_percent"],
                    y=dual["neighbor_gap_pct"],
                    mode="markers",
                    marker=dict(color="#FF4500", size=10, opacity=0.9, line=dict(color="white", width=0.5)),
                    name="Dual outlier",
                    hovertemplate="<b>%{customdata[0]}</b><br>Adjacent counties used: %{customdata[1]}<extra></extra>",
                    customdata=np.column_stack([
                        (dual["countyname"] + ", " + dual["stusps"]).values,
                        dual["neighbor_count"].astype(str).values,
                    ]),
                )
            )
            fig_ng.add_vline(x=-25, line_dash="dash", line_color="#FF8C00", line_width=1.5)
            fig_ng.add_hline(y=-15, line_dash="dash", line_color="#FF8C00", line_width=1.5)
            fig_ng.update_layout(
                **DARK,
                height=400,
                xaxis=dict(title="Gap vs Metro Median (%)", gridcolor="#1E1E3A"),
                yaxis=dict(title="Gap vs Literal Neighbor Median (%)", gridcolor="#1E1E3A"),
                legend=dict(bgcolor="#1A1A2E"),
                margin=dict(l=60, r=20, t=10, b=40),
            )
            st.plotly_chart(fig_ng, use_container_width=True)

        with cb:
            if not dual.empty:
                st.dataframe(
                    dual[["stusps", "countyname", "fmr_2", "gap_percent", "neighbor_gap_pct", "neighbor_count"]]
                    .rename(
                        columns={
                            "stusps": "State",
                            "countyname": "County",
                            "fmr_2": "2BR ($)",
                            "gap_percent": "Metro Gap%",
                            "neighbor_gap_pct": "Literal Neighbor Gap%",
                            "neighbor_count": "Adjacent Counties",
                        }
                    )
                    .style.format(
                        {
                            "2BR ($)": "${:,.0f}",
                            "Metro Gap%": "{:+.1f}%",
                            "Literal Neighbor Gap%": "{:+.1f}%",
                            "Adjacent Counties": "{:,.0f}",
                        }
                    ),
                    use_container_width=True,
                    height=400,
                )
            else:
                st.info("No dual outliers with current filter.")

with t4:
    st.markdown('<div class="sec-hdr">Phase 2B — Recovery Signal Score</div>', unsafe_allow_html=True)

    with st.expander("📐 How is the Recovery Score calculated?"):
        st.markdown(
            """
        The Recovery Score (0–100) is composed of three components:

        | Component | Max Points | What it measures |
        |---|---|---|
        | **Undervaluation** | 40 | How far below metro median the county is (deeper discount = higher score) |
        | **Momentum / CAGR** | 35 | Compound annual growth rate of 2BR FMR from FY2021–FY2025 |
        | **Consistency** | 25 | Absolute regression correlation `|r|` of the rent trend line |

        **Formula:**
        - Undervaluation = `clip(40 × (-gap_pct / 0.50), 0, 40)`
        - Momentum = `clip(35 × (CAGR / 95th-pct CAGR), 0, 35)`
        - Consistency = `|r| × 25`

        **Opportunity Tiers:** Weak (0–30) · Moderate (30–50) · Strong (50–70) · Prime (70–100)

        A "Prime" county is cheap today **and** has documented, consistent rent growth over multiple years — we are not predicting recovery, we are identifying markets where recovery signals are already visible in the data.
        """
        )

    if not show_recovery or df["recovery_score"].isna().all():
        st.info(
            "Enable **Phase 2B: Recovery Score** in the sidebar and upload historical files. This answers the professor's question: *'How do we know rent will increase?'* — by showing counties that are cheap today and already demonstrating multi-year rent momentum."
        )
    else:
        st.info(
            "**The professor's question answered:** a county that is cheap today and has positive multi-year CAGR with steady trend correlation is already showing recovery signals — we are not guessing, we are documenting momentum already present in the HUD data."
        )

        rec = df.dropna(subset=["recovery_score"]).copy()
        opps = rec[(rec["gap_pct"] <= -0.20) & (rec["cagr"] > 0)].sort_values("recovery_score", ascending=False)

        st.markdown('<div class="sec-hdr">Opportunity Quadrant</div>', unsafe_allow_html=True)
        st.caption("Top-left = cheap AND recovering = opportunity zone  ·  Bottom-left = cheap but stagnant = value trap")

        fig_q = go.Figure()
        norm_q = rec[rec["gap_pct"] > -0.20]
        fig_q.add_trace(
            go.Scatter(
                x=norm_q["gap_percent"],
                y=norm_q["cagr"],
                mode="markers",
                marker=dict(color="#3A3A5A", size=4, opacity=0.35),
                name="Normal",
                hoverinfo="skip",
            )
        )
        tier_colors = {"Weak": "#FF8C00", "Moderate": "#FFA500", "Strong": "#7B68EE", "Prime": "#FF4500"}
        for tier, col in tier_colors.items():
            sub = opps[opps["opportunity_tier"] == tier]
            if sub.empty:
                continue
            fig_q.add_trace(
                go.Scatter(
                    x=sub["gap_percent"],
                    y=sub["cagr"],
                    mode="markers",
                    marker=dict(color=col, size=9, opacity=0.9, line=dict(color="white", width=0.5)),
                    name=f"{tier} ({len(sub)})",
                    hovertemplate="<b>%{customdata[0]}, %{customdata[1]}</b><br>Gap: %{x:.1f}%<br>CAGR: %{y:.1f}%/yr<extra></extra>",
                    customdata=sub[["countyname", "stusps"]].values,
                )
            )
        fig_q.add_vline(x=-20, line_dash="dash", line_color="#FF8C00", line_width=1.5)
        fig_q.add_hline(y=0, line_dash="dash", line_color="#FF8C00", line_width=1.5)
        fig_q.update_layout(
            **DARK,
            height=430,
            xaxis=dict(title="Rent Gap vs Metro Median (%)", gridcolor="#1E1E3A"),
            yaxis=dict(title="CAGR (% per year)", gridcolor="#1E1E3A"),
            legend=dict(bgcolor="#1A1A2E"),
            margin=dict(l=60, r=20, t=20, b=40),
        )
        st.plotly_chart(fig_q, use_container_width=True)

        st.markdown('<div class="sec-hdr">Top 15 Recovery Candidates</div>', unsafe_allow_html=True)
        top15 = opps.head(15).copy()
        top15["label"] = top15["countyname"] + ", " + top15["stusps"]
        fig_r = go.Figure()
        fig_r.add_trace(
            go.Bar(x=top15["score_underval"], y=top15["label"], orientation="h", name="Undervaluation (0-40)", marker_color="#FF4500")
        )
        fig_r.add_trace(
            go.Bar(
                x=top15["score_momentum"],
                y=top15["label"],
                orientation="h",
                name="Momentum/CAGR (0-35)",
                marker_color="#7B68EE",
                base=top15["score_underval"],
            )
        )
        fig_r.add_trace(
            go.Bar(
                x=top15["score_consist"],
                y=top15["label"],
                orientation="h",
                name="Consistency (0-25)",
                marker_color="#FF8C00",
                base=top15["score_underval"] + top15["score_momentum"],
            )
        )
        fig_r.update_layout(
            **DARK,
            barmode="stack",
            height=460,
            xaxis_title="Recovery Score (0-100)",
            legend=dict(bgcolor="#1A1A2E"),
            margin=dict(l=200, r=30, t=20, b=40),
        )
        st.plotly_chart(fig_r, use_container_width=True)

with t5:
    st.markdown('<div class="sec-hdr">AI County Brief Generator</div>', unsafe_allow_html=True)

    if has_recovery_metrics:
        st.success("Enhanced Brief Mode: historical trend metrics are available and will be included.")
    else:
        st.info("Base Brief Mode: brief will use current undervaluation and neighbor signals only. Upload historical files to enrich it with CAGR and recovery score.")

    if outliers.empty:
        st.info("No flagged counties available for brief generation.")
    else:
        brief_source = outliers.copy()
        if has_recovery_metrics:
            brief_source = brief_source.sort_values(by=["recovery_score", "gap_percent"], ascending=[False, True])
        else:
            brief_source = brief_source.sort_values(by=["gap_percent"], ascending=[True])

        brief_source = brief_source.reset_index(drop=True)
        brief_source["county_label"] = brief_source["countyname"] + ", " + brief_source["stusps"]

        selected_label = st.selectbox(
            "Select a flagged county",
            brief_source["county_label"].tolist(),
            key="ai_county_select",
        )

        selected_row = brief_source.loc[brief_source["county_label"] == selected_label].iloc[0]

        backend_label = str(get_secret("ai_backend", "ollama")).upper()
        st.caption(f"Current AI backend: {backend_label}")

        s1, s2, s3, s4 = st.columns(4)
        with s1:
            st.metric("2BR FMR", f"${selected_row['fmr_2']:,.0f}")
        with s2:
            st.metric("Gap %", f"{selected_row['gap_percent']:+.1f}%")
        with s3:
            neighbor_val = selected_row.get("neighbor_gap_pct", np.nan)
            st.metric("Literal Neighbor Gap %", f"{neighbor_val:+.1f}%" if pd.notna(neighbor_val) else "N/A")
        with s4:
            if has_recovery_metrics and pd.notna(selected_row.get("recovery_score", np.nan)):
                st.metric("Recovery Score", f"{selected_row['recovery_score']:.1f}")
            else:
                st.metric("Recovery Score", "N/A")

        if st.button("Generate County Brief", key="generate_ai_brief"):
            with st.spinner("Generating county brief..."):
                try:
                    brief = generate_county_brief(selected_row)

                    st.markdown(f"### {brief['headline']}")
                    st.write(brief["investment_thesis"])

                    c_left, c_right = st.columns(2)

                    with c_left:
                        st.markdown("**Strengths**")
                        for item in brief["strengths"]:
                            st.write(f"- {item}")

                    with c_right:
                        st.markdown("**Risks**")
                        for item in brief["risks"]:
                            st.write(f"- {item}")

                    st.markdown(f"**Recommendation:** {brief['recommendation']}")
                    st.caption(brief["confidence_note"])

                except Exception as e:
                    st.error(f"County brief generation failed: {e}")

st.markdown("---")
st.caption(
    "Team 15 — Aditya Pola · Srijan Srivatsava Ganji · Raviteja Thode  |  HUD Office of Policy Development & Research  |  Visual Analytics Spring 2025"
)
