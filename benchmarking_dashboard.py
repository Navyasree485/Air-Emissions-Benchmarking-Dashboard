import os
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(
    page_title="Environmental Benchmarking",
    page_icon="🌿",
    layout="wide",
    initial_sidebar_state="expanded",
)

BASE = Path(__file__).resolve().parent
FILE = Path(os.getenv("BENCHMARKING_EXCEL_PATH", BASE / "Benchmarking_FY26.xlsx"))

NAVY = "#123F5D"
BLUE = "#0878D1"
COMBINED = ["FY26/CY25", "FY25/CY24", "FY24/CY23"]
INDIAN_YEARS = ["FY26", "FY25", "FY24"]
SHORTLIST = ["POSCO", "JSW Steel", "Hyundai Steel", "Cleveland Cliff", "USS"]

BROWN = {
    "FY24/CY23": "#E9C28E",
    "FY25/CY24": "#C8782C",
    "FY26/CY25": "#7A3C10",
    "FY24": "#E9C28E",
    "FY25": "#C8782C",
    "FY26": "#7A3C10",
}
BLUE_PERIOD = {
    "FY24/CY23": "#A9C9EE",
    "FY25/CY24": "#4F8FD5",
    "FY26/CY25": "#155AA8",
}
PERIOD_MAP = {
    "FY26": "FY26/CY25", "CY25": "FY26/CY25",
    "FY25": "FY25/CY24", "CY24": "FY25/CY24",
    "FY24": "FY24/CY23", "CY23": "FY24/CY23",
}
AQI_BANDS = [
    (0, 50, "Good", "#00B050"),
    (50, 100, "Satisfactory", "#92D050"),
    (100, 200, "Moderately Polluted", "#FFFF00"),
    (200, 300, "Poor", "#FFC000"),
    (300, 400, "Very Poor", "#FF0000"),
    (400, 500, "Severe", "#C00000"),
]

st.markdown(
    f"""
<style>
.stApp {{ background:#F4F7FA; color:{NAVY}; }}
.block-container {{ padding-top:.7rem; padding-bottom:2rem; max-width:1500px; }}
#MainMenu, footer {{ visibility:hidden; }}
[data-testid="stSidebar"] {{ background:{NAVY}; }}
[data-testid="stSidebar"] * {{ color:white !important; }}
[data-testid="stSidebar"] button {{ background:transparent !important; border:1px solid #ffffff45 !important; text-align:left !important; }}
.topline {{ height:6px; background:{BLUE}; border-radius:5px; }}
.eyebrow {{ color:{BLUE}; font-size:.73rem; font-weight:900; }}
.page-title {{ color:{NAVY}; font-size:1.7rem; font-weight:900; border-bottom:1px solid #AEBBC5; margin-bottom:12px; }}
.ref, .kpi, .peer-card {{ background:white; border:1px solid #D9E3EB; border-radius:9px; box-shadow:0 2px 7px #123f5d10; }}
.ref {{ padding:14px 17px; border-left:6px solid {BLUE}; min-height:118px; }}
.kpi {{ padding:11px 13px; min-height:88px; }}
.label {{ color:#40576A; font-size:.68rem; font-weight:900; }}
.big {{ color:{NAVY}; font-size:1.25rem; font-weight:900; }}
.small {{ color:#40576A; font-size:.74rem; line-height:1.35; }}
.pill {{ display:inline-block; background:{BLUE}; color:white; padding:6px 12px; border-radius:15px; margin:4px 3px; font-size:.7rem; font-weight:900; }}
.kpi-value {{ color:{NAVY}; font-size:1.1rem; font-weight:900; }}
.peer-card {{ padding:10px 12px; margin-bottom:8px; }}
.rank {{ float:right; color:{BLUE}; font-weight:900; }}
/* Home cards: native buttons, fully clickable and protected from browser dark mode. */
[data-testid="stMain"] .home-card div[data-testid="stButton"] > button {{
    width:100%; min-height:150px; white-space:pre-wrap; text-align:left; justify-content:flex-start;
    background:#FFFFFF !important; color:{NAVY} !important;
    border:1px solid #D9E3EB !important; border-left:7px solid {BLUE} !important;
    border-radius:9px !important; padding:16px !important;
    font-size:.82rem !important; line-height:1.55 !important; font-weight:700 !important;
    box-shadow:0 2px 7px #123f5d10 !important;
}}
[data-testid="stMain"] .home-card div[data-testid="stButton"] > button:hover {{
    border-color:{BLUE} !important; transform:translateY(-2px); box-shadow:0 7px 18px #123f5d20 !important;
}}
div[data-testid="stRadio"] label, div[data-testid="stMultiSelect"] label, div[data-testid="stSelectbox"] label {{ color:{NAVY} !important; font-weight:900 !important; }}
</style>
""",
    unsafe_allow_html=True,
)


@st.cache_data(show_spinner=False)
def load_data(path_string):
    path = Path(path_string)
    if not path.exists():
        return None

    def read(sheet, header, usecols=None, nrows=None):
        return pd.read_excel(
            path, sheet_name=sheet, header=header, usecols=usecols,
            nrows=nrows, engine="openpyxl"
        ).dropna(how="all")

    data = {
        "global": read("Global", 1),
        "aqi": read("AQI", 2),
        "prod": read("Prod_BF BOF", 4),
        "prod_r0": read("Prod_BF BOF R0", 4),
        "tata": read("TataSteel", 1),
        "indian": read("Indian", 1, "B:F", 18),
    }

    numeric = {
        "global": ["CS Production (MTPA)", "Dust load (kg/tcs)2", "SO2 load (kg/tcs)3", "NOx load (kg/tcs)4"],
        "aqi": ["average", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar"],
        "prod": ["CS production", "BF-BOF Share"],
        "prod_r0": ["CS production", "BF-BOF Share"],
        "tata": ["CS Production (MTPA)", "Dust load (kg/tcs)2", "SO2 load (kg/tcs)3", "NOx load (kg/tcs)4"],
        "indian": ["Dust load (kg/tcs)", "SO2 load (kg/tcs)", "NOx load (kg/tcs)"],
    }
    for key, columns in numeric.items():
        for column in columns:
            if column in data[key].columns:
                data[key][column] = pd.to_numeric(data[key][column], errors="coerce")

    # Indian_Peers table only.
    data["indian"] = data["indian"].rename(columns={data["indian"].columns[0]: "Company"})
    data["indian"]["Company"] = data["indian"]["Company"].astype(str).str.replace("India-", "", regex=False)
    data["indian"] = data["indian"][data["indian"]["FY/CY"].isin(INDIAN_YEARS)]

    # Peer data from both source sheets.
    peer_columns = ["Company", "FY/CY", "CS production", "BF-BOF Share", "comment"]
    aliases = {
        "TSL": "Tata Steel", "Tata Steel Ltd": "Tata Steel", "Tata Steel Limited": "Tata Steel",
        "Cleveland-Cliffs": "Cleveland Cliff", "Cleveland Cliffs": "Cleveland Cliff",
        "US Steel": "USS", "U.S. Steel": "USS", "JSW": "JSW Steel", "Hyundai": "Hyundai Steel",
    }
    peer_frames = []
    for key in ["prod", "prod_r0"]:
        for column in peer_columns:
            if column not in data[key].columns:
                data[key][column] = pd.NA
        frame = data[key][peer_columns].copy()
        frame["Company"] = frame["Company"].astype(str).str.strip().replace(aliases)
        frame["FY/CY"] = frame["FY/CY"].astype(str).str.strip().replace(PERIOD_MAP)
        peer_frames.append(frame)
    peers = pd.concat(peer_frames, ignore_index=True).dropna(subset=["CS production", "BF-BOF Share"])
    peers = peers.drop_duplicates(["Company", "FY/CY"], keep="last")

    # Ensure six reference/shortlist companies are available in FY26/CY25.
    required = ["Tata Steel", "POSCO", "USS", "JSW Steel", "Cleveland Cliff", "Hyundai Steel"]
    additions = []
    for company in required:
        has_current = ((peers["Company"] == company) & (peers["FY/CY"] == "FY26/CY25")).any()
        if not has_current:
            previous = peers[(peers["Company"] == company) & (peers["FY/CY"] == "FY25/CY24")]
            if not previous.empty:
                row = previous.iloc[0]
                additions.append({
                    "Company": company, "FY/CY": "FY26/CY25",
                    "CS production": row["CS production"], "BF-BOF Share": row["BF-BOF Share"],
                    "comment": "Latest available carried forward",
                })
    if not ((peers["Company"] == "Tata Steel") & (peers["FY/CY"] == "FY26/CY25")).any():
        additions.append({"Company": "Tata Steel", "FY/CY": "FY26/CY25", "CS production": 30.36, "BF-BOF Share": 0.772, "comment": "Reference value"})
    if additions:
        peers = pd.concat([peers, pd.DataFrame(additions)], ignore_index=True)
    data["prod"] = peers

    # Global combined reporting periods.
    data["global"]["Company"] = data["global"]["Company"].astype(str).str.strip()
    data["global"]["DisplayPeriod"] = data["global"]["FY/CY"].astype(str).str.strip().map(PERIOD_MAP)

    # Add TSUK and TSN to Tata Steel sites from the Global sheet.
    site_aliases = {
        "Tata Steel UK": "TSUK", "TSUK": "TSUK",
        "Tata Steel Netherlands": "TSN", "TSN": "TSN",
    }
    external = data["global"][data["global"]["Company"].isin(site_aliases)].copy()
    external["Site"] = external["Company"].replace(site_aliases)
    site_columns = ["Site", "FY/CY", "CS Production (MTPA)", "Dust load (kg/tcs)2", "SO2 load (kg/tcs)3", "NOx load (kg/tcs)4"]
    for column in site_columns:
        if column not in external.columns:
            external[column] = pd.NA
    data["tata"] = pd.concat([data["tata"][site_columns], external[site_columns]], ignore_index=True)
    data["tata"] = data["tata"].drop_duplicates(["Site", "FY/CY"], keep="last")

    # AQI period mapping and annual average.
    data["aqi"]["Site"] = data["aqi"]["Site"].astype(str).str.strip()
    data["aqi"]["FY"] = data["aqi"]["FY"].astype(str).str.strip()
    data["aqi"]["DisplayPeriod"] = data["aqi"]["FY"].map(PERIOD_MAP)
    month_columns = [m for m in ["Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar"] if m in data["aqi"].columns]
    data["aqi"]["average"] = data["aqi"]["average"].fillna(data["aqi"][month_columns].mean(axis=1, skipna=True))
    return data


D = load_data(str(FILE))
if D is None:
    st.error(f"Workbook not found: {FILE}")
    st.stop()
G, A, P, T, I = D["global"], D["aqi"], D["prod"], D["tata"], D["indian"]
page = st.query_params.get("page", "home")


def navigate(target):
    st.query_params["page"] = target
    st.rerun()


def header(title, subtitle="Environmental Benchmarking"):
    # Compact header: remove the hidden blue eyebrow above the page title.
    st.markdown(
        f'<div class="page-title" style="margin-top:0;padding-top:0">{title}</div>',
        unsafe_allow_html=True,
    )


def fmt(value, digits=2):
    return "N/A" if pd.isna(value) else f"{value:.{digits}f}"


def style_plot(fig, height=540, bottom=90):
    fig.update_layout(
        height=height, plot_bgcolor="white", paper_bgcolor="white",
        font=dict(color=NAVY, size=13), title_font=dict(color=NAVY, size=19),
        legend=dict(font=dict(color=NAVY, size=13), title_font=dict(color=NAVY, size=13)),
        margin=dict(l=55, r=30, t=60, b=bottom),
    )
    fig.update_xaxes(color=NAVY, title_font=dict(color=NAVY, size=14), tickfont=dict(color=NAVY, size=12), gridcolor="#D7E1E9", linecolor=NAVY)
    fig.update_yaxes(color=NAVY, title_font=dict(color=NAVY, size=14), tickfont=dict(color=NAVY, size=12), gridcolor="#D7E1E9", linecolor=NAVY)
    return fig


def minimum_record(df, field, period=None, period_field="FY/CY", name_field="Company"):
    work = df.dropna(subset=[field]).copy()
    work = work[work[field] > 0]
    if period is not None:
        work = work[work[period_field] == period]
    if work.empty:
        return "N/A", "No data"
    row = work.loc[work[field].idxmin()]
    return fmt(row[field]), str(row[name_field])


with st.sidebar:
    st.markdown("## Environmental Benchmarking")
    for label, target in [
        ("Home", "home"), ("Peer Selection", "peer"), ("Global Emissions", "global"),
        ("Indian Peers", "indian"), ("Tata Steel Sites", "tata"),
        ("Ambient Air Quality", "aqi"), ("Best Practices", "practices"),
    ]:
        if st.button(label, key="side_" + target, type="primary" if page == target else "secondary"):
            navigate(target)
    st.divider()
    if st.button("Refresh workbook data"):
        st.cache_data.clear()
        st.rerun()


def home_button(column, key, target, title, value, line1, line2):
    with column:
        st.markdown('<div class="home-card">', unsafe_allow_html=True)
        text = f"{title}\n\n{value}\n{line1}\n{line2}"
        if st.button(text, key=key):
            navigate(target)
        st.markdown('</div>', unsafe_allow_html=True)


def kpi_card(column, label, value, company):
    with column:
        st.markdown(
            f'<div class="kpi"><div class="label">{label}</div>'
            f'<div class="kpi-value">{value}</div><div class="small">{company}</div></div>',
            unsafe_allow_html=True,
        )





# HOME_PAGE_STYLE_START
st.markdown(
    """
<style>
/* Home page only: white cards that remain readable with browser dark mode. */
.home-grid-card {
    display: block !important;
    text-decoration: none !important;
    background: #FFFFFF !important;
    border: 1px solid #D9E3EB !important;
    border-left: 7px solid var(--accent) !important;
    border-radius: 10px !important;
    padding: 16px 17px !important;
    min-height: 180px !important;
    box-shadow: 0 3px 10px rgba(18, 63, 93, 0.10) !important;
    transition: transform .15s ease, box-shadow .15s ease !important;
    margin-bottom: 14px !important;
}
.home-grid-card:hover {
    transform: translateY(-3px) !important;
    box-shadow: 0 9px 24px rgba(18, 63, 93, 0.18) !important;
    border-color: var(--accent) !important;
}
.home-card-kicker {
    display:block !important;
    color:#526B7E !important;
    font-size:0.66rem !important;
    font-weight:900 !important;
    letter-spacing:0.055em !important;
    text-transform:uppercase !important;
}
.home-card-title {
    display:block !important;
    color:#123F5D !important;
    font-size:1.03rem !important;
    font-weight:900 !important;
    margin-top:4px !important;
}
.home-card-primary {
    display:block !important;
    color:var(--accent) !important;
    font-size:1.15rem !important;
    font-weight:900 !important;
    line-height:1.25 !important;
    margin-top:12px !important;
}
.home-card-secondary {
    display:block !important;
    color:#294B63 !important;
    font-size:0.76rem !important;
    font-weight:750 !important;
    line-height:1.45 !important;
    margin-top:8px !important;
}
.home-card-footnote {
    display:block !important;
    color:#667E8F !important;
    font-size:0.68rem !important;
    line-height:1.35 !important;
    margin-top:10px !important;
    padding-top:8px !important;
    border-top:1px solid #E5ECF1 !important;
}
.home-section-label {
    color:#123F5D !important;
    font-size:0.92rem !important;
    font-weight:900 !important;
    margin:10px 0 7px !important;
}
</style>
""",
    unsafe_allow_html=True,
)
# HOME_PAGE_STYLE_END

def render_home():
    """Executive-summary home page. No detail-page logic is changed here."""
    header("Air Emissions Benchmarking | Executive Summary")

    reporting_period = st.selectbox(
        "Reference reporting period",
        ["FY26/CY25", "FY25/CY24", "FY24/CY23"],
        index=0,
        key="home_reporting_period",
    )
    financial_year = reporting_period.split("/")[0]

    def clean_positive(frame, field):
        if field not in frame.columns:
            return frame.iloc[0:0].copy()
        result = frame.dropna(subset=[field]).copy()
        return result[result[field] > 0]

    def lowest(frame, field, name_field):
        result = clean_positive(frame, field)
        if result.empty:
            return "N/A", "No data"
        row = result.loc[result[field].idxmin()]
        return fmt(row[field]), str(row[name_field])

    def value_for(frame, field, name_field, names):
        names = {str(name).strip() for name in names}
        result = clean_positive(frame, field)
        result = result[result[name_field].astype(str).str.strip().isin(names)]
        if result.empty:
            return "N/A"
        return fmt(result.iloc[0][field])

    def combined_frame(frame, period_column="FY/CY"):
        result = frame.copy()
        if "DisplayPeriod" not in result.columns:
            period_map = {
                "FY26": "FY26/CY25", "CY25": "FY26/CY25",
                "FY25": "FY25/CY24", "CY24": "FY25/CY24",
                "FY24": "FY24/CY23", "CY23": "FY24/CY23",
            }
            result["DisplayPeriod"] = result[period_column].astype(str).str.strip().map(period_map)
        return result[result["DisplayPeriod"] == reporting_period].copy()

    def card_html(route, accent, kicker, title, primary, secondary, footnote):
        # Spans are used instead of nested block divs so Streamlit renders one clean clickable card.
        return (
            f'<a class="home-grid-card" style="--accent:{accent}" '
            f'href="?page={route}" target="_self">'
            f'<span class="home-card-kicker">{kicker}</span>'
            f'<span class="home-card-title">{title}</span>'
            f'<span class="home-card-primary">{primary}</span>'
            f'<span class="home-card-secondary">{secondary}</span>'
            f'<span class="home-card-footnote">{footnote}</span>'
            f'</a>'
        )

    # Reference company panel, automated by the selected reporting period.
    peer_period_data = P[P["FY/CY"].astype(str).str.strip() == reporting_period].copy()
    reference = peer_period_data[
        peer_period_data["Company"].astype(str).str.strip().isin(
            ["Tata Steel", "TSL", "Tata Steel Limited"]
        )
    ]
    if reference.empty:
        reference_cs = "N/A"
        reference_share = "N/A"
    else:
        reference_cs = fmt(reference.iloc[0]["CS production"])
        reference_share = f'{reference.iloc[0]["BF-BOF Share"]:.1%}'

    left, right = st.columns([1.05, 1.65])
    with left:
        st.markdown(
            f'<div class="ref"><div class="label">REFERENCE COMPANY</div>'
            f'<div class="big">Tata Steel</div><div class="small">{reporting_period}<br>'
            f'<b>CS Production:</b> {reference_cs} MTPA<br>'
            f'<b>BF-BOF Share:</b> {reference_share}</div></div>',
            unsafe_allow_html=True,
        )
    with right:
        peer_names = ["POSCO", "USS", "JSW Steel", "Cleveland-Cliffs", "Hyundai Steel"]
        peer_pills = "".join(f'<span class="pill">{name}</span>' for name in peer_names)
        st.markdown(
            f'<div class="ref"><div class="label">SCREENED PEER COMPANIES</div>'
            f'{peer_pills}<div class="small">Closest peer set screened on crude-steel scale and BF-BOF route share.</div></div>',
            unsafe_allow_html=True,
        )

    # Automate Global values for the selected combined period.
    global_period = combined_frame(G)
    global_dust = lowest(global_period, "Dust load (kg/tcs)2", "Company")
    global_so2 = lowest(global_period, "SO2 load (kg/tcs)3", "Company")
    global_nox = lowest(global_period, "NOx load (kg/tcs)4", "Company")
    tsl_dust = value_for(global_period, "Dust load (kg/tcs)2", "Company", ["TSL", "Tata Steel", "Tata Steel Limited"])
    tsl_so2 = value_for(global_period, "SO2 load (kg/tcs)3", "Company", ["TSL", "Tata Steel", "Tata Steel Limited"])
    tsl_nox = value_for(global_period, "NOx load (kg/tcs)4", "Company", ["TSL", "Tata Steel", "Tata Steel Limited"])

    # Automate Indian peer values for the FY component of the selected period.
    indian_period = I[I["FY/CY"].astype(str).str.strip() == financial_year].copy()
    indian_dust = lowest(indian_period, "Dust load (kg/tcs)", "Company")
    indian_so2 = lowest(indian_period, "SO2 load (kg/tcs)", "Company")
    indian_nox = lowest(indian_period, "NOx load (kg/tcs)", "Company")

    # Automate Tata Steel site values. TSN/TSUK CY values are mapped to the selected combined period.
    tata_period = combined_frame(T)
    tata_dust = lowest(tata_period, "Dust load (kg/tcs)2", "Site")
    tata_so2 = lowest(tata_period, "SO2 load (kg/tcs)3", "Site")
    tata_nox = lowest(tata_period, "NOx load (kg/tcs)4", "Site")

    # Automate AQI values, including TSN CY mapping when present in the AQI dataset.
    aqi_period = combined_frame(A, period_column="FY")
    if "average" in aqi_period.columns:
        aqi_clean = clean_positive(aqi_period, "average")
    else:
        aqi_clean = aqi_period.iloc[0:0].copy()
    if aqi_clean.empty:
        aqi_best_value, aqi_best_site, aqi_high_value, aqi_high_site = "N/A", "No data", "N/A", "No data"
    else:
        best_aqi_row = aqi_clean.loc[aqi_clean["average"].idxmin()]
        high_aqi_row = aqi_clean.loc[aqi_clean["average"].idxmax()]
        aqi_best_value, aqi_best_site = fmt(best_aqi_row["average"], 1), str(best_aqi_row["Site"])
        aqi_high_value, aqi_high_site = fmt(high_aqi_row["average"], 1), str(high_aqi_row["Site"])

    st.markdown('<div class="home-section-label">Executive benchmarking snapshot</div>', unsafe_allow_html=True)

    # Six cards retained. Only content and appearance are changed.
    row1 = st.columns(3)
    with row1[0]:
        st.markdown(card_html(
            "global", "#00A36C", "GLOBAL BENCHMARK", "Dust Emissions",
            f'{global_dust[0]} kg/tcs · {global_dust[1]}',
            f'Tata Steel: {tsl_dust} kg/tcs',
            f'{reporting_period} · Lower particulate intensity indicates stronger dust control.',
        ), unsafe_allow_html=True)
    with row1[1]:
        st.markdown(card_html(
            "global", "#D88716", "GLOBAL BENCHMARK", "SO₂ Emissions",
            f'{global_so2[0]} kg/tcs · {global_so2[1]}',
            f'Tata Steel: {tsl_so2} kg/tcs',
            f'{reporting_period} · Focus areas include fuel sulphur, desulphurisation and coke-oven gas treatment.',
        ), unsafe_allow_html=True)
    with row1[2]:
        st.markdown(card_html(
            "global", "#E91E63", "GLOBAL BENCHMARK", "NOx Emissions",
            f'{global_nox[0]} kg/tcs · {global_nox[1]}',
            f'Tata Steel: {tsl_nox} kg/tcs',
            f'{reporting_period} · Priority levers include low-NOx burners, SCR/SNCR and combustion optimisation.',
        ), unsafe_allow_html=True)

    row2 = st.columns(3)
    with row2[0]:
        st.markdown(card_html(
            "indian", "#0878D1", "INDIAN PEER BENCHMARK", financial_year,
            f'Dust {indian_dust[0]} · {indian_dust[1]}',
            f'SO₂ {indian_so2[0]} · {indian_so2[1]}  |  NOx {indian_nox[0]} · {indian_nox[1]}',
            'Indian peer comparison supports focused performance-gap identification.',
        ), unsafe_allow_html=True)
    with row2[1]:
        st.markdown(card_html(
            "tata", "#008C95", "TATA STEEL SITES", "Best Site Intensities",
            f'Dust {tata_dust[0]} · {tata_dust[1]}',
            f'SO₂ {tata_so2[0]} · {tata_so2[1]}  |  NOx {tata_nox[0]} · {tata_nox[1]}',
            f'{reporting_period} · Includes Indian sites and TSN/TSUK where data is available.',
        ), unsafe_allow_html=True)
    with row2[2]:
        st.markdown(card_html(
            "aqi", "#F59E0B", "AMBIENT AIR QUALITY", "Annual Average AQI",
            f'Best: {aqi_best_value} · {aqi_best_site}',
            f'Highest: {aqi_high_value} · {aqi_high_site}',
            f'{reporting_period} · Monthly trends and CPCB AQI categories are available on the detail page.',
        ), unsafe_allow_html=True)



# PEER_TILE_COMPACT_STYLE_START
st.markdown(
    """
<style>
/* Peer Selection only: compact recommendation tiles and safe top spacing. */
.peer-card {
    padding: 9px 12px !important;
    margin-bottom: 7px !important;
    min-height: 74px !important;
    box-sizing: border-box !important;
}
.peer-card b {
    display: block !important;
    color: #123F5D !important;
    font-size: 0.94rem !important;
    line-height: 1.15 !important;
    padding-right: 34px !important;
}
.peer-card .small {
    color: #526B7E !important;
    font-size: 0.69rem !important;
    line-height: 1.25 !important;
    margin-top: 3px !important;
}
.peer-card .rank {
    color: #0878D1 !important;
    font-size: 0.84rem !important;
    line-height: 1.1 !important;
}
</style>
""",
    unsafe_allow_html=True,
)
# PEER_TILE_COMPACT_STYLE_END

def render_peer():
    """Peer-selection page only. Other dashboard pages remain unchanged."""
    header("Selecting the Right Benchmarking Peers for Tata Steel", "Benchmarking")
    period = st.selectbox(
        "Reporting year",
        ["FY26/CY25", "FY25/CY24"],
        key="peer_reporting_period",
    )

    source = P.copy()
    source["Company"] = source["Company"].astype(str).str.strip()
    source["FY/CY"] = source["FY/CY"].astype(str).str.strip()

    # Keep one point per company in the selected period.
    selected = (
        source[source["FY/CY"] == period]
        .dropna(subset=["CS production", "BF-BOF Share"])
        .drop_duplicates(subset=["Company"], keep="last")
        .copy()
    )
    selected["Data basis"] = period
    selected["Carried forward"] = False

    # FY26/CY25 can be incomplete in the source table. Add only companies missing
    # from FY26/CY25 using their FY25/CY24 point, and label the basis transparently.
    if period == "FY26/CY25":
        previous = (
            source[source["FY/CY"] == "FY25/CY24"]
            .dropna(subset=["CS production", "BF-BOF Share"])
            .drop_duplicates(subset=["Company"], keep="last")
            .copy()
        )
        missing = previous[~previous["Company"].isin(selected["Company"])].copy()
        missing["FY/CY"] = period
        missing["Data basis"] = "FY25/CY24 latest available"
        missing["Carried forward"] = True
        selected = pd.concat([selected, missing], ignore_index=True)

    data = selected.drop_duplicates(subset=["Company"], keep="first").copy()
    required = ["Tata Steel"] + SHORTLIST
    cluster = data[data["Company"].isin(required)].copy()

    if cluster.empty:
        st.error("No peer data found for the selected reporting period.")
        return

    # The ellipse contains Tata Steel and all five shortlisted companies.
    x_pad = max(2, (cluster["CS production"].max() - cluster["CS production"].min()) * 0.18)
    y_pad = max(0.04, (cluster["BF-BOF Share"].max() - cluster["BF-BOF Share"].min()) * 0.35)

    figure = go.Figure()
    figure.add_shape(
        type="circle",
        x0=cluster["CS production"].min() - x_pad,
        x1=cluster["CS production"].max() + x_pad,
        y0=cluster["BF-BOF Share"].min() - y_pad,
        y1=cluster["BF-BOF Share"].max() + y_pad,
        line=dict(color="#9B5A20", width=2),
        fillcolor="rgba(184,103,34,.10)",
        layer="below",
    )

    universe = data[~data["Company"].isin(required)].copy()
    universe_current = universe[~universe["Carried forward"]]
    universe_carried = universe[universe["Carried forward"]]
    peers = cluster[cluster["Company"] != "Tata Steel"].copy()
    tata = cluster[cluster["Company"] == "Tata Steel"].copy()

    # Universe company names are available on hover. This avoids the FY25 label
    # collisions visible when several companies share BF-BOF values near 100%.
    if not universe_current.empty:
        figure.add_trace(go.Scatter(
            x=universe_current["CS production"],
            y=universe_current["BF-BOF Share"],
            mode="markers",
            text=universe_current["Company"],
            customdata=universe_current[["Data basis"]],
            name="Peer universe",
            marker=dict(color="#D7A66F", size=9),
            hovertemplate=(
                "%{text}<br>CS %{x:.2f} MTPA<br>BF-BOF %{y:.1%}"
                "<br>Basis: %{customdata[0]}<extra></extra>"
            ),
        ))

    if not universe_carried.empty:
        figure.add_trace(go.Scatter(
            x=universe_carried["CS production"],
            y=universe_carried["BF-BOF Share"],
            mode="markers",
            text=universe_carried["Company"],
            customdata=universe_carried[["Data basis"]],
            name="Latest available universe",
            marker=dict(color="#D7A66F", size=9, symbol="circle-open", line=dict(width=2)),
            hovertemplate=(
                "%{text}<br>CS %{x:.2f} MTPA<br>BF-BOF %{y:.1%}"
                "<br>Basis: %{customdata[0]}<extra></extra>"
            ),
        ))

    # Stagger shortlisted labels so company names do not merge.
    label_positions = {
        "POSCO": "top center",
        "JSW Steel": "top right",
        "Hyundai Steel": "bottom center",
        "Cleveland Cliff": "top left",
        "USS": "top left",
    }
    peer_positions = [label_positions.get(name, "top center") for name in peers["Company"]]
    if not peers.empty:
        figure.add_trace(go.Scatter(
            x=peers["CS production"],
            y=peers["BF-BOF Share"],
            mode="markers+text",
            text=peers["Company"],
            textposition=peer_positions,
            customdata=peers[["Data basis"]],
            name="Shortlisted peers",
            marker=dict(color="#B86722", size=12),
            textfont=dict(color="#5B2E0A", size=12),
            hovertemplate=(
                "%{text}<br>CS %{x:.2f} MTPA<br>BF-BOF %{y:.1%}"
                "<br>Basis: %{customdata[0]}<extra></extra>"
            ),
        ))

    if not tata.empty:
        figure.add_trace(go.Scatter(
            x=tata["CS production"],
            y=tata["BF-BOF Share"],
            mode="markers+text",
            text=tata["Company"],
            textposition="middle right",
            customdata=tata[["Data basis"]],
            name="Tata Steel",
            marker=dict(color="#7A3C10", size=16, symbol="diamond"),
            textfont=dict(color=NAVY, size=12),
            hovertemplate=(
                "%{text}<br>CS %{x:.2f} MTPA<br>BF-BOF %{y:.1%}"
                "<br>Basis: %{customdata[0]}<extra></extra>"
            ),
        ))

    figure.update_layout(
        title="Peer universe: crude steel production vs BF-BOF share",
        legend=dict(
            orientation="h",
            x=0.5,
            xanchor="center",
            y=-0.24,
            yanchor="top",
            bgcolor="rgba(255,255,255,0.92)",
        ),
    )
    figure.update_yaxes(
        tickformat=".0%",
        range=[0, max(1.15, data["BF-BOF Share"].max() * 1.08)],
        title="BF-BOF Share",
    )
    figure.update_xaxes(
        range=[0, data["CS production"].max() * 1.10],
        title="Crude Steel Production (MTPA)",
    )

    chart, peer_list = st.columns([2.7, 1])
    with chart:
        st.plotly_chart(style_plot(figure, 590, 135), width="stretch")
        if period == "FY26/CY25" and data["Carried forward"].any():
            carried_count = int(data["Carried forward"].sum())
            st.caption(
                f"Open markers: {carried_count} companies use FY25/CY24 as the latest available value; "
                "all companies remain visible through hover labels."
            )

    with peer_list:
        for rank, company in enumerate(SHORTLIST, 1):
            record = data[data["Company"] == company]
            if record.empty:
                detail = "No data"
                basis = ""
            else:
                detail = (
                    f"{record.iloc[0]['CS production']:.1f} MTPA | "
                    f"{record.iloc[0]['BF-BOF Share']:.1%}"
                )
                basis = str(record.iloc[0]["Data basis"])
            st.markdown(
                f'<div class="peer-card"><span class="rank">{rank:02d}</span>'
                f'<b>{company}</b><div class="small">{detail}</div>'
                f'<div class="small">{basis}</div></div>',
                unsafe_allow_html=True,
            )


def global_bar_figure(data, field, metric, periods, unit="kg/tcs"):
    company_order = data.groupby("Company")[field].mean().sort_values().index.tolist()
    blue_names = {"TSL", "Tata Steel", "Tata Steel Limited", "TSN", "TSUK", "Tata Steel UK", "Tata Steel Netherlands"}
    figure = go.Figure()
    for period in periods:
        period_data = data[data["DisplayPeriod"] == period].set_index("Company").reindex(company_order).reset_index()
        colors = [BLUE_PERIOD[period] if str(company).strip() in blue_names else BROWN[period] for company in period_data["Company"]]
        figure.add_trace(go.Bar(
            name=period, x=period_data["Company"], y=period_data[field], marker_color=colors,
            text=period_data[field].apply(lambda v: "" if pd.isna(v) else f"{v:.2f}"),
            textposition="outside", textfont=dict(color=NAVY, size=10),
            hovertemplate="%{x}<br>" + period + "<br>%{y:.2f} " + unit + "<extra></extra>",
        ))
    figure.update_layout(barmode="group", title=f"{metric} intensity by reporting period")
    figure.update_yaxes(
        title="Crude Steel Production (MTPA)" if metric == "Production"
        else f"{metric} concentration (kg/tcs)"
    )
    figure.update_xaxes(tickangle=-40)
    return figure


def render_global():
    header("Global Air Emissions Benchmarking")
    fields = {
        "Dust": "Dust load (kg/tcs)2",
        "SO₂": "SO2 load (kg/tcs)3",
        "NOx": "NOx load (kg/tcs)4",
        "Production": "CS Production (MTPA)",
    }
    left, right = st.columns([1, 2])
    with left:
        metric = st.radio("Pollutant", list(fields), horizontal=True)
    with right:
        periods = st.multiselect("FY / CY reporting periods", COMBINED, default=COMBINED)
    field = fields[metric]
    data = G[G["DisplayPeriod"].isin(periods)].dropna(subset=[field]).copy()
    data = data[data[field] > 0]
    if data.empty:
        st.warning("No global emission data for the selected filters.")
        return

    best = data.loc[data[field].idxmax() if metric == "Production" else data[field].idxmin()]
    recent = data[data["DisplayPeriod"] == "FY26/CY25"]
    recent_best = None if recent.empty else recent.loc[
        recent[field].idxmax() if metric == "Production" else recent[field].idxmin()
    ]
    tsl = data[data["Company"].isin(["TSL", "Tata Steel", "Tata Steel Limited"])]
    tsl_recent = tsl[tsl["DisplayPeriod"] == "FY26/CY25"]
    tsl_record = None if tsl_recent.empty else tsl_recent.iloc[0]

    columns = st.columns(3)
    global_unit = "MTPA" if metric == "Production" else "kg/tcs"
    benchmark_label = "HIGHEST PRODUCTION" if metric == "Production" else "GLOBAL BENCHMARK"
    recent_label = "FY26/CY25 HIGHEST" if metric == "Production" else "FY26/CY25 BEST"
    kpi_card(columns[0], benchmark_label, f"{fmt(best[field])} {global_unit}", best["Company"])
    kpi_card(columns[1], recent_label, f"{fmt(recent_best[field]) if recent_best is not None else 'N/A'} {global_unit}", recent_best["Company"] if recent_best is not None else "No recent data")
    kpi_card(columns[2], "TATA STEEL FY26/CY25", f"{fmt(tsl_record[field]) if tsl_record is not None else 'N/A'} {global_unit}", "TSL")

    global_unit = "MTPA" if metric == "Production" else "kg/tcs"
    st.plotly_chart(
        style_plot(global_bar_figure(data, field, metric, periods, global_unit), 610, 135),
        width="stretch",
    )


def render_indian():
    header("Indian Peer Air Emissions Benchmarking")
    fields = {
        "Dust": "Dust load (kg/tcs)",
        "SO₂": "SO2 load (kg/tcs)",
        "NOx": "NOx load (kg/tcs)",
        "Production": "CS Production (MTPA)",
    }
    left, right = st.columns([1, 2])
    with left:
        metric = st.radio("Pollutant", list(fields), horizontal=True)
    with right:
        periods = st.multiselect("Financial year", INDIAN_YEARS, default=INDIAN_YEARS)
    field = fields[metric]
    if metric == "Production":
        # Indian_Peers table has emissions only. Production is taken from the
        # Global table for the same Indian companies and selected FY values.
        indian_aliases = {
            "Tata Steel": "TSL", "Tata Steel Limited": "TSL",
            "JSW Steel": "JSW", "Jindal Steel & Power": "JSPL",
            "ArcelorMittal Nippon Steel": "AMNS",
        }
        data = G[
            (G["Country"].astype(str).str.strip() == "India")
            & (G["FY/CY"].astype(str).str.strip().isin(periods))
        ].copy()
        data["Company"] = data["Company"].astype(str).str.strip().replace(indian_aliases)
        data = data.dropna(subset=[field])
    else:
        data = I[I["FY/CY"].isin(periods)].dropna(subset=[field]).copy()
    data = data[data[field] > 0]
    if data.empty:
        st.warning("No Indian emission data for the selected filters.")
        return

    best = data.loc[data[field].idxmax() if metric == "Production" else data[field].idxmin()]
    recent = data[data["FY/CY"] == "FY26"]
    recent_best = None if recent.empty else recent.loc[
        recent[field].idxmax() if metric == "Production" else recent[field].idxmin()
    ]
    tsl_recent = recent[recent["Company"] == "TSL"]
    tsl_record = None if tsl_recent.empty else tsl_recent.iloc[0]

    columns = st.columns(3)
    indian_unit = "MTPA" if metric == "Production" else "kg/tcs"
    benchmark_label = "HIGHEST PRODUCTION" if metric == "Production" else "INDIAN BENCHMARK"
    recent_label = "FY26 HIGHEST" if metric == "Production" else "FY26 BEST"
    kpi_card(columns[0], benchmark_label, f"{fmt(best[field])} {indian_unit}", best["Company"])
    kpi_card(columns[1], recent_label, f"{fmt(recent_best[field]) if recent_best is not None else 'N/A'} {indian_unit}", recent_best["Company"] if recent_best is not None else "No FY26 data")
    kpi_card(columns[2], "TSL FY26", f"{fmt(tsl_record[field]) if tsl_record is not None else 'N/A'} {indian_unit}", "TSL")

    company_order = data.groupby("Company")[field].mean().sort_values().index.tolist()
    figure = go.Figure()
    for period in periods:
        period_data = data[data["FY/CY"] == period].set_index("Company").reindex(company_order).reset_index()
        figure.add_trace(go.Bar(
            name=period, x=period_data["Company"], y=period_data[field], marker_color=BROWN[period],
            text=period_data[field].apply(lambda v: "" if pd.isna(v) else f"{v:.2f}"),
            textposition="outside", textfont=dict(color=NAVY, size=10),
        ))
    figure.update_layout(barmode="group", title=f"Indian peers | {metric} intensity")
    figure.update_yaxes(
        title="Crude Steel Production (MTPA)" if metric == "Production"
        else f"{metric} concentration (kg/tcs)"
    )
    st.plotly_chart(style_plot(figure, 555), width="stretch")


def render_tata():
    header("Tata Steel Site Benchmarking")
    fields = {"Dust": "Dust load (kg/tcs)2", "SO₂": "SO2 load (kg/tcs)3", "NOx": "NOx load (kg/tcs)4", "Production": "CS Production (MTPA)"}
    left, right = st.columns([1, 2])
    with left:
        metric = st.selectbox("Metric", list(fields))
    with right:
        periods = st.multiselect("Reporting period", COMBINED, default=COMBINED)
    field = fields[metric]
    data = T.copy()
    data["DisplayPeriod"] = data["FY/CY"].astype(str).str.strip().map(PERIOD_MAP)
    data = data[data["DisplayPeriod"].isin(periods)].dropna(subset=[field])
    data = data[data[field] > 0]
    if data.empty:
        st.warning("No Tata Steel site data for the selected filters.")
        return

    best = data.loc[data[field].idxmin()]
    recent = data[data["DisplayPeriod"] == "FY26/CY25"]
    recent_best = None if recent.empty else recent.loc[recent[field].idxmin()]
    columns = st.columns(2)
    kpi_card(columns[0], "BEST SITE IN SELECTION", fmt(best[field]) + (" MTPA" if metric == "Production" else " kg/tcs"), best["Site"])
    kpi_card(columns[1], "FY26/CY25 BEST SITE", (fmt(recent_best[field]) if recent_best is not None else "N/A") + (" MTPA" if metric == "Production" else " kg/tcs"), recent_best["Site"] if recent_best is not None else "No recent data")

    sites = sorted(data["Site"].unique())
    figure = go.Figure()
    for period in periods:
        period_data = data[data["DisplayPeriod"] == period].set_index("Site").reindex(sites).reset_index()
        figure.add_trace(go.Bar(
            name=period, x=period_data["Site"], y=period_data[field], marker_color=BLUE_PERIOD[period],
            text=period_data[field].apply(lambda v: "" if pd.isna(v) else f"{v:.2f}"),
            textposition="outside", textfont=dict(color=NAVY, size=10),
        ))
    figure.update_layout(barmode="group", title=f"Tata Steel sites | {metric}")
    figure.update_yaxes(title="MTPA" if metric == "Production" else f"{metric} concentration (kg/tcs)")
    st.plotly_chart(style_plot(figure, 555), width="stretch")


def render_aqi():
    header("Ambient Air Quality Status")
    period = st.selectbox("Reporting period", COMBINED)
    data = A[A["DisplayPeriod"] == period].copy()
    months = [m for m in ["Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar"] if m in data.columns]

    monthly = go.Figure()
    for low, high, label, color in AQI_BANDS:
        monthly.add_hrect(y0=low, y1=high, fillcolor=color, opacity=.92, line_width=0, layer="below")
    site_colors = ["#8B4513", "#FF00FF", "#00A651", "#00AEEF", "#7030A0", "#002060", "#7F6000"]
    for (_, row), color in zip(data.iterrows(), site_colors * 3):
        monthly.add_trace(go.Scatter(
            x=months, y=[row.get(m) for m in months], mode="lines+markers",
            name=str(row["Site"]), line=dict(color=color, width=2), marker=dict(size=5),
        ))
    monthly.update_layout(
        title=f"{period}: Monthly AQI",
        legend=dict(
            orientation="h",
            x=0.5,
            xanchor="center",
            y=-0.22,
            yanchor="top",
            bgcolor="rgba(255,255,255,0.92)",
        ),
    )
    monthly.update_yaxes(range=[0, 500], title="AQI")

    bar_colors = []
    for value in data["average"]:
        color = "#C00000"
        for low, high, _, candidate in AQI_BANDS:
            if low <= value <= high:
                color = candidate
                break
        bar_colors.append(color)
    annual = go.Figure(go.Bar(
        x=data["Site"], y=data["average"], marker_color=bar_colors,
        text=data["average"].round(0), textposition="outside", textfont=dict(color=NAVY, size=11),
    ))
    annual.update_layout(title=f"{period}: Annual Average AQI")
    annual.update_yaxes(title="Annual average AQI")

    left, right = st.columns([2.3, 1])
    with left:
        st.plotly_chart(style_plot(monthly, 500, 135), width="stretch")
    with right:
        st.plotly_chart(style_plot(annual, 500), width="stretch")

    # Restore CPCB AQI colour-scale legend below both charts.
    legend_columns = st.columns(6)
    for column, (low, high, label, color) in zip(legend_columns, AQI_BANDS):
        with column:
            range_text = f"{int(low)}-{int(high)}" if high < 500 else f">{int(low)}"
            st.markdown(
                f'<div style="background:{color};color:#111;text-align:center;padding:8px 3px;'
                f'font-size:.68rem;font-weight:900;border:1px solid #333">{label}<br>{range_text}</div>',
                unsafe_allow_html=True,
            )


def render_practices():
    header("Best Practices in Tata Steel")
    st.info("TSJ: De-NOx at Merchant Mill. TSK: Waste heat recovery and De-NOx systems. TSM: SNCR NH3 injection at BFPP.")


ROUTES = {
    "home": render_home,
    "peer": render_peer,
    "global": render_global,
    "indian": render_indian,
    "tata": render_tata,
    "aqi": render_aqi,
    "practices": render_practices,
}
ROUTES.get(page, render_home)()
