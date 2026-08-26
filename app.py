# ============================================================
# M2P Consulting – Airport Solar Feasibility Tool
# Version 2.0  |  26 August 2026
# ============================================================
import streamlit as st
import pandas as pd
import numpy_financial as npf
import plotly.graph_objects as go
import plotly.express as px

# ── BRAND COLOURS ──────────────────────────────────────────
M2P_NAVY   = "#063443"
M2P_TEAL   = "#184A41"
M2P_BLUE   = "#2581C4"
M2P_GREEN  = "#008A65"
M2P_AMBER  = "#BD7119"
M2P_LIGHT  = "#F2F2F2"
RED        = "#C0392B"
AMBER_WARN = "#E67E22"
GREEN_OK   = "#27AE60"

# ── PAGE CONFIG ────────────────────────────────────────────
st.set_page_config(
    page_title="M2P Solar Feasibility Tool",
    page_icon="☀️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── GLOBAL CSS ─────────────────────────────────────────────
st.markdown(f"""
<style>
    /* Base */
    html, body, [class*="css"] {{
        font-family: 'Century Gothic', 'Trebuchet MS', sans-serif;
        background-color: #FAFAFA;
    }}
    /* Header bar */
    .m2p-header {{
        background: linear-gradient(135deg, {M2P_NAVY} 0%, {M2P_TEAL} 100%);
        padding: 1.4rem 2rem;
        border-radius: 0 0 8px 8px;
        margin-bottom: 1.5rem;
    }}
    .m2p-header h1 {{
        color: white;
        font-size: 1.6rem;
        font-weight: 700;
        margin: 0;
        letter-spacing: 0.3px;
    }}
    .m2p-header p {{
        color: rgba(255,255,255,0.72);
        font-size: 0.85rem;
        margin: 0.25rem 0 0 0;
    }}
    /* Gate cards */
    .gate-card {{
        background: white;
        border: 1px solid #E8E8E8;
        border-left: 4px solid {M2P_BLUE};
        border-radius: 6px;
        padding: 1.2rem 1.4rem;
        margin-bottom: 1rem;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }}
    .gate-card h3 {{
        color: {M2P_NAVY};
        font-size: 1rem;
        font-weight: 700;
        margin: 0 0 0.6rem 0;
    }}
    .gate-card.warn {{ border-left-color: {AMBER_WARN}; }}
    .gate-card.pass {{ border-left-color: {GREEN_OK}; }}
    .gate-card.fail {{ border-left-color: {RED}; }}
    /* KPI tiles */
    .kpi-row {{ display: flex; gap: 1rem; margin: 1rem 0; flex-wrap: wrap; }}
    .kpi-tile {{
        flex: 1; min-width: 160px;
        background: {M2P_NAVY};
        border-radius: 8px;
        padding: 1.1rem 1.3rem;
        text-align: center;
    }}
    .kpi-tile .val {{
        font-size: 1.9rem;
        font-weight: 700;
        color: white;
        line-height: 1.1;
    }}
    .kpi-tile .lbl {{
        font-size: 0.72rem;
        color: rgba(255,255,255,0.65);
        margin-top: 0.3rem;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }}
    .kpi-tile.amber {{ background: {M2P_AMBER}; }}
    .kpi-tile.teal  {{ background: {M2P_TEAL}; }}
    .kpi-tile.blue  {{ background: {M2P_BLUE}; }}
    .kpi-tile.green {{ background: {M2P_GREEN}; }}
    /* Status badges */
    .badge {{
        display: inline-block;
        padding: 0.18rem 0.7rem;
        border-radius: 20px;
        font-size: 0.75rem;
        font-weight: 700;
        letter-spacing: 0.4px;
    }}
    .badge-pass {{ background: #D5F5E3; color: #1E8449; }}
    .badge-warn {{ background: #FDEBD0; color: #A04000; }}
    .badge-fail {{ background: #FADBD8; color: #922B21; }}
    .badge-info {{ background: #D6EAF8; color: #1A5276; }}
    /* Section labels */
    .section-label {{
        font-size: 0.7rem;
        font-weight: 700;
        color: {M2P_BLUE};
        text-transform: uppercase;
        letter-spacing: 1px;
        margin-bottom: 0.4rem;
    }}
    /* Divider */
    .m2p-divider {{
        border: none;
        border-top: 2px solid {M2P_LIGHT};
        margin: 1.5rem 0;
    }}
    /* Sidebar — base */
    [data-testid="stSidebar"] {{
        background: {M2P_NAVY};
    }}
    [data-testid="stSidebar"] * {{ color: rgba(255,255,255,0.88); }}

    /* Sidebar labels */
    [data-testid="stSidebar"] label,
    [data-testid="stSidebar"] .stSelectbox label,
    [data-testid="stSidebar"] .stNumberInput label,
    [data-testid="stSidebar"] .stSlider label,
    [data-testid="stSidebar"] .stTextInput label {{
        color: rgba(255,255,255,0.70) !important;
        font-size: 0.78rem;
    }}

    /* Sidebar input boxes — white background, dark text so values are readable */
    [data-testid="stSidebar"] input,
    [data-testid="stSidebar"] textarea,
    [data-testid="stSidebar"] .stTextInput input,
    [data-testid="stSidebar"] .stNumberInput input {{
        background-color: #FFFFFF !important;
        color: {M2P_NAVY} !important;
        border: 1px solid rgba(255,255,255,0.25) !important;
        border-radius: 4px !important;
    }}

    /* Sidebar select boxes */
    [data-testid="stSidebar"] .stSelectbox div[data-baseweb="select"] > div,
    [data-testid="stSidebar"] select {{
        background-color: #FFFFFF !important;
        color: {M2P_NAVY} !important;
        border-radius: 4px !important;
    }}

    /* Sidebar slider track and thumb — keep visible on dark bg */
    [data-testid="stSidebar"] .stSlider [data-testid="stTickBar"] {{
        color: rgba(255,255,255,0.50) !important;
    }}

    /* Sidebar markdown / headings */
    [data-testid="stSidebar"] h1,
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3,
    [data-testid="stSidebar"] p,
    [data-testid="stSidebar"] span:not([class]) {{
        color: rgba(255,255,255,0.88) !important;
    }}
    /* Disclaimer banner */
    .disclaimer {{
        background: #EBF5FB;
        border: 1px solid #AED6F1;
        border-radius: 6px;
        padding: 0.8rem 1.2rem;
        font-size: 0.78rem;
        color: #1A5276;
        margin-top: 2rem;
    }}
</style>
""", unsafe_allow_html=True)

# ── HEADER ─────────────────────────────────────────────────
st.markdown("""
<div class="m2p-header">
  <h1>☀️ Airport Solar Feasibility Tool</h1>
  <p>M2P Consulting · Airport Planning & Design · Internal Use Only · v2.0</p>
</div>
""", unsafe_allow_html=True)

# ── SIDEBAR ────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### Project Setup")
    project_name   = st.text_input("Project name", value="Airport Solar Assessment")
    airport_type   = st.selectbox("Airport development type", ["Brownfield", "Greenfield"])
    airport_region = st.selectbox("Region (for solar resource)", [
        "Europe – North (UK, Netherlands, Norway) · 2.4 h",
        "Europe – Central (Germany, France) · 2.9 h",
        "Europe – South (Spain, Italy, Greece) · 4.2 h",
        "Middle East (UAE, Qatar, Saudi) · 5.8 h",
        "South / SE Asia (Malaysia, India) · 4.6 h",
        "North America – East · 3.5 h",
        "North America – West (California) · 5.0 h",
        "Custom",
    ])
    custom_psh = None
    if "Custom" in airport_region:
        custom_psh = st.number_input("Peak sun hours (h/day)", 1.5, 7.0, 4.0, 0.1)

    psh_lookup = {
        "Europe – North": 2.4, "Europe – Central": 2.9, "Europe – South": 4.2,
        "Middle East": 5.8, "South / SE Asia": 4.6,
        "North America – East": 3.5, "North America – West": 5.0,
    }
    peak_sun = custom_psh if custom_psh else next(
        (v for k, v in psh_lookup.items() if k in airport_region), 3.5
    )

    st.markdown("---")
    st.markdown("### Market Assumptions")
    capex_mw       = st.number_input("Solar CAPEX (€ / MW)", 500000, 2000000, 800000, 50000,
                                      help="European brownfield rooftop ~€900k; Middle East ground-mount ~€650k")
    ppa_price      = st.number_input("PPA / avoided cost (€ / kWh)", 0.05, 0.25, 0.09, 0.01)
    opex_pct       = st.slider("OPEX (% of CAPEX / yr)", 0.5, 3.0, 1.5, 0.1) / 100
    ppa_escalation = st.slider("Annual PPA escalation (%)", 0.0, 3.0, 1.5, 0.5) / 100
    opex_inflation = st.slider("Annual OPEX inflation (%)", 0.0, 5.0, 2.5, 0.5) / 100

    st.markdown("---")
    st.markdown("### Finance")
    debt_pct       = st.slider("Debt ratio (%)", 0, 90, 70, 5) / 100
    interest_pct   = st.slider("Interest rate (%)", 2.0, 10.0, 5.5, 0.5) / 100
    loan_years     = st.slider("Loan term (years)", 5, 20, 15, 1)
    tax_pct        = st.slider("Corporate tax rate (%)", 0, 40, 25, 1) / 100
    project_life   = 25


# ── IRRADIANCE PRESETS (kWh/kWp/yr) ───────────────────────
annual_yield_per_kwp = peak_sun * 365 * 0.80   # system_efficiency = 0.80

# ── SOLAR RESOURCE ─────────────────────────────────────────
def regional_yield():
    return annual_yield_per_kwp   # kWh per kWp per year

# ── GATE TABS ──────────────────────────────────────────────
tabs = st.tabs([
    "**Gate 1** Site",
    "**Gate 2** Aviation Safety",
    "**Gate 3** Grid",
    "**Gate 4** Commercial",
    "**Gate 5** GF vs BF",
    "**Gate 6** Sizing & Finance",
    "**Export**",
])

# ── GATE 1 ─────────────────────────────────────────────────
with tabs[0]:
    st.markdown('<div class="section-label">Gate 1 — Site Viability</div>', unsafe_allow_html=True)
    st.markdown("Confirm that the site can physically accommodate solar before any further analysis.")

    c1, c2 = st.columns(2)
    with c1:
        total_site_ha      = st.number_input("Total site area (ha)", 1.0, 10000.0, 500.0, 10.0)
        terrain            = st.selectbox("Terrain type", [
            "Flat / optimal (multiplier: ×1.0)",
            "Gentle slope 0–5% (multiplier: ×1.05 CAPEX, ×1.10 area)",
            "Complex / steep >5% (multiplier: ×1.15 CAPEX, ×1.25 area)",
        ])
        land_ownership     = st.selectbox("Land ownership", [
            "Owned by airport – no acquisition cost",
            "Owned by airport – part leased to third party",
            "Must be acquired / leased",
        ])
    with c2:
        built_pct          = st.slider("Built / operational area (%)", 0, 80, 40, 5)
        green_protected_ha = st.number_input("Protected green / ecology area (ha)", 0.0, 5000.0, 0.0, 10.0,
                                              help="e.g. bird protection zones, flood plain")
        committed_dev_ha   = st.number_input("Area committed to other development (ha)", 0.0, 5000.0, 0.0, 10.0)

    terrain_area_mult = 1.0
    terrain_capex_mult = 1.0
    if "Gentle" in terrain:
        terrain_area_mult = 1.10; terrain_capex_mult = 1.05
    elif "steep" in terrain:
        terrain_area_mult = 1.25; terrain_capex_mult = 1.15

    built_ha = total_site_ha * built_pct / 100
    gate1_available_ha = total_site_ha - built_ha - green_protected_ha - committed_dev_ha
    gate1_available_ha = max(0.0, gate1_available_ha)

    # Acquisition cost
    land_cost_total = 0.0
    if "acquired" in land_ownership:
        land_cost_per_ha = st.number_input("Land acquisition cost (€ / ha)", 0, 5000000, 50000, 10000)
        land_cost_total = gate1_available_ha * land_cost_per_ha

    # Flags
    g1_flags = []
    if gate1_available_ha < 20:
        g1_flags.append(("fail", "Less than 20 ha available — minimum for a meaningful installation"))
    elif gate1_available_ha < 50:
        g1_flags.append(("warn", f"{gate1_available_ha:.0f} ha available — sufficient for ~{gate1_available_ha/2:.0f}–{gate1_available_ha:.0f} MW depending on technology"))
    else:
        g1_flags.append(("pass", f"{gate1_available_ha:.0f} ha available — sufficient for preliminary sizing"))
    if "acquired" in land_ownership and land_cost_total > 0:
        g1_flags.append(("warn", f"Land acquisition adds €{land_cost_total:,.0f} to CAPEX — verify against project economics"))
    if green_protected_ha > 0:
        g1_flags.append(("info", f"{green_protected_ha:.0f} ha excluded as protected — confirm boundary with ecology survey"))

    st.markdown('<hr class="m2p-divider">', unsafe_allow_html=True)
    st.markdown("**Gate 1 assessment**")
    for status, msg in g1_flags:
        badge = {"pass": "badge-pass", "warn": "badge-warn", "fail": "badge-fail", "info": "badge-info"}[status]
        label = status.upper()
        st.markdown(f'<span class="badge {badge}">{label}</span> {msg}', unsafe_allow_html=True)

    st.info(f"**Available area carried forward to Gate 2:** {gate1_available_ha:.1f} ha")

# ── GATE 2 ─────────────────────────────────────────────────
with tabs[1]:
    st.markdown('<div class="section-label">Gate 2 — Aviation Safety Constraints</div>', unsafe_allow_html=True)
    st.markdown("Aviation safety constraints typically reduce usable area. Complete each check before committing to a layout.")

    c1, c2 = st.columns(2)
    with c1:
        glare_study_done  = st.checkbox("Formal glare study completed")
        glare_zones_ha    = st.number_input("Area eliminated by glare constraints (ha)", 0.0, gate1_available_ha, 0.0, 1.0,
                                             disabled=not glare_study_done,
                                             help="From the glare study report")
        ols_study_done    = st.checkbox("Obstacle limitation surface (OLS) mapping completed")
        ols_eliminated_ha = st.number_input("Area eliminated by OLS height restrictions (ha)", 0.0, gate1_available_ha, 0.0, 1.0,
                                             disabled=not ols_study_done)
    with c2:
        rpz_eliminated_ha  = st.number_input("Area within runway protection zones (ha)", 0.0, gate1_available_ha, 0.0, 1.0,
                                              help="No solar permitted within RPZ regardless of other checks")
        naa_consulted      = st.checkbox("Pre-application consultation with national aviation authority completed")
        antenna_ok         = st.checkbox("Confirmed no interference with navigation/communication equipment")

    g2_eliminated = glare_zones_ha + ols_eliminated_ha + rpz_eliminated_ha
    gate2_cleared_ha = max(0.0, gate1_available_ha - g2_eliminated)

    # Flags
    g2_flags = []
    if not glare_study_done:
        g2_flags.append(("warn", "Glare study not yet completed — area available may be overstated"))
    if not ols_study_done:
        g2_flags.append(("warn", "OLS mapping not completed — height restrictions unknown"))
    if rpz_eliminated_ha > 0:
        g2_flags.append(("fail", f"{rpz_eliminated_ha:.0f} ha within runway protection zone — this area is unavailable"))
    if not naa_consulted:
        g2_flags.append(("warn", "National aviation authority consultation required before design is committed"))
    if not antenna_ok:
        g2_flags.append(("fail", "Navigation/communication interference unresolved — do not proceed without clearance"))
    if gate2_cleared_ha > 0 and len([f for f in g2_flags if f[0] == "fail"]) == 0:
        g2_flags.append(("pass", f"{gate2_cleared_ha:.1f} ha cleared for preliminary layout — subject to completing outstanding studies"))

    reduction_pct = (1 - gate2_cleared_ha / gate1_available_ha) * 100 if gate1_available_ha > 0 else 0

    st.markdown('<hr class="m2p-divider">', unsafe_allow_html=True)
    st.markdown("**Gate 2 assessment**")
    for status, msg in g2_flags:
        badge = {"pass": "badge-pass", "warn": "badge-warn", "fail": "badge-fail", "info": "badge-info"}[status]
        st.markdown(f'<span class="badge {badge}">{status.upper()}</span> {msg}', unsafe_allow_html=True)

    col_a, col_b = st.columns(2)
    col_a.metric("Aviation-cleared area", f"{gate2_cleared_ha:.1f} ha")
    col_b.metric("Reduction from Gate 1", f"{reduction_pct:.0f}%",
                 delta=f"−{g2_eliminated:.1f} ha", delta_color="inverse")

# ── GATE 3 ─────────────────────────────────────────────────
with tabs[2]:
    st.markdown('<div class="section-label">Gate 3 — Grid Connection</div>', unsafe_allow_html=True)
    st.markdown("Solar is only useful if the grid can absorb it. Resolve grid capacity before sizing the system.")

    c1, c2 = st.columns(2)
    with c1:
        annual_demand_gwh   = st.number_input("Airport annual electricity demand (GWh/yr)", 0.1, 2000.0, 150.0, 10.0)
        grid_headroom_mw    = st.number_input("Available grid connection headroom (MW)", 0.0, 2000.0, 50.0, 5.0,
                                               help="From distribution network operator (DNO) pre-application enquiry")
        grid_study_done     = st.checkbox("DNO pre-application enquiry completed")
        export_allowed      = st.checkbox("Grid export permitted (excess generation can be sold)")
    with c2:
        self_consume_pct    = st.slider("Estimated self-consumption (%)", 10, 100, 70, 5,
                                         help="The share of generation consumed on site. Higher = better economics")
        add_bess            = st.checkbox("Include battery storage (BESS) to shift generation to peak demand?")
        bess_mwh            = 0.0
        bess_capex_mwh      = 250000.0
        bess_revenue_mwh_yr = 0.0
        if add_bess:
            bess_mwh            = st.number_input("Battery capacity (MWh)", 10.0, 500.0, 50.0, 10.0)
            bess_capex_mwh      = st.number_input("BESS CAPEX (€ / MWh)", 150000, 500000, 250000, 10000)
            bess_revenue_mwh_yr = st.number_input("BESS annual value (€ / MWh installed)", 0, 50000, 8000, 1000,
                                                   help="From arbitrage, capacity payments and ancillary services. Conservative = €5k–€10k/MWh/yr")

    g3_flags = []
    if not grid_study_done:
        g3_flags.append(("warn", "DNO pre-application enquiry not completed — grid headroom is unconfirmed"))
    if grid_headroom_mw < 10:
        g3_flags.append(("fail", "Grid headroom under 10 MW — grid upgrade required before proceeding"))
    elif grid_headroom_mw < 30:
        g3_flags.append(("warn", f"Grid headroom {grid_headroom_mw:.0f} MW — limits maximum capacity; consider BESS to reduce peak export"))
    else:
        g3_flags.append(("pass", f"Grid headroom {grid_headroom_mw:.0f} MW — sufficient for preliminary sizing"))
    if not export_allowed and self_consume_pct < 100:
        g3_flags.append(("warn", "Export not permitted — size system to self-consumption only, or generation will be curtailed"))

    st.markdown('<hr class="m2p-divider">', unsafe_allow_html=True)
    for status, msg in g3_flags:
        badge = {"pass": "badge-pass", "warn": "badge-warn", "fail": "badge-fail", "info": "badge-info"}[status]
        st.markdown(f'<span class="badge {badge}">{status.upper()}</span> {msg}', unsafe_allow_html=True)

# ── GATE 4 ─────────────────────────────────────────────────
with tabs[3]:
    st.markdown('<div class="section-label">Gate 4 — Commercial Structure</div>', unsafe_allow_html=True)
    st.markdown("The commercial model determines who builds it, who owns it, and who carries the risk.")

    commercial_model = st.selectbox("Commercial model", [
        "Direct ownership (airport owns and operates)",
        "Power Purchase Agreement – airport buys output from a third-party developer",
        "Special Purpose Vehicle / concession (developer owns, airport receives carbon benefit)",
        "Rooftop lease (developer pays rent, airport receives rental income)",
    ])

    c1, c2 = st.columns(2)
    with c1:
        airport_ownership  = st.selectbox("Airport ownership", [
            "Public (government / municipal)",
            "Private",
            "Mixed (public majority)",
        ])
        can_fund_capex     = st.checkbox("Airport can fund full CAPEX from own balance sheet")
        ppa_market_exists  = st.checkbox("A PPA market exists in this jurisdiction")
    with c2:
        permit_required    = st.checkbox("Generation / renewable energy licence required")
        permit_obtained    = st.checkbox("Licence obtained", disabled=not permit_required)
        aca_target         = st.checkbox("Airport has an ACA or equivalent carbon target")
        aca_level          = st.selectbox("ACA level (if applicable)", ["Not applicable", "Level 1", "Level 2",
                                                                         "Level 3", "Level 3+", "Level 4", "Level 4+", "Level 5"],
                                           disabled=not aca_target)

    model_notes = {
        "Direct ownership": "Airport carries full CAPEX. Best when balance sheet is strong and tax position allows depreciation benefits.",
        "Power Purchase": "Zero CAPEX for airport. Developer takes construction and operational risk. Airport pays for electricity at agreed rate.",
        "Special Purpose": "SPV carries CAPEX. Airport contributes land. Used when airport cannot fund capital — see KLIA / Cenergi precedent.",
        "Rooftop lease": "Simplest structure. Developer pays rent for roof space. Airport receives income but no direct carbon benefit in the energy balance.",
    }
    matched_note = next((v for k, v in model_notes.items() if k in commercial_model), "")

    g4_flags = []
    if "Direct" in commercial_model and not can_fund_capex:
        g4_flags.append(("warn", "Direct ownership selected but airport cannot fund CAPEX — consider SPV or PPA route"))
    if "PPA" in commercial_model and not ppa_market_exists:
        g4_flags.append(("fail", "PPA selected but no PPA market in this jurisdiction — developer cannot be remunerated"))
    if permit_required and not permit_obtained:
        g4_flags.append(("warn", "Generation licence required but not yet obtained — add 6–12 months to programme"))
    if not g4_flags:
        g4_flags.append(("pass", f"Commercial model selected: {commercial_model.split('(')[0].strip()}"))

    if matched_note:
        st.info(f"**Model note:** {matched_note}")

    st.markdown('<hr class="m2p-divider">', unsafe_allow_html=True)
    for status, msg in g4_flags:
        badge = {"pass": "badge-pass", "warn": "badge-warn", "fail": "badge-fail", "info": "badge-info"}[status]
        st.markdown(f'<span class="badge {badge}">{status.upper()}</span> {msg}', unsafe_allow_html=True)

# ── GATE 5 ─────────────────────────────────────────────────
with tabs[4]:
    st.markdown('<div class="section-label">Gate 5 — Greenfield vs Brownfield Checklist</div>', unsafe_allow_html=True)
    st.markdown(f"Additional checks for **{airport_type}** developments.")

    if airport_type == "Greenfield":
        st.markdown("##### Greenfield checks")
        checks_gf = {
            "Solar integrated into master plan (not added after layout decisions)":
                "Retrofitting solar after layout is fixed can reduce yield 15–20% due to suboptimal orientation.",
            "Terminal roof pitch and orientation optimised for solar at this stage":
                "This is an architectural decision. Once fixed it cannot be changed without redesign.",
            "Ground-mount land safeguarded before parcels are allocated to other uses":
                "Once allocated, land is gone. Safeguarding costs nothing at masterplan stage.",
            "Grid connection designed to include solar generation from day one":
                "Retrofit grid upgrade can cost 2–3× a designed-in connection.",
            "Approach paths and obstacle surfaces designed to leave solar zones clear":
                "Cheaper to plan around them than conduct glare studies on a fixed layout.",
        }
        results_gf = {}
        for check, detail in checks_gf.items():
            results_gf[check] = st.checkbox(check, help=detail)
        incomplete_gf = [c for c, v in results_gf.items() if not v]
        if incomplete_gf:
            st.warning(f"**{len(incomplete_gf)} check(s) outstanding** — resolve before concept design is frozen.")
        else:
            st.success("All greenfield checks passed.")

    else:  # Brownfield
        st.markdown("##### Brownfield checks")
        checks_bf = {
            "Structural survey of existing roofs confirms load-bearing capacity for panels":
                "Standard panels add 15–20 kg/m². Many older terminal roofs cannot take this.",
            "Roof remaining life matches or exceeds 25-year panel life":
                "Replacing the roof mid-life requires full removal and reinstallation of panels.",
            "Shadow study completed for buildings, stands and parked aircraft":
                "Aircraft shadows can reduce rooftop yield by 20–30% in some configurations.",
            "Glare study covers all existing approach paths and taxiways (including any new ones)":
                "New taxiways or stand reconfigurations since original design may have changed the picture.",
            "Confirmed whether a land bank exists that is not committed to other development":
                "The best brownfield solar sites are often surplus operational land, not rooftops.",
            "Grid connection upgrade cost and lead time obtained from DNO":
                "Lead time can exceed 24 months and can be the longest item on the programme.",
        }
        results_bf = {}
        for check, detail in checks_bf.items():
            results_bf[check] = st.checkbox(check, help=detail)
        incomplete_bf = [c for c, v in results_bf.items() if not v]
        if incomplete_bf:
            st.warning(f"**{len(incomplete_bf)} check(s) outstanding** — resolve before committing to a layout.")
        else:
            st.success("All brownfield checks passed.")

# ── GATE 6 ─────────────────────────────────────────────────
with tabs[5]:
    st.markdown('<div class="section-label">Gate 6 — Indicative Sizing & Financial Model</div>', unsafe_allow_html=True)
    st.markdown("Run this only once Gates 1–5 are substantially clear. All figures are indicative only — not for final investment decisions.")

    c1, c2, c3 = st.columns(3)
    with c1:
        install_type       = st.selectbox("Installation type", [
            "Ground-mount – fixed tilt",
            "Ground-mount – single-axis tracking",
            "Rooftop – flat",
            "Rooftop – pitched / optimised",
            "Mixed (ground-mount + rooftop)",
        ])
        ha_per_mw_lookup = {
            "Ground-mount – fixed tilt": 2.0,
            "Ground-mount – single-axis tracking": 2.5,
            "Rooftop – flat": 0.10,
            "Rooftop – pitched / optimised": 0.08,
            "Mixed (ground-mount + rooftop)": 1.8,
        }
        ha_per_mw          = ha_per_mw_lookup[install_type] * terrain_area_mult
        usable_fraction    = st.slider("Usable fraction after setbacks / shading (%)", 40, 90, 70, 5) / 100
    with c2:
        effective_ha       = gate2_cleared_ha * usable_fraction
        indicative_mw      = effective_ha / ha_per_mw
        capacity_override  = st.checkbox("Override with a target capacity (client-specified)")
        if capacity_override:
            indicative_mw  = st.number_input("Target capacity (MW)", 1.0, 1000.0,
                                              float(max(1.0, round(indicative_mw, 0))), 1.0)
    with c3:
        system_eff         = st.slider("System efficiency (%)", 70, 90, 80, 1) / 100
        degradation_yr     = st.slider("Annual degradation (%)", 0.3, 1.0, 0.5, 0.1) / 100

    # ── CALCULATION ENGINE ──────────────────────────────────
    capex_solar         = indicative_mw * capex_mw * terrain_capex_mult
    capex_bess          = bess_mwh * bess_capex_mwh if add_bess else 0.0
    capex_land          = land_cost_total if "acquired" in land_ownership else 0.0
    total_capex         = capex_solar + capex_bess + capex_land
    depreciable_base    = capex_solar + capex_bess       # land NOT depreciated
    equity_invest       = total_capex * (1 - debt_pct)
    debt_principal      = total_capex * debt_pct

    base_kwh            = indicative_mw * 1000 * peak_sun * system_eff * 365
    annual_co2_t        = base_kwh * 0.233 / 1000       # EU average emission factor kg/kWh

    if debt_pct > 0 and interest_pct > 0 and loan_years > 0:
        annual_debt_svc = (debt_principal * interest_pct) / (1 - (1 + interest_pct) ** -loan_years)
    else:
        annual_debt_svc = 0.0
    annual_dep          = depreciable_base / project_life
    outstanding_debt    = debt_principal
    cash_flows          = [-equity_invest]
    records             = []

    for yr in range(1, project_life + 1):
        yr_kwh      = base_kwh * ((1 - degradation_yr) ** (yr - 1))
        rev_solar   = yr_kwh * ppa_price * ((1 + ppa_escalation) ** (yr - 1)) * (self_consume_pct / 100)
        rev_bess    = (bess_mwh * bess_revenue_mwh_yr) if add_bess else 0.0
        revenue     = rev_solar + rev_bess
        opex        = (capex_solar * opex_pct) * ((1 + opex_inflation) ** (yr - 1))
        lifecycle   = (indicative_mw * 60000) if yr == 15 else 0.0   # inverter replacement, capitalised separately

        if yr <= loan_years:
            interest = outstanding_debt * interest_pct
            principal_pay = annual_debt_svc - interest
            debt_svc  = annual_debt_svc
            outstanding_debt = max(0, outstanding_debt - principal_pay)
        else:
            interest = 0.0; debt_svc = 0.0

        taxable     = revenue - opex - annual_dep - interest
        if yr == 15:
            taxable -= lifecycle * 0.03   # partial expense allowance
        tax         = max(0.0, taxable * tax_pct)
        net_cf      = revenue - opex - lifecycle - debt_svc - tax

        cash_flows.append(net_cf)
        records.append({"Year": yr, "Revenue": revenue, "OPEX": -opex,
                         "Debt Service": -debt_svc, "Tax": -tax,
                         "Net Cash Flow": net_cf})

    df = pd.DataFrame(records)
    df["Cumulative"] = df["Net Cash Flow"].cumsum() - equity_invest

    # IRR
    irr_val = npf.irr(cash_flows)
    irr_str = f"{irr_val:.1%}" if (irr_val is not None and not (irr_val != irr_val)) else "Not achievable at these inputs"

    pb_series = df[df["Cumulative"] > 0]["Year"]
    pb_str = f"Year {int(pb_series.min())}" if not pb_series.empty else "Beyond 25 years"

    # ── KPI TILES ───────────────────────────────────────────
    st.markdown('<hr class="m2p-divider">', unsafe_allow_html=True)
    st.markdown(f"### {project_name} — Indicative results")

    st.markdown(f"""
    <div class="kpi-row">
      <div class="kpi-tile blue">
        <div class="val">{indicative_mw:.1f} MW</div>
        <div class="lbl">Indicative capacity</div>
      </div>
      <div class="kpi-tile">
        <div class="val">{effective_ha:.1f} ha</div>
        <div class="lbl">Effective site area</div>
      </div>
      <div class="kpi-tile teal">
        <div class="val">{base_kwh/1e6:.1f} GWh</div>
        <div class="lbl">Year 1 generation</div>
      </div>
      <div class="kpi-tile green">
        <div class="val">{annual_co2_t:,.0f} t</div>
        <div class="lbl">CO₂ avoided / yr</div>
      </div>
      <div class="kpi-tile amber">
        <div class="val">€{total_capex/1e6:.1f}M</div>
        <div class="lbl">Total CAPEX</div>
      </div>
      <div class="kpi-tile">
        <div class="val">{irr_str}</div>
        <div class="lbl">After-tax levered IRR</div>
      </div>
      <div class="kpi-tile blue">
        <div class="val">{pb_str}</div>
        <div class="lbl">Equity payback</div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # ── CHARTS ──────────────────────────────────────────────
    col_chart1, col_chart2 = st.columns(2)

    with col_chart1:
        fig1 = go.Figure()
        fig1.add_bar(x=df["Year"], y=df["Net Cash Flow"]/1e6,
                     name="Annual net cash flow", marker_color=M2P_BLUE, opacity=0.75)
        fig1.add_scatter(x=df["Year"], y=df["Cumulative"]/1e6,
                         mode="lines+markers", name="Cumulative return",
                         line=dict(color=M2P_AMBER, width=2.5))
        fig1.add_hline(y=0, line_dash="dot", line_color="grey", opacity=0.5)
        fig1.update_layout(
            title="Cash flows over project life",
            xaxis_title="Year", yaxis_title="€ million",
            legend=dict(orientation="h", y=-0.2),
            plot_bgcolor="white", paper_bgcolor="white",
            font=dict(family="Century Gothic, Trebuchet MS, sans-serif", size=11),
            margin=dict(l=40, r=20, t=50, b=40),
        )
        st.plotly_chart(fig1, use_container_width=True)

    with col_chart2:
        capex_breakdown = {
            "Solar panels & BOS": capex_solar / 1e6,
            "Battery storage (BESS)": capex_bess / 1e6,
            "Land acquisition": capex_land / 1e6,
        }
        capex_breakdown = {k: v for k, v in capex_breakdown.items() if v > 0}
        fig2 = go.Figure(go.Pie(
            labels=list(capex_breakdown.keys()),
            values=list(capex_breakdown.values()),
            marker_colors=[M2P_BLUE, M2P_TEAL, M2P_AMBER],
            textinfo="label+percent",
            hole=0.42,
        ))
        fig2.update_layout(
            title="CAPEX breakdown",
            font=dict(family="Century Gothic, Trebuchet MS, sans-serif", size=11),
            paper_bgcolor="white",
            margin=dict(l=20, r=20, t=50, b=20),
        )
        st.plotly_chart(fig2, use_container_width=True)

    # ── DETAILED TABLE ──────────────────────────────────────
    with st.expander("View year-by-year cash flow table"):
        display_df = df.copy()
        for col in ["Revenue", "OPEX", "Debt Service", "Tax", "Net Cash Flow", "Cumulative"]:
            display_df[col] = display_df[col].apply(lambda x: f"€{x:,.0f}")
        st.dataframe(display_df, use_container_width=True, hide_index=True)

    # ── BENCHMARKS ──────────────────────────────────────────
    st.markdown('<hr class="m2p-divider">', unsafe_allow_html=True)
    st.markdown('<div class="section-label">Calibration against M2P case studies</div>', unsafe_allow_html=True)
    bench_data = {
        "Airport":      ["Munich (MUC)",         "KLIA (KUL)",                      "SFO"],
        "Country":      ["Germany",              "Malaysia",                         "USA"],
        "Capacity":     ["50 MWp (target 2030)", "30 MW (due 2027)",                "~5 MW"],
        "Annual yield": ["~45 GWh",              "~46 GWh",                          "~8 GWh"],
        "CO₂ avoided":  ["~18,000 t/yr",         "~35,000 t/yr",                    "~1,800 t/yr"],
        "Structure":    ["Direct ownership",     "SPV with Cenergi SEA",             "Grid procurement"],
        "Notes":        ["14% complete; 1,600 ha site; 2/3 protected",
                         "Airport land bank; no balance sheet cost",
                         "1.6% of 290 GWh demand; procurement dominant"],
    }
    st.dataframe(pd.DataFrame(bench_data), use_container_width=True, hide_index=True)

    # ── ASSUMPTION FLAGS ────────────────────────────────────
    st.markdown('<hr class="m2p-divider">', unsafe_allow_html=True)
    st.markdown('<div class="section-label">Model assumptions and limitations</div>', unsafe_allow_html=True)
    assumptions = [
        ("info", "Battery storage revenue is modelled as a user-input annual value per MWh. A full BESS dispatch model requires a separate analysis."),
        ("info", "CO₂ avoided uses the EU average grid emission factor (0.233 kg/kWh). Use a local factor for non-EU sites."),
        ("info", "Land cost is zero where airport owns the site. Update the acquisition figure at Gate 1 if purchasing land."),
        ("warn", "IRR and payback are pre-feasibility estimates. A full financial model for investment decisions requires audited demand data, a formal grid study, and tax advice in the relevant jurisdiction."),
        ("warn", "Glare and OLS assessments must be completed before any layout is committed. Uncompleted studies are flagged at Gate 2."),
    ]
    for status, msg in assumptions:
        badge = {"warn": "badge-warn", "info": "badge-info"}[status]
        st.markdown(f'<span class="badge {badge}">{status.upper()}</span> {msg}', unsafe_allow_html=True)

# ── EXPORT TAB ─────────────────────────────────────────────
with tabs[6]:
    st.markdown('<div class="section-label">Export</div>', unsafe_allow_html=True)
    st.markdown("Generate a summary for the client file or internal review.")

    summary = f"""
M2P CONSULTING – AIRPORT SOLAR FEASIBILITY SUMMARY
Project: {project_name}
Generated: 26 August 2026
Tool version: 2.0

SITE
  Airport type:        {airport_type}
  Region:              {airport_region}
  Total site area:     {total_site_ha:.0f} ha
  Gate 1 available:    {gate1_available_ha:.1f} ha
  Gate 2 cleared:      {gate2_cleared_ha:.1f} ha

SIZING
  Installation type:   {install_type}
  Indicative capacity: {indicative_mw:.1f} MW
  Effective area:      {effective_ha:.1f} ha
  Year 1 generation:   {base_kwh/1e6:.1f} GWh
  CO2 avoided/yr:      {annual_co2_t:,.0f} t

FINANCIALS (INDICATIVE ONLY)
  Total CAPEX:         €{total_capex:,.0f}
  Equity required:     €{equity_invest:,.0f}
  After-tax IRR:       {irr_str}
  Equity payback:      {pb_str}

COMMERCIAL MODEL
  {commercial_model}

DISCLAIMER
  All figures are indicative. Not for final investment decisions.
  Requires completed glare study, OLS mapping, DNO grid enquiry
  and jurisdiction-specific tax and legal advice before use.
"""
    st.download_button(
        label="Download summary (.txt)",
        data=summary,
        file_name=f"M2P_Solar_Feasibility_{project_name.replace(' ', '_')}.txt",
        mime="text/plain",
    )
    st.text_area("Preview", value=summary, height=400)

# ── FOOTER ─────────────────────────────────────────────────
st.markdown("""
<div class="disclaimer">
  <strong>Disclaimer:</strong> This tool produces pre-feasibility estimates for M2P internal use only.
  All outputs are indicative and require a completed glare study, OLS mapping, DNO grid connection enquiry,
  and jurisdiction-specific tax, legal and engineering advice before being presented to clients or used for
  investment decisions. Source: M2P Airport Solar Feasibility Tool v2.0, August 2026.
</div>
""", unsafe_allow_html=True)
