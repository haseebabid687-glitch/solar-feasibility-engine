import streamlit as st
import pandas as pd
import numpy_financial as npf
import plotly.express as px

# ---------------------------------------------------------
# 1. THE UPGRADED CALCULATION ENGINE
# ---------------------------------------------------------
def run_after_tax_scenario(capex_per_mw, ppa_price_kwh, capacity_mw, opex_per_mw, 
                           peak_sun_hours, system_efficiency, degradation_rate, 
                           inverter_cost_per_mw, acres_per_mw, land_cost_per_acre,
                           debt_ratio, interest_rate, loan_term_years, tax_rate, depreciation_years,
                           terrain_type, bess_mwh, bess_capex_per_mwh, bess_acres_per_mwh):
    
    # Apply Terrain Penalties
    terrain_capex_mult = 1.0
    terrain_land_mult = 1.0
    if terrain_type == "Gentle Slope (0-5%)":
        terrain_capex_mult = 1.05
        terrain_land_mult = 1.10
    elif terrain_type == "Complex / Steep (>5%)":
        terrain_capex_mult = 1.15
        terrain_land_mult = 1.25

    # Solar Hardware & Land
    solar_land_acres = capacity_mw * acres_per_mw * terrain_land_mult
    solar_capex = capacity_mw * capex_per_mw * terrain_capex_mult
    
    # Battery Storage (BESS) Hardware & Land
    bess_land_acres = bess_mwh * bess_acres_per_mwh
    bess_capex = bess_mwh * bess_capex_per_mwh
    
    # Totals
    total_gross_land = solar_land_acres + bess_land_acres
    total_land_cost = total_gross_land * land_cost_per_acre
    total_capex = solar_capex + bess_capex + total_land_cost
    
    debt_principal = total_capex * debt_ratio
    equity_investment = total_capex - debt_principal
    
    capacity_kw = capacity_mw * 1000
    base_annual_kwh = capacity_kw * peak_sun_hours * system_efficiency * 365
    
    if interest_rate > 0 and loan_term_years > 0:
        annual_debt_service = (debt_principal * interest_rate) / (1 - (1 + interest_rate)**-loan_term_years)
    else:
        annual_debt_service = 0
        
    annual_depreciation = total_capex / depreciation_years 
    
    cash_flows = [-equity_investment]
    cash_flow_records = [{
        "Year": 0, "Gross_Revenue": 0, "OPEX": 0, "Interest_Expense": 0, 
        "Tax_Paid": 0, "Debt_Service": 0, "After_Tax_Cash_Flow": -equity_investment
    }]
    
    outstanding_debt = debt_principal
    
    for year in range(1, 26):
        current_yield = base_annual_kwh * ((1 - degradation_rate) ** (year - 1))
        
        revenue = current_yield * ppa_price_kwh
        opex = capacity_mw * opex_per_mw
        lifecycle_capex = (capacity_mw * inverter_cost_per_mw) if year == 15 else 0
        
        if year <= loan_term_years:
            interest_expense = outstanding_debt * interest_rate
            principal_payment = annual_debt_service - interest_expense
            current_debt_service = annual_debt_service
            outstanding_debt -= principal_payment
        else:
            interest_expense = 0
            current_debt_service = 0
            
        taxable_income = revenue - opex - annual_depreciation - interest_expense - lifecycle_capex
        tax_paid = max(0, taxable_income * tax_rate)
        
        net_cf = revenue - opex - lifecycle_capex - current_debt_service - tax_paid
        cash_flows.append(net_cf)
        
        cash_flow_records.append({
            "Year": year, "Gross_Revenue": revenue, "OPEX": -opex, 
            "Interest_Expense": -interest_expense, "Tax_Paid": -tax_paid, 
            "Debt_Service": -current_debt_service, "After_Tax_Cash_Flow": net_cf
        })
        
    after_tax_irr = npf.irr(cash_flows)
    
    df_scenario = pd.DataFrame(cash_flow_records)
    df_scenario["Cumulative_Cash_Flow"] = df_scenario["After_Tax_Cash_Flow"].cumsum()
    
    return after_tax_irr, total_capex, equity_investment, total_gross_land, df_scenario

# ---------------------------------------------------------
# 2. THE EXECUTIVE UI WITH TERRAIN & BATTERIES
# ---------------------------------------------------------
st.set_page_config(page_title="Solar Feasibility Engine", layout="wide")

st.title("☀️ Comprehensive Solar & BESS Feasibility Engine")
st.markdown("Instantly model capital requirements and land footprints for complex infrastructure sites.")

# --- THE MASTER INPUTS ---
col_a, col_b, col_c = st.columns(3)
with col_a:
    capacity_mw = st.number_input("Target Solar Capacity (MW)", min_value=1, max_value=1000, value=100, step=10)
with col_b:
    terrain_type = st.selectbox("Site Topography", ["Flat / Optimal", "Gentle Slope (0-5%)", "Complex / Steep (>5%)"])
with col_c:
    add_battery = st.checkbox("Include Battery Storage (BESS)?")
    bess_mwh = st.number_input("Battery Capacity (MWh)", min_value=10, max_value=500, value=50, step=10) if add_battery else 0

# --- ADVANCED SETTINGS ---
with st.expander("⚙️ View & Edit Baseline Market Assumptions"):
    colA, colB, colC = st.columns(3)
    with colA:
        capex_per_mw = st.number_input("Base Solar CAPEX per MW (€)", value=800000, step=50000)
        ppa_price_kwh = st.number_input("Blended PPA Price (€/kWh)", value=0.08, format="%.3f")
        peak_sun_hours = st.slider("Peak Sun Hours", 1.5, 6.0, 2.9, 0.1)
    with colB:
        acres_per_mw = st.number_input("Solar Acres per MW", value=6.0)
        land_cost_per_acre = st.number_input("Land Cost per Acre (€)", value=15000)
        tax_rate = st.slider("Corporate Tax Rate (%)", 0.0, 0.40, 0.25)
    with colC:
        debt_ratio = st.slider("Debt Leverage (%)", 0.0, 1.0, 0.70, 0.05)
        interest_rate = st.number_input("Bank Interest Rate (%)", value=0.055, format="%.3f")
        loan_term_years = st.slider("Loan Term (Years)", 5, 25, 15)

# Fixed backend assumptions
opex_per_mw, system_efficiency, degradation_rate = 15000, 0.80, 0.005
inverter_cost_per_mw, depreciation_years = 60000, 25
bess_capex_per_mwh, bess_acres_per_mwh = 250000, 0.1

# --- RUN THE MODEL ---
irr, total_capex, equity_investment, total_gross_land, df_output = run_after_tax_scenario(
    capex_per_mw, ppa_price_kwh, capacity_mw, opex_per_mw, peak_sun_hours, 
    system_efficiency, degradation_rate, inverter_cost_per_mw, acres_per_mw, 
    land_cost_per_acre, debt_ratio, interest_rate, loan_term_years, tax_rate, depreciation_years,
    terrain_type, bess_mwh, bess_capex_per_mwh, bess_acres_per_mwh
)

# --- DYNAMIC PAYBACK CALCULATOR ---
payback_year_series = df_output[df_output['Cumulative_Cash_Flow'] > 0]['Year']
payback_text = f"Year {payback_year_series.min()}" if not payback_year_series.empty else "No Payback"

# --- EXECUTIVE READOUT (INSTANT KPIs) ---
st.divider()
st.subheader(f"Strategic Overview: {capacity_mw} MW Solar + {bess_mwh} MWh Storage")

col1, col2, col3, col4, col5 = st.columns(5)
col1.metric("Total Land Required", f"{total_gross_land:,.1f} Acres")
col2.metric("Total Project CAPEX", f"€{total_capex:,.0f}")
col3.metric("Required Cash Equity", f"€{equity_investment:,.0f}")
col4.metric("After-Tax Levered IRR", f"{irr:.2%}")
col5.metric("Breakeven Point", payback_text)

st.divider()

# --- CHARTING ---
fig = px.bar(df_output, x="Year", y="After_Tax_Cash_Flow", 
              title="Annual Net Cash Flows vs. Cumulative Return",
              labels={"After_Tax_Cash_Flow": "Net Cash (€)"},
              template="plotly_white")
fig.add_scatter(x=df_output["Year"], y=df_output["Cumulative_Cash_Flow"], 
                mode='lines+markers', name='Cumulative Return', 
                line=dict(color='orange', width=3))

st.plotly_chart(fig, use_container_width=True)
