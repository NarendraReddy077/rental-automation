import streamlit as st
import pandas as pd
import openpyxl
import datetime
import os
import io

import sheets.main_sheet as main_sheet
import sheets.rent_calculation as rent_calculation
import sheets.lease_size as lease_size
import sheets.capex_pm as capex_pm
import sheets.security_deposit as security_deposit

st.set_page_config(
    page_title="Siemens - Rental Automation System",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- TEMPLATE UTILITIES ---
def load_template(filename, directory="templates"):
    """Reads raw file contents from templates directory for dynamic HTML/CSS injection."""
    path = os.path.join(directory, filename)
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return f.read()
        except Exception:
            return ""
    return ""

# Inject css file from templates
style_css = load_template("style.css")
if style_css:
    st.markdown(f"<style>{style_css}</style>", unsafe_allow_html=True)


# --- DYNAMIC EXCEL PARSER ---
def parse_input_xlsx(file_bytes):
    """Parses user-uploaded input Excel sheet with extensive error resilience."""
    try:
        df = pd.read_excel(file_bytes, sheet_name="Main")
        df.columns = [c.strip() for c in df.columns]
        
        # Build dictionary from Name-Value pairs
        raw_params = {}
        for idx, row in df.iterrows():
            if 'Field Name' in df.columns and 'Sample Value' in df.columns:
                name = str(row['Field Name']).strip()
                val = row['Sample Value']
                if pd.isna(val):
                    val = None
                raw_params[name] = val
        
        # Helper to parse dates
        def parse_date(date_val, default=None):
            if date_val is None:
                return default
            if isinstance(date_val, (datetime.date, datetime.datetime)):
                return date_val.date() if isinstance(date_val, datetime.datetime) else date_val
            try:
                return pd.to_datetime(str(date_val).strip()).date()
            except Exception:
                return default

        # Normalize values
        params = {}
        params["Lease ID"] = raw_params.get("Lease ID", "LEASE-001")
        params["REU Name"] = raw_params.get("REU Name", "IN-CHEK2")
        params["Lease Type"] = raw_params.get("Lease Type", "Office")
        params["Building Name"] = raw_params.get("Building Name", "SKCL Tech Park")
        params["City"] = raw_params.get("City", "Chennai")
        params["Country"] = raw_params.get("Country", "India")
        params["Currency"] = raw_params.get("Currency", "INR")
        
        params["Chargeable Area Sqft"] = float(raw_params.get("Chargeable Area Sqft", 9515))
        params["Parking Slots"] = int(raw_params.get("Parking Slots", 20))
        
        start_date = parse_date(raw_params.get("Agreement Start Date"), datetime.date(2026, 4, 1))
        end_date = parse_date(raw_params.get("Agreement End Date"), datetime.date(2031, 3, 31))
        params["Agreement Start Date"] = start_date
        params["Agreement End Date"] = end_date
        params["Rent Start Date"] = parse_date(raw_params.get("Rent Start Date"), start_date)
        
        # Lease Term calculations
        term_months = raw_params.get("Lease Term Months")
        if term_months is None or pd.isna(term_months):
            # Calculate months from start/end dates
            term_months = round((end_date - start_date).days / 30.4167)
        params["Lease Term Months"] = int(term_months)
        
        params["Rent Per Sqft"] = float(raw_params.get("Rent Per Sqft", 120))
        params["Quoted CAM"] = float(raw_params.get("Quoted CAM", 15.48))
        
        esc_val = raw_params.get("Escalation %", 0.15)
        params["Escalation %"] = float(esc_val) if esc_val is not None else 0.15
        
        # Robust check for frequency e.g. 0.36 -> 36 months
        freq_val = raw_params.get("Escalation Frequency Months", 36)
        if freq_val is not None:
            freq_val = float(freq_val)
            if freq_val < 1.0:
                freq_val = int(round(freq_val * 100))
            else:
                freq_val = int(freq_val)
        else:
            freq_val = 36
        params["Escalation Frequency Months"] = freq_val
        
        params["CAM Escalation %"] = 0.05
        params["Billing Frequency"] = raw_params.get("Billing Frequency", "Monthly")
        params["Security Deposit Months"] = float(raw_params.get("Security Deposit Months", 6))
        params["Security Deposit Amount"] = float(raw_params.get("Security Deposit Amount", 11418000))
        params["Refundable Deposit"] = raw_params.get("Refundable Deposit", "Yes")
        
        # Capex & Cost factors
        params["Fitout Cost"] = float(raw_params.get("Fitout Cost", 34000000.0))
        params["Useful Life Years"] = float(raw_params.get("Useful Life Years", 5))
        params["Residual Value"] = float(raw_params.get("Residual Value", 0))
        params["Discount Rate"] = float(raw_params.get("Discount Rate", 0.08))
        params["Incremental Borrowing Rate"] = float(raw_params.get("Incremental Borrowing Rate", 0.08))
        params["Cost of Capital"] = float(raw_params.get("Cost of Capital", 0.105))
        params["Addnl.Deposit -energy(Refundable)"] = float(raw_params.get("Addnl.Deposit -energy(Refundable)", 500000))
        
        # Add new parameters missing in original but parsed if they exist
        params["Imputed Interest Rate"] = float(raw_params.get("Imputed Interest Rate", 0.0711))
        params["Ready Reckoner Rate"] = float(raw_params.get("Ready Reckoner Rate", 15000.0))
        params["Exchange Rate"] = float(raw_params.get("Exchange Rate", 105.02))
        params["Incremental Restoration Cost Sqft"] = float(raw_params.get("Incremental Restoration Cost Sqft", 82.6))
        params["PM Cost Over Lease"] = float(raw_params.get("PM Cost Over Lease", 2500000.0))
        
        return params, None
    except Exception as e:
        return {}, str(e)


def compile_output_workbook(template_path, params):
    """Loads specimen template, populates the 5 sheets sequentially, and saves new workbook."""
    wb = openpyxl.load_workbook(template_path, data_only=False)
    
    # Populate Main sheet
    main_sheet.inject(wb["Main"], params)
    
    # Populate Rent Calculation sheet
    rent_calculation.inject(wb["Rent Calculation"], params)
    
    # Populate Lease Size sheet
    lease_size.inject(wb["Lease Size"], params)
    
    # Populate CAPEX and PM sheet
    capex_pm.inject(wb["CAPEX and PM"], params)
    
    # Populate Security Deposit sheet
    security_deposit.inject(wb["Security Deposit"], params)
    
    for ws in wb.worksheets:
        if ws.sheet_properties:
            ws.sheet_properties.tabColor = None
            
    # Save to a memory stream for download
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# --- STREAMLIT UI LAYOUT & CONTROLS ---

# 1. Render custom premium header
header_html = load_template("header.html")
if header_html:
    st.markdown(header_html, unsafe_allow_html=True)
else:
    st.title("🏢 Rental Automation System")

# 2. Sidebar Ingestion Panel
st.sidebar.markdown('<p style="font-size: 1.3rem; font-weight:700; color:#f8fafc; margin-bottom: 15px;">📁 Data Ingestion</p>', unsafe_allow_html=True)
uploaded_file = st.sidebar.file_uploader("Upload Lease Parameter Workbook (.xlsx)", type=["xlsx"])

# Default fallback values matching Specimen
default_params = {
    "REU Name": "IN-CHEK2",
    "Building Name": "SKCL Tech Park",
    "City": "Chennai",
    "Country": "India",
    "Currency": "INR",
    "Lease Type": "Office",
    "Chargeable Area Sqft": 9515.0,
    "Agreement Start Date": datetime.date(2026, 4, 1),
    "Agreement End Date": datetime.date(2031, 3, 31),
    "Rent Per Sqft": 120.0,
    "Quoted CAM": 15.48,
    "Escalation %": 0.15,
    "Escalation Frequency Months": 36,
    "CAM Escalation %": 0.05,
    "Security Deposit Amount": 11418000.0,
    "Addnl.Deposit -energy(Refundable)": 500000.0,
    "Fitout Cost": 34000000.0,
    "Fitout Cost Breakdown": [14000000.0, 5000000.0, 15000000.0],
    "Capex Schedule": {2026: 14000000.0, 2027: 12000000.0, 2028: 5000000.0, 2029: 6000000.0, 2030: 6000000.0},
    "PM Cost Over Lease": 2500000.0,
    "PM Schedule": {2026: 1500000.0, 2027: 200000.0, 2028: 200000.0, 2029: 200000.0, 2030: 200000.0, 2031: 200000.0},
    "Cost of Capital": 0.105,
    "Discount Rate": 0.08,
    "Incremental Borrowing Rate": 0.08,
    "Imputed Interest Rate": 0.0711,
    "Ready Reckoner Rate": 15000.0,
    "Exchange Rate": 105.02,
    "Incremental Restoration Cost Sqft": 82.6,
    "Lease Term Months": 60
}

params = default_params
parse_error_msg = None
upload_occurred = False

if uploaded_file is not None:
    upload_occurred = True
    parsed_vals, err = parse_input_xlsx(uploaded_file)
    if err:
        parse_error_msg = err
    else:
        params = parsed_vals

# 3. Sidebar Override & Configuration Panel
st.sidebar.markdown('<p style="font-size: 1.25rem; font-weight:600; color:#cbd5e1; margin-top: 20px;">⚙️ Parameter Configuration</p>', unsafe_allow_html=True)

# Expander 1: Real Estate & Lease Properties
with st.sidebar.expander("▶ Real Estate & Lease Properties", expanded=True):
    reu_name = st.text_input("REU Name", value=params["REU Name"])
    bldg_name = st.text_input("Building Name", value=params["Building Name"])
    city = st.text_input("City", value=params["City"])
    area = st.number_input("Chargeable Area (Sq.ft.)", value=int(params["Chargeable Area Sqft"]), step=100)
    
    col_d1, col_d2 = st.columns(2)
    with col_d1:
        start_date = st.date_input("Start Date", value=params["Agreement Start Date"])
    with col_d2:
        end_date = st.date_input("End Date", value=params["Agreement End Date"])
        
    rent_sqft = st.number_input("Quoted Rentals (per Sqft/mo)", value=float(params["Rent Per Sqft"]), step=1.0)
    cam_sqft = st.number_input("Quoted CAM (per Sqft/mo)", value=float(params["Quoted CAM"]), step=0.1)
    
    rent_esc = st.slider("Rent Escalation %", min_value=0.0, max_value=0.5, value=float(params["Escalation %"]), step=0.01)
    rent_esc_freq = st.number_input("Rent Escalation Freq (Months)", value=int(params["Escalation Frequency Months"]), step=12)
    cam_esc = st.slider("CAM Escalation %", min_value=0.0, max_value=0.2, value=float(params["CAM Escalation %"]), step=0.01)
    
    sec_deposit = st.number_input("Security Deposit Amount", value=float(params["Security Deposit Amount"]), step=50000.0)
    energy_dep = st.number_input("Energy Deposit Amount", value=float(params["Addnl.Deposit -energy(Refundable)"]), step=10000.0)

# Expander 2: CAPEX & PM Investment Schedule
with st.sidebar.expander("▶ CAPEX & PM Schedule"):
    st.markdown("**Fitout Investments (Breakdown)**")
    fitout_breakdown = params.get("Fitout Cost Breakdown", [14000000.0, 5000000.0, 15000000.0])
    fitout_1 = st.number_input("Fitout Phase 1 Cost", value=float(fitout_breakdown[0]), step=100000.0)
    fitout_2 = st.number_input("Fitout Phase 2 Cost", value=float(fitout_breakdown[1]), step=100000.0)
    fitout_3 = st.number_input("Fitout Phase 3 Cost", value=float(fitout_breakdown[2]), step=100000.0)
    
    fitout_total = fitout_1 + fitout_2 + fitout_3
    st.info(f"Total Fitout Cost: {params['Currency']} {fitout_total:,.2f}")
    
    st.markdown("**Capex Investments by Year**")
    capex_dict = params.get("Capex Schedule", {2026: 14000000.0, 2027: 12000000.0, 2028: 5000000.0, 2029: 6000000.0, 2030: 6000000.0})
    cap_26 = st.number_input("Capex FY2026", value=float(capex_dict.get(2026, 0)), step=100000.0)
    cap_27 = st.number_input("Capex FY2027", value=float(capex_dict.get(2027, 0)), step=100000.0)
    cap_28 = st.number_input("Capex FY2028", value=float(capex_dict.get(2028, 0)), step=100000.0)
    cap_29 = st.number_input("Capex FY2029", value=float(capex_dict.get(2029, 0)), step=100000.0)
    cap_30 = st.number_input("Capex FY2030", value=float(capex_dict.get(2030, 0)), step=100000.0)
    
    st.markdown("**Maintenance (PM) Schedule**")
    pm_cost_total = st.number_input("PM Total Cost", value=float(params["PM Cost Over Lease"]), step=100000.0)
    pm_dict = params.get("PM Schedule", {2026: 1500000.0, 2027: 200000.0, 2028: 200000.0, 2029: 200000.0, 2030: 200000.0, 2031: 200000.0})
    pm_26 = st.number_input("PM FY2026", value=float(pm_dict.get(2026, 0)), step=50000.0)
    pm_27 = st.number_input("PM FY2027", value=float(pm_dict.get(2027, 0)), step=10000.0)
    pm_28 = st.number_input("PM FY2028", value=float(pm_dict.get(2028, 0)), step=10000.0)
    pm_29 = st.number_input("PM FY2029", value=float(pm_dict.get(2029, 0)), step=10000.0)
    pm_30 = st.number_input("PM FY2030", value=float(pm_dict.get(2030, 0)), step=10000.0)
    pm_31 = st.number_input("PM FY2031", value=float(pm_dict.get(2031, 0)), step=10000.0)

# Expander 3: Financial Rates & ARO Workings
with st.sidebar.expander("▶ Rates & ARO Restoration"):
    wacc = st.slider("Cost of Capital (WACC) %", min_value=0.0, max_value=0.25, value=float(params["Cost of Capital"]), step=0.005)
    borrow_rate = st.slider("Incremental Borrowing Rate %", min_value=0.0, max_value=0.25, value=float(params["Incremental Borrowing Rate"]), step=0.005)
    imputed_rate = st.slider("Imputed Interest Rate %", min_value=0.0, max_value=0.20, value=float(params["Imputed Interest Rate"]), step=0.001)
    reckoner_rate = st.number_input("Ready Reckoner Rate (INR/sqm)", value=float(params["Ready Reckoner Rate"]), step=1000.0)
    aro_rate = st.number_input("Restoration Cost per Sqft (ARO)", value=float(params["Incremental Restoration Cost Sqft"]), step=5.0)
    exchange_rate = st.number_input("Forex Rate (INR/Euro)", value=float(params["Exchange Rate"]), step=0.1)

# Pack overrides into active UI parameters
ui_params = {
    "REU Name": reu_name,
    "Building Name": bldg_name,
    "City": city,
    "Country": params["Country"],
    "Currency": params["Currency"],
    "Lease Type": params["Lease Type"],
    "Chargeable Area Sqft": float(area),
    "Agreement Start Date": start_date,
    "Agreement End Date": end_date,
    "Rent Per Sqft": float(rent_sqft),
    "Quoted CAM": float(cam_sqft),
    "Escalation %": float(rent_esc),
    "Escalation Frequency Months": int(rent_esc_freq),
    "CAM Escalation %": float(cam_esc),
    "Security Deposit Amount": float(sec_deposit),
    "Addnl.Deposit -energy(Refundable)": float(energy_dep),
    
    "Fitout Cost": float(fitout_total),
    "Fitout Cost Breakdown": [fitout_1, fitout_2, fitout_3],
    "Capex Schedule": {2026: cap_26, 2027: cap_27, 2028: cap_28, 2029: cap_29, 2030: cap_30},
    "PM Cost Over Lease": float(pm_cost_total),
    "PM Schedule": {2026: pm_26, 2027: pm_27, 2028: pm_28, 2029: pm_29, 2030: pm_30, 2031: pm_31},
    
    "Cost of Capital": float(wacc),
    "Discount Rate": float(borrow_rate),
    "Incremental Borrowing Rate": float(borrow_rate),
    "Imputed Interest Rate": float(imputed_rate),
    "Ready Reckoner Rate": float(reckoner_rate),
    "Exchange Rate": float(exchange_rate),
    "Incremental Restoration Cost Sqft": float(aro_rate),
    "Lease Term Months": round((end_date - start_date).days / 30.4167)
}


# --- LIVE SIMULATION MODEL W wake-up ---
capex_pm_df = capex_pm.simulate(ui_params)
sd_results = security_deposit.simulate(ui_params)
rent_calc_results = rent_calculation.simulate(ui_params)
lease_size_df, npv_value = lease_size.simulate(ui_params, capex_pm_df)



# --- DYNAMIC STATUS CARD SECTION ---
if parse_error_msg:
    error_tpl = load_template("error_card.html")
    if error_tpl:
        st.markdown(error_tpl.format(error_message=f"Failed to parse sheet: {parse_error_msg}"), unsafe_allow_html=True)
    else:
        st.error(f"Failed to parse sheet: {parse_error_msg}")
elif upload_occurred:
    success_tpl = load_template("success_card.html")
    if success_tpl:
        total_sd = ui_params["Security Deposit Amount"] + ui_params["Addnl.Deposit -energy(Refundable)"]
        duration_yrs = ui_params["Lease Term Months"] / 12.0
        st.markdown(success_tpl.format(
            reu_name=ui_params["REU Name"],
            area=f"{int(ui_params['Chargeable Area Sqft']):,}",
            duration=f"{duration_yrs:.2f}",
            currency=ui_params["Currency"],
            total_deposit=f"{total_sd:,.2f}"
        ), unsafe_allow_html=True)
    else:
        st.success(f"Workbook parameters for {ui_params['REU Name']} compiled successfully!")
else:
    info_tpl = load_template("info_card.html")
    if info_tpl:
        st.markdown(info_tpl, unsafe_allow_html=True)


# --- LIVE KPI DASHBOARD (HTML GRID) ---
total_rent_cam = ui_params["Rent Per Sqft"] + ui_params["Quoted CAM"]
duration_yrs = ui_params["Lease Term Months"] / 12.0
total_sd_amount = ui_params["Security Deposit Amount"] + ui_params["Addnl.Deposit -energy(Refundable)"]

metrics_html = f"""
<div class="stats-grid">
    <div class="stat-cell">
        <span class="stat-label">Chargeable Area</span>
        <span class="stat-value">{int(ui_params['Chargeable Area Sqft']):,} sqft</span>
        <span class="stat-meta">≈ {ui_params['Chargeable Area Sqft']/10.764:,.2f} sq.m.</span>
    </div>
    <div class="stat-cell">
        <span class="stat-label">Initial Rent + CAM Rate</span>
        <span class="stat-value">{ui_params['Currency']} {total_rent_cam:.2f}</span>
        <span class="stat-meta">Rent: {ui_params['Rent Per Sqft']:.1f} | CAM: {ui_params['Quoted CAM']:.2f}</span>
    </div>
    <div class="stat-cell">
        <span class="stat-label">Project Net Outflow (NPV)</span>
        <span class="stat-value">Euro {npv_value/1000000:,.2f} M</span>
        <span class="stat-meta">WACC Rate: {ui_params['Cost of Capital']*100:.2f}%</span>
    </div>
    <div class="stat-cell">
        <span class="stat-label">Lease Duration</span>
        <span class="stat-value">{duration_yrs:.2f} yrs</span>
        <span class="stat-meta">{ui_params['Agreement Start Date'].strftime('%b %d, %Y')} to {ui_params['Agreement End Date'].strftime('%b %d, %Y')}</span>
    </div>
</div>
"""
st.markdown(metrics_html, unsafe_allow_html=True)


# --- EXCEL WORKBOOK COMPILER GENERATION ---
template_file_path = r"C:\Users\z0050s8t\Documents\rental-automaton\artifacts\Rental Specimen.xlsx"

if os.path.exists(template_file_path):
    try:
        output_stream = compile_output_workbook(template_file_path, ui_params)
        
        st.download_button(
            label="💾 Generate and Download Excel Workbook",
            data=output_stream,
            file_name=f"rental_workbook_{ui_params['REU Name']}_{datetime.date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    except Exception as e:
        st.error(f"Excel template compilation failed: {e}")
else:
    st.warning(f"Excel template specimen not found in path: {template_file_path}")

