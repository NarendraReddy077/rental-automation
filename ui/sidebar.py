import streamlit as st
import datetime
import os
from sheets.input_parser import parse_input_xlsx
from sheets.capex_pm import calculate_fy_months
from core.parameters import blank_params

def render_sidebar():
    # 2. Sidebar Ingestion Panel
    st.sidebar.markdown('<p style="font-size: 1.3rem; font-weight:700; color:#f8fafc; margin-bottom: 15px;">📁 Data Ingestion</p>', unsafe_allow_html=True)
    uploaded_file = st.sidebar.file_uploader("Upload Lease Parameter Workbook (.xlsx)", type=["xlsx"], help="Upload an Excel sheet (.xlsx) containing main lease configuration parameters.")

    # Input template is located in artifacts/input_template.xlsx relative to workspace root
    input_template_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "artifacts", "input_template.xlsx")
    if os.path.exists(input_template_path):
        try:
            with open(input_template_path, "rb") as f:
                template_bytes = f.read()
            st.sidebar.download_button(
                label="📥 Download Input Template",
                data=template_bytes,
                file_name="input_template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                help="Download the blank template sheet to fill in lease parameters."
            )
        except Exception as e:
            st.sidebar.error(f"Error loading template: {e}")

    params = blank_params
    parse_error_msg = None
    upload_occurred = False

    # Ensure upload_counter is in session_state
    if "upload_counter" not in st.session_state:
        st.session_state["upload_counter"] = 0

    if uploaded_file is not None:
        upload_occurred = True
        file_name = uploaded_file.name
        # If the file uploaded is different from the last one, increment counter to reset keys
        if st.session_state.get("last_uploaded_file") != file_name:
            st.session_state["last_uploaded_file"] = file_name
            st.session_state["upload_counter"] += 1
            
        parsed_vals, err = parse_input_xlsx(uploaded_file)
        if err:
            parse_error_msg = err
        else:
            params = parsed_vals
    else:
        # If file was cleared, increment counter to reset back to default params
        if "last_uploaded_file" in st.session_state:
            del st.session_state["last_uploaded_file"]
            st.session_state["upload_counter"] += 1

    key_suffix = f"_up_{st.session_state['upload_counter']}"

    # 3. Sidebar Override & Configuration Panel
    st.sidebar.markdown('<p style="font-size: 1.25rem; font-weight:600; color:#cbd5e1; margin-top: 20px;">⚙️ Parameter Configuration</p>', unsafe_allow_html=True)

    # Expander 1: Real Estate & Lease Properties
    with st.sidebar.expander("▶ Real Estate & Lease Properties", expanded=True):
        reu_name = st.text_input("REU Name", value=params["REU Name"], key=f"reu_name{key_suffix}", help="Unique identifier code for the Real Estate Unit.")
        bldg_name = st.text_input("Building Name", value=params["Building Name"], key=f"bldg_name{key_suffix}", help="The name of the building/facility where the premises are leased.")
        city = st.text_input("City", value=params["City"], key=f"city{key_suffix}", help="The city location of the leased property.")
        area = st.number_input("Chargeable Area (sq ft)", value=int(params["Chargeable Area Sqft"]) if params["Chargeable Area Sqft"] is not None else None, key=f"area{key_suffix}", step=100, help="The chargeable/rentable area in square feet.")
        
        col_d1, col_d2 = st.columns(2)
        with col_d1:
            start_date = st.date_input("Start Date", value=params["Agreement Start Date"], key=f"start_date{key_suffix}", help="The official lease agreement commencement date.")
        with col_d2:
            end_date = st.date_input("End Date", value=params["Agreement End Date"], key=f"end_date{key_suffix}", help="The official lease agreement expiration date.")
            
        rent_sqft = st.number_input("Quoted Rent (per sq ft/month)", value=float(params["Rent Per Sqft"]) if params["Rent Per Sqft"] is not None else None, key=f"rent_sqft{key_suffix}", step=1.0, help="The starting rental rate per square foot per month.")
        cam_sqft = st.number_input("Quoted CAM (per sq ft/month)", value=float(params["Quoted CAM"]) if params["Quoted CAM"] is not None else None, key=f"cam_sqft{key_suffix}", step=0.1, help="The starting Common Area Maintenance (CAM) rate per square foot per month.")
        
        rent_esc = st.slider("Rent Escalation %", min_value=0.0, max_value=0.5, value=float(params["Escalation %"]) if params["Escalation %"] is not None else 0.0, key=f"rent_esc{key_suffix}", step=0.01, help="The rental rate escalation percentage (e.g. 0.15 = 15%).")
        rent_esc_freq = st.number_input("Rent Escalation Freq (Months)", value=int(params["Escalation Frequency Months"]) if params["Escalation Frequency Months"] is not None else None, key=f"rent_esc_freq{key_suffix}", step=12, help="Interval in months between rent escalations (typically 36 months).")
        cam_esc = st.slider("CAM Escalation %", min_value=0.0, max_value=0.2, value=float(params["CAM Escalation %"]) if params["CAM Escalation %"] is not None else 0.0, key=f"cam_esc{key_suffix}", step=0.01, help="Escalation percentage applied to CAM rates.")
        cam_esc_freq = st.number_input("CAM Escalation Freq (Months)", value=int(params.get("CAM Escalation Frequency Months")) if params.get("CAM Escalation Frequency Months") is not None else None, key=f"cam_esc_freq{key_suffix}", step=12, help="Interval in months between CAM rate escalations (typically 12 months).")
        
        sec_deposit = st.number_input("Security Deposit Amount (INR)", value=int(params["Security Deposit Amount"]) if params["Security Deposit Amount"] is not None else None, key=f"sec_deposit{key_suffix}", step=50000, help="Interest-free refundable security deposit amount.")
        energy_dep = st.number_input("Energy Deposit Amount (INR)", value=int(params["Addnl.Deposit -energy(Refundable)"]) if params["Addnl.Deposit -energy(Refundable)"] is not None else None, key=f"energy_dep{key_suffix}", step=10000, help="Additional refundable security deposit specifically for power/utilities.")

    # Expander 1.5: Parking Configuration
    with st.sidebar.expander("▶ Parking Configuration"):
        col_p1, col_p2 = st.columns(2)
        with col_p1:
            four_w_slots = st.number_input("4 Wheeler Slots", value=int(params.get("4 Wheeler Slots")) if params.get("4 Wheeler Slots") is not None else None, key=f"four_w_slots{key_suffix}", step=1, help="Number of chargeable 4-wheeler parking slots.")
        with col_p2:
            four_w_rate = st.number_input("4 Wheeler Rate (per mo)", value=float(params.get("4 Wheeler Rate")) if params.get("4 Wheeler Rate") is not None else None, key=f"four_w_rate{key_suffix}", step=100.0, help="Monthly rate per 4-wheeler slot.")
        col_p3, col_p4 = st.columns(2)
        with col_p3:
            two_w_slots = st.number_input("2 Wheeler Slots", value=int(params.get("2 Wheeler Slots")) if params.get("2 Wheeler Slots") is not None else None, key=f"two_w_slots{key_suffix}", step=1, help="Number of chargeable 2-wheeler parking slots.")
        with col_p4:
            two_w_rate = st.number_input("2 Wheeler Rate (per mo)", value=float(params.get("2 Wheeler Rate")) if params.get("2 Wheeler Rate") is not None else None, key=f"two_w_rate{key_suffix}", step=100.0, help="Monthly rate per 2-wheeler slot.")
            
        col_p5, col_p6 = st.columns(2)
        with col_p5:
            parking_esc = st.slider("Parking Escalation %", min_value=0.0, max_value=0.5, value=float(params.get("Parking Escalation %")) if params.get("Parking Escalation %") is not None else 0.0, key=f"parking_esc{key_suffix}", step=0.01, help="The parking rate escalation percentage.")
        with col_p6:
            parking_esc_freq = st.number_input("Parking Escalation Freq (Months)", value=int(params.get("Parking Escalation Frequency Months")) if params.get("Parking Escalation Frequency Months") is not None else None, key=f"parking_esc_freq{key_suffix}", step=1, help="Interval in months between parking rate escalations.")

    # Expander 1.6: OpEx Configuration
    with st.sidebar.expander("▶ OpEx Configuration"):
        opex_others = st.number_input("Opex I (Others per month)", value=float(params.get("Opex Others Per Month")) if params.get("Opex Others Per Month") is not None else None, key=f"opex_others{key_suffix}", step=10.0, help="Opex Others monthly amount.")
        opex_ii = st.number_input("Opex II (per month)", value=float(params.get("Opex II Per Month")) if params.get("Opex II Per Month") is not None else None, key=f"opex_ii{key_suffix}", step=10.0, help="Opex II monthly amount (OpEx Add-on).")

    # Expander 2: CAPEX & PM Investment Schedule
    with st.sidebar.expander("▶ CAPEX & PM Schedule"):
        st.markdown("**Fitout Investments (Breakdown)**")
        fitout_breakdown = params.get("Fitout Cost Breakdown", [])
        fitout_lifes = params.get("Fitout Useful Lives", [])
        
        # Filter out zeros from breakdown to determine configured phases
        non_zero_fitouts = [f for f in fitout_breakdown if f > 0]
        default_num_fitouts = max(1, len(non_zero_fitouts))
        max_fitouts = max(8, default_num_fitouts)
        num_fitouts = st.selectbox("Number of Fitout Phases", list(range(1, max_fitouts + 1)), index=default_num_fitouts - 1, key=f"num_fitouts{key_suffix}", help="Select the number of dynamic fitout phases.")
        
        fitout_costs = []
        fitout_useful_lives = []
        fitout_active_months_breakdown = []
        
        # Calculate years of the lease
        lease_years = list(range(start_date.year, end_date.year + 1)) if (start_date and end_date) else []
        
        term_months = round((end_date - start_date).days / 30.4167) if (start_date and end_date) else 0
        parsed_active_months_list = params.get("Fitout Active Months Breakdown", [])
        
        for i in range(num_fitouts):
            col1, col2 = st.columns(2)
            default_c = fitout_breakdown[i] if i < len(fitout_breakdown) else None
            
            if i == 0:
                default_l = term_months
                life_key = f"fitout_life_{i}_{term_months}{key_suffix}"
            else:
                default_l = fitout_lifes[i] if i < len(fitout_lifes) else None
                life_key = f"fitout_life_{i}{key_suffix}"
                
            with col1:
                c = st.number_input(f"Fitout Phase {i+1} Cost", value=int(default_c) if default_c is not None else None, key=f"fitout_cost_{i}{key_suffix}", step=100000, help=f"Capital expenditure for phase {i+1} of fitout works.")
            with col2:
                l = st.number_input(f"Phase {i+1} Useful Life (mo)", value=int(default_l) if default_l is not None else None, key=life_key, step=6, help=f"Amortization period for phase {i+1} in months.")
                
            fitout_costs.append(float(c) if c is not None else 0.0)
            fitout_useful_lives.append(int(l) if l is not None else 0)
            
            # Expandable year-by-year active months customization
            with st.expander(f"Phase {i+1} Year-by-Year Months Active"):
                default_m_dist = calculate_fy_months(start_date, term_months, int(l)) if (start_date and term_months and l is not None) else {}
                phase_m_dict = parsed_active_months_list[i] if i < len(parsed_active_months_list) else {}
                
                # Check if the user has overridden spreadsheet defaults for useful life or dates
                is_override = (
                    (default_l is not None and l != default_l) or
                    (start_date != params.get("Agreement Start Date")) or
                    (end_date != params.get("Agreement End Date"))
                )
                
                phase_m_custom = {}
                for yr in lease_years:
                    if is_override:
                        default_m = default_m_dist.get(yr, 0)
                    else:
                        default_m = phase_m_dict.get(yr, default_m_dist.get(yr, 0))
                        
                    # Include l, start_date, and term_months in the key to recreate the widget when dependencies are altered
                    key = f"fitout_{i}_m_{yr}_{l}_{start_date}_{term_months}{key_suffix}"
                    m = st.number_input(
                        f"FY{yr} Active Months",
                        min_value=0,
                        max_value=12,
                        value=int(default_m),
                        key=key
                    )
                    phase_m_custom[yr] = int(m)
                fitout_active_months_breakdown.append(phase_m_custom)
            
        fitout_total = sum(fitout_costs)
        st.info(f"Total Fitout Cost: {params.get('Currency', 'INR')} {fitout_total:,.2f}")
        
        st.markdown("**Capex Investments by Year**")
        capex_dict = params.get("Capex Schedule", {})
        # Determine non-zero elements
        non_zero_capex_years = [y for y, val in capex_dict.items() if val > 0]
        start_year = start_date.year if start_date else datetime.date.today().year
        
        # Determine default number of years based on max year configured
        if non_zero_capex_years:
            default_num_capex = max(1, max(non_zero_capex_years) - start_year + 1)
        else:
            default_num_capex = 5
            
        max_capex = max(8, default_num_capex)
        num_capex = st.selectbox("Number of Capex Years", list(range(1, max_capex + 1)), index=default_num_capex - 1, key=f"num_capex{key_suffix}", help="Select the number of years for CAPEX injection.")
        
        capex_schedule = {}
        capex_useful_lives = {}
        capex_lives_dict = params.get("Capex Useful Lives", {})
        
        for i in range(num_capex):
            yr = start_year + i
            default_val = capex_dict.get(yr, None)
            
            # Calculate default life (remaining lease term)
            if i == 0:
                default_life = term_months if term_months > 0 else None
            else:
                first_year_months = 12 - start_date.month + 1 if start_date else 12
                elapsed = first_year_months + 12 * (i - 1)
                default_life = max(0, term_months - elapsed) if term_months > 0 else None
                
            default_l = capex_lives_dict.get(yr, default_life)
            
            col1, col2 = st.columns(2)
            with col1:
                val = st.number_input(f"Capex FY{yr} Cost", value=int(default_val) if default_val is not None else None, key=f"capex_cost_{i}{key_suffix}", step=100000, help=f"Estimated Capital Expenditures for Fiscal Year {yr}.")
            with col2:
                life = st.number_input(f"FY{yr} Useful Life (mo)", value=int(default_l) if default_l is not None else None, key=f"capex_life_{i}{key_suffix}", step=12, help=f"Amortization period for CAPEX in Fiscal Year {yr} in months.")
                
            capex_useful_lives[yr] = int(life) if life is not None else 0
            if val is not None and val > 0:
                capex_schedule[yr] = float(val)
                
        capex_total = sum(capex_schedule.values())
        st.info(f"Total Capex Cost: {params.get('Currency', 'INR')} {capex_total:,.2f}")
        
        st.markdown("**Maintenance (PM) Schedule**")
        pm_dict = params.get("PM Schedule", {})
        non_zero_pm_years = [y for y, val in pm_dict.items() if val > 0]
        
        if non_zero_pm_years:
            default_num_pm = max(1, max(non_zero_pm_years) - start_year + 1)
        else:
            default_num_pm = 6
            
        max_pm = max(10, default_num_pm)
        num_pm = st.selectbox("Number of PM Years", list(range(1, max_pm + 1)), index=default_num_pm - 1, key=f"num_pm{key_suffix}", help="Select the number of years for Preventive Maintenance schedule.")
        
        pm_schedule = {}
        for i in range(num_pm):
            yr = start_year + i
            default_val = pm_dict.get(yr, None)
            val = st.number_input(f"PM FY{yr}", value=int(default_val) if default_val is not None else None, key=f"pm_cost_{i}{key_suffix}", step=50000, help=f"Preventive Maintenance budget allocation for Fiscal Year {yr}.")
            if val is not None and val > 0:
                pm_schedule[yr] = float(val)
                
        pm_cost_total = sum(pm_schedule.values())
        st.info(f"Total PM Cost: {params.get('Currency', 'INR')} {pm_cost_total:,.2f}")

    # Expander 3: Financial Rates & ARO Restoration
    with st.sidebar.expander("▶ Rates & ARO Restoration"):
        wacc_val = params.get("Cost of Capital")
        wacc_pct = float(wacc_val * 100) if wacc_val is not None else None
        wacc = st.number_input("Cost of Capital (WACC) %", value=wacc_pct, format="%.2f", step=0.01, key=f"wacc{key_suffix}", help="Weighted Average Cost of Capital used for investment analysis and discounting cash flows.")
        
        imputed_val = params.get("Imputed Interest Rate")
        imputed_pct = float(imputed_val * 100) if imputed_val is not None else None
        imputed_rate = st.number_input("Imputed Interest Rate %", value=imputed_pct, format="%.2f", step=0.01, key=f"imputed_rate{key_suffix}", help="The implicit rate of interest in the lease, or estimated discount rate.")
        
        exchange_rate = st.number_input("Forex Rate (LC/Euro)", value=float(params["Exchange Rate"]) if params["Exchange Rate"] is not None else None, key=f"exchange_rate{key_suffix}", step=0.1, help="The foreign currency exchange rate (INR per 1 Euro) for reporting.")

    # Pack overrides into active UI parameters
    ui_params = {
        "REU Name": reu_name,
        "Building Name": bldg_name,
        "City": city,
        "Country": params["Country"],
        "Currency": params["Currency"],
        "Lease Type": params["Lease Type"],
        "Chargeable Area Sqft": float(area) if area is not None else None,
        "Agreement Start Date": start_date,
        "Agreement End Date": end_date,
        "Rent Per Sqft": float(rent_sqft) if rent_sqft is not None else None,
        "Quoted CAM": float(cam_sqft) if cam_sqft is not None else None,
        "Escalation %": float(rent_esc) if rent_esc is not None else 0.0,
        "Escalation Frequency Months": int(rent_esc_freq) if rent_esc_freq is not None else None,
        "CAM Escalation %": float(cam_esc) if cam_esc is not None else 0.0,
        "CAM Escalation Frequency Months": int(cam_esc_freq) if cam_esc_freq is not None else None,
        "Security Deposit Amount": float(sec_deposit) if sec_deposit is not None else None,
        "Addnl.Deposit -energy(Refundable)": float(energy_dep) if energy_dep is not None else None,
        "4 Wheeler Slots": int(four_w_slots) if four_w_slots is not None else None,
        "4 Wheeler Rate": float(four_w_rate) if four_w_rate is not None else None,
        "2 Wheeler Slots": int(two_w_slots) if two_w_slots is not None else None,
        "2 Wheeler Rate": float(two_w_rate) if two_w_rate is not None else None,
        "Parking Escalation %": float(parking_esc) if parking_esc is not None else 0.0,
        "Parking Escalation Frequency Months": int(parking_esc_freq) if parking_esc_freq is not None else None,
        
        "Fitout Cost": float(fitout_total) if fitout_total is not None else 0.0,
        "Fitout Cost Breakdown": fitout_costs,
        "Fitout Useful Lives": fitout_useful_lives,
        "Fitout Active Months Breakdown": fitout_active_months_breakdown,
        "Capex Schedule": capex_schedule,
        "Capex Useful Lives": capex_useful_lives,
        "PM Cost Over Lease": float(pm_cost_total) if pm_cost_total is not None else 0.0,
        "PM Schedule": pm_schedule,
        
        "Cost of Capital": (float(wacc) / 100.0) if wacc is not None else None,
        "Discount Rate": float(params.get("Discount Rate", params.get("Incremental Borrowing Rate", 0.08))) if params.get("Discount Rate") is not None else 0.08,
        "Incremental Borrowing Rate": float(params.get("Incremental Borrowing Rate", params.get("Discount Rate", 0.08))) if params.get("Incremental Borrowing Rate") is not None else 0.08,
        "Imputed Interest Rate": (float(imputed_rate) / 100.0) if imputed_rate is not None else None,
        "Ready Reckoner Rate": float(params.get("Ready Reckoner Rate", 15000.0)) if params.get("Ready Reckoner Rate") is not None else 15000.0,
        "Exchange Rate": float(exchange_rate) if exchange_rate is not None else None,
        "Incremental Restoration Cost Sqft": float(params["Incremental Restoration Cost Sqft"]) if params.get("Incremental Restoration Cost Sqft") is not None else None,
        "Opex Others Per Month": float(opex_others) if opex_others is not None else 0.0,
        "Opex II Per Month": float(opex_ii) if opex_ii is not None else 0.0,
        "Lease Term Months": round((end_date - start_date).days / 30.4167) if (start_date and end_date) else 0
    }

    return ui_params, upload_occurred, parse_error_msg
