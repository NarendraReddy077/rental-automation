import streamlit as st
import datetime
import os
import sheets.capex_pm as capex_pm
import sheets.security_deposit as security_deposit
import sheets.rent_calculation as rent_calculation
import sheets.lease_size as lease_size
from ui.styles import load_template
from core.excel_compiler import compile_output_workbook

def render_dashboard(ui_params, upload_occurred, parse_error_msg):
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
                area=f"{int(ui_params['Chargeable Area Sqft']):,}" if ui_params['Chargeable Area Sqft'] is not None else "0",
                duration=f"{duration_yrs:.2f}",
                currency=ui_params["Currency"],
                total_deposit=f"{total_sd:,.2f}" if total_sd is not None else "0.00"
            ), unsafe_allow_html=True)
        else:
            st.success(f"Workbook parameters for {ui_params['REU Name']} compiled successfully!")

        # --- LIVE SIMULATION MODEL W wake-up ---
        capex_pm_df = capex_pm.simulate(ui_params)
        sd_results = security_deposit.simulate(ui_params)
        rent_calc_results = rent_calculation.simulate(ui_params)
        lease_size_df, npv_value = lease_size.simulate(ui_params, capex_pm_df)

        # --- LIVE KPI DASHBOARD (HTML GRID) ---
        total_rent_cam = (ui_params["Rent Per Sqft"] or 0.0) + (ui_params["Quoted CAM"] or 0.0)
        duration_yrs = ui_params["Lease Term Months"] / 12.0
        total_sd_amount = (ui_params["Security Deposit Amount"] or 0.0) + (ui_params["Addnl.Deposit -energy(Refundable)"] or 0.0)
        total_fitout = ui_params["Fitout Cost"]
        total_capex = sum(ui_params["Capex Schedule"].values())
        total_pm = ui_params["PM Cost Over Lease"]
        
        net_rent_1 = rent_calc_results["Net Rent I (Standard)"]
        net_rent_2 = rent_calc_results["Net Rent II (Refinancing)"]
        opex_1 = rent_calc_results["Opex Others Per Month"]
        opex_2 = rent_calc_results["Opex II Per Month"]
        total_occupancy_cost = rent_calc_results["Total Occupancy Cost"]
        
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
                <span class="stat-label">Lease Duration</span>
                <span class="stat-value">{duration_yrs:.2f} yrs</span>
                <span class="stat-meta">{ui_params['Agreement Start Date'].strftime('%b %d, %Y')} to {ui_params['Agreement End Date'].strftime('%b %d, %Y')}</span>
            </div>
            <div class="stat-cell">
                <span class="stat-label">Total Fitout Cost</span>
                <span class="stat-value">{ui_params['Currency']} {total_fitout/1000000:,.2f} M</span>
                <span class="stat-meta">{len(ui_params['Fitout Cost Breakdown'])} Phases Configured</span>
            </div>
            <div class="stat-cell">
                <span class="stat-label">Total CAPEX</span>
                <span class="stat-value">{ui_params['Currency']} {total_capex/1000000:,.2f} M</span>
                <span class="stat-meta">{len(ui_params['Capex Schedule'])} Years Scheduled</span>
            </div>
            <div class="stat-cell">
                <span class="stat-label">Total PM Cost</span>
                <span class="stat-value">{ui_params['Currency']} {total_pm/1000000:,.2f} M</span>
                <span class="stat-meta">{len(ui_params['PM Schedule'])} Years Scheduled</span>
            </div>
            <div class="stat-cell">
                <span class="stat-label">Project NPV</span>
                <span class="stat-value">€ {npv_value:,.2f} M</span>
                <span class="stat-meta">WACC: {ui_params['Cost of Capital']*100:.2f}% | Forex: {ui_params['Exchange Rate']:.2f}</span>
            </div>
            <div class="stat-cell">
                <span class="stat-label">Total Occupancy Cost</span>
                <span class="stat-value">{ui_params['Currency']} {total_occupancy_cost:,.2f}</span>
                <div class="stat-sub-grid">
                    <div class="stat-sub-item">
                        <span class="stat-sub-label">Net Rent 1</span>
                        <span class="stat-sub-value">{net_rent_1:,.2f}</span>
                    </div>
                    <div class="stat-sub-item">
                        <span class="stat-sub-label">Net Rent 2</span>
                        <span class="stat-sub-value">{net_rent_2:,.2f}</span>
                    </div>
                    <div class="stat-sub-item">
                        <span class="stat-sub-label">Opex 1</span>
                        <span class="stat-sub-value">{opex_1:,.2f}</span>
                    </div>
                    <div class="stat-sub-item">
                        <span class="stat-sub-label">Opex 2</span>
                        <span class="stat-sub-value">{opex_2:,.2f}</span>
                    </div>
                </div>
            </div>
        </div>
        """
        st.markdown(metrics_html, unsafe_allow_html=True)

        # --- EXCEL WORKBOOK COMPILER GENERATION ---
        template_file_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "artifacts", "Rental Specimen.xlsx")
        
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
    else:
        info_tpl = load_template("info_card.html")
        if info_tpl:
            st.markdown(info_tpl, unsafe_allow_html=True)
