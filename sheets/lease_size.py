import pandas as pd
import numpy as np

def inject(ws, params):
    """
    Descriptive injection for 'Lease Size' worksheet.
    Updates column I cell I26 to dynamically sum fitout phases starting from Phase 2.
    """
    fitouts = list(params.get("Fitout Cost Breakdown", []))
    N = len(fitouts)
    
    # Update row 26 column I to sum fitouts starting from Phase 2 (row 12, 18, 24, ...)
    if N > 1:
        terms = [f"'CAPEX and PM'!B{6 + 6 * k}" for k in range(1, N)]
        ws.cell(row=26, column=9, value="=" + "+".join(terms))
    else:
        ws.cell(row=26, column=9, value=0)

def simulate(params, capex_pm_df=None):
    """
    Simulates year-by-year cash outflows and NPV matching cell N36 of Lease Size.
    """
    area_sqft = params.get("Chargeable Area Sqft", 9515.0) or 9515.0
    rent_rate = params.get("Rent Per Sqft", 120.0) or 120.0
    cam_rate = params.get("Quoted CAM", 15.48) or 15.48
    rent_esc_pct = params.get("Escalation %", 0.15)
    rent_esc_freq = params.get("Escalation Frequency Months", 36) or 36
    cam_esc_pct = params.get("CAM Escalation %", 0.05)
    cam_esc_freq = params.get("CAM Escalation Frequency Months", 12) or 12
    parking_esc_pct = params.get("Parking Escalation %", rent_esc_pct)
    parking_esc_freq = params.get("Parking Escalation Frequency Months", rent_esc_freq) or 36
    
    four_w_slots = params.get("4 Wheeler Slots", 80)
    four_w_rate = params.get("4 Wheeler Rate", 1500.0)
    two_w_slots = params.get("2 Wheeler Slots", 50)
    two_w_rate = params.get("2 Wheeler Rate", 1000.0)
    
    sd_amount = params.get("Security Deposit Amount", 0.0)
    energy_deposit = params.get("Addnl.Deposit -energy(Refundable)", 0.0)
    fitout_cost = params.get("Fitout Cost", 0.0)
    wacc = params.get("Cost of Capital", 0.105) or 0.105
    ex_rate = params.get("Exchange Rate", 105.02) or 105.02
    term_months = params.get("Lease Term Months", 72) or 72
    start_date = params.get("Agreement Start Date")
    
    # If start_date is a string, convert to date
    if isinstance(start_date, str):
        import datetime
        try:
            start_date = datetime.datetime.strptime(start_date, "%Y-%m-%d").date()
        except ValueError:
            start_date = None
            
    # 1. Calculate monthly values (Rent, CAM, Parking)
    months_rent = []
    months_cam = []
    months_parking = []
    
    current_rent = rent_rate
    current_cam = cam_rate
    initial_parking = (four_w_slots * four_w_rate) + (two_w_slots * two_w_rate)
    current_parking = initial_parking
    
    for m in range(1, term_months + 1):
        if m > 1:
            if (m - 1) % rent_esc_freq == 0:
                current_rent *= (1.0 + rent_esc_pct)
            if (m - 1) % cam_esc_freq == 0:
                current_cam *= (1.0 + cam_esc_pct)
            if (m - 1) % parking_esc_freq == 0:
                current_parking *= (1.0 + parking_esc_pct)
        
        months_rent.append(current_rent * area_sqft)
        months_cam.append(current_cam * area_sqft)
        months_parking.append(current_parking)
        
    total_rent_inr = sum(months_rent)
    total_cam_inr = sum(months_cam)
    total_parking_inr = sum(months_parking)
    
    # Stamp Duty & Registration calculation matching workbook's formula:
    # 1.5% of (Average Rent + Average Car Park + Average CAM + Security Deposit) * 1.18
    g200 = (total_rent_inr / term_months) * 12
    g201 = (total_parking_inr * 1.15) / 5
    g202 = (total_cam_inr * 1.15) / 5
    g203 = sd_amount
    
    g204 = (g200 + g201 + g202 + g203) * 1.18
    total_stamp_duty = g204 * 0.015
    
    # 2. Group months into 10 years (Year 1 is 8 months, others are 12 months)
    years_rent = [0.0] * 10
    years_cam = [0.0] * 10
    years_parking = [0.0] * 10
    
    for m in range(1, term_months + 1):
        if m <= 8:
            y_idx = 0
        else:
            y_idx = (m - 9) // 12 + 1
            
        if y_idx < 10:
            years_rent[y_idx] += months_rent[m - 1]
            years_cam[y_idx] += months_cam[m - 1]
            years_parking[y_idx] += months_parking[m - 1]
            
    # Capex schedule and PM schedule by year index
    years_capex = [0.0] * 10
    years_pm = [0.0] * 10
    start_year = start_date.year if start_date else 2026
    
    for yr, val in params.get("Capex Schedule", {}).items():
        y_idx = yr - start_year
        if 0 <= y_idx < 10:
            years_capex[y_idx] = val
            
    for yr, val in params.get("PM Schedule", {}).items():
        y_idx = yr - start_year
        if 0 <= y_idx < 10:
            years_pm[y_idx] = val
            
    # Security deposit flows (Year 1 outflow, Year 7 refund)
    years_sd = [0.0] * 10
    sd_flow = sd_amount + energy_deposit + energy_deposit  # Energy deposit is added twice in formula
    years_sd[0] = sd_flow
    years_sd[6] = -sd_flow
    
    # 3. Create dataframe rows (for return value and details)
    rows = []
    years_total_euro = [0.0] * 10
    
    for i in range(10):
        total_inr = years_rent[i] + years_cam[i] + years_parking[i] + years_capex[i] + years_pm[i] + years_sd[i]
        years_total_euro[i] = total_inr / ex_rate
        
        discount_factor = 1.0 / ((1.0 + wacc) ** (i + 1))
        pv_outflow = years_total_euro[i] * discount_factor
        
        rows.append({
            "Year": f"Year {i+1}",
            "Rent Cost": round(years_rent[i] / ex_rate, 2),
            "CAM Cost": round(years_cam[i] / ex_rate, 2),
            "Parking Cost": round(years_parking[i] / ex_rate, 2),
            "Capex": round(years_capex[i] / ex_rate, 2),
            "Maintenance": round(years_pm[i] / ex_rate, 2),
            "Security Deposit Flow": round(years_sd[i] / ex_rate, 2),
            "Total Outflow": round(years_total_euro[i], 2),
            "PV Factor": round(discount_factor, 4),
            "Present Value Outflow": round(pv_outflow, 2)
        })
        
    df = pd.DataFrame(rows)
    
    # Calculate NPV of Euro cash flows
    npv_euro = 0.0
    for i in range(10):
        npv_euro += years_total_euro[i] / ((1.0 + wacc) ** (i + 1))
        
    # One time costs in Euro
    one_time_cost_euro = (total_stamp_duty + fitout_cost) / ex_rate
    
    # Total Project NPV in Euro Millions
    total_pv_euro = npv_euro + one_time_cost_euro
    npv_millions = total_pv_euro / 1000000.0
    
    return df, npv_millions
