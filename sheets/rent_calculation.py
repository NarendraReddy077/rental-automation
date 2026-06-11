import pandas as pd
import datetime

def inject(ws, params):
    """
    Injects parameters into 'Rent Calculation' sheet.
    """
    ws["J5"] = params.get("Exchange Rate", 105.02)
    ws["F10"] = params.get("City", "Chennai")
    ws["F17"] = params.get("Lease Type", "Warm Shell")
    ws["F27"] = "='CAPEX and PM'!E3"
    ws["F33"] = params.get("Security Deposit Months", 10)
    ws["F69"] = params.get("Addnl.Deposit -energy(Refundable)", 0.0)
    ws["F117"] = params.get("Opex Others Per Month", 654.0)
    ws["F118"] = "=F117*7%"
    ws["F122"] = params.get("Opex II Per Month", 0.0)
    ws["G122"] = "=F122"
    
    # Link parking counts and rates dynamically to the Lease Rent sheet input table
    ws["F43"] = "='Lease Rent'!D22"
    ws["F44"] = "='Lease Rent'!C22"
    ws["F46"] = "='Lease Rent'!D23"
    ws["F47"] = "='Lease Rent'!C23"

    # Link expected date of Agreement to close to the correct dynamic Lease Rent cell
    term_months = params.get("Lease Term Months", 72)
    ws["F51"] = f"='Lease Rent'!C{30 + term_months}"

    # Disable Outline Level selectors / group symbols in the sheet
    if hasattr(ws, 'sheet_properties') and ws.sheet_properties:
        if hasattr(ws.sheet_properties, 'outlinePr') and ws.sheet_properties.outlinePr:
            ws.sheet_properties.outlinePr.showOutlineSymbols = False
    
    if hasattr(ws, 'sheet_view') and ws.sheet_view:
        ws.sheet_view.showOutlineSymbols = False
        
    if hasattr(ws, 'views') and ws.views and hasattr(ws.views, 'sheetView'):
        for sv in ws.views.sheetView:
            sv.showOutlineSymbols = False

def simulate(params):
    """
    Simulates evaluated figures in the Rent Calculation sheet matching Excel's formulas exactly.
    """
    import datetime
    
    currency = params.get("Currency", "INR")
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
    capital_rate = params.get("Cost of Capital", 0.105) or 0.105
    imputed_rate = params.get("Imputed Interest Rate", 0.0711) or 0.0711
    term_months = params.get("Lease Term Months", 72) or 72
    
    start_date = params.get("Agreement Start Date")
    if isinstance(start_date, str):
        try:
            start_date = datetime.datetime.strptime(start_date, "%Y-%m-%d").date()
        except ValueError:
            start_date = None
    if start_date is None:
        start_date = datetime.date(2026, 4, 1)
        
    end_date = params.get("Agreement End Date")
    if isinstance(end_date, str):
        try:
            end_date = datetime.datetime.strptime(end_date, "%Y-%m-%d").date()
        except ValueError:
            end_date = None
    if end_date is None:
        end_date = datetime.date(2032, 3, 31)

    # 1. Period in months using EOMONTH logic of Excel
    def get_eomonth(d, m_offset):
        y = d.year + (d.month + m_offset - 1) // 12
        m = (d.month + m_offset - 1) % 12 + 1
        if m == 12:
            next_month = datetime.date(y + 1, 1, 1)
        else:
            next_month = datetime.date(y, m + 1, 1)
        eom = next_month - datetime.timedelta(days=1)
        return eom

    end_date_for_period = get_eomonth(start_date, term_months - 1)
    period_months = (end_date_for_period - start_date).days / 365.0 * 12.0

    # 2. Month-by-month escalations
    current_rent = rent_rate
    current_cam = cam_rate
    initial_parking = (four_w_slots * four_w_rate) + (two_w_slots * two_w_rate)
    current_parking = initial_parking

    months_rent = []
    months_cam = []
    months_parking = []

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

    total_rent = sum(months_rent)
    total_cam = sum(months_cam)
    total_parking = sum(months_parking)

    avg_rent = total_rent / (area_sqft * period_months)
    avg_cam = total_cam / (area_sqft * period_months)
    avg_parking_per_sqft = (total_parking / period_months) / area_sqft

    # 3. Security deposit carrying cost
    sd_interest = sd_amount * capital_rate / 12.0 * period_months
    energy_interest = energy_deposit * capital_rate * 3.0
    total_sd_interest = sd_interest + energy_interest
    carrying_cost_per_sqft = total_sd_interest / (area_sqft * period_months)

    # 4. Stamp duty
    avg_rent_yr = total_rent / period_months * 12.0
    avg_parking_yr = (total_parking * 1.15) / 5.0
    avg_cam_yr = (total_cam * 1.15) / 5.0
    g204 = (avg_rent_yr + avg_parking_yr + avg_cam_yr + sd_amount) * 1.18
    total_stamp_duty = g204 * 0.015
    eff_stamp_duty = total_stamp_duty / (area_sqft * period_months)

    # Helper function for distributing months
    def calculate_fy_months_local(s_dt, t_m, f_t_m):
        fy_months = {}
        c_yr = s_dt.year
        rem_m = min(t_m, f_t_m)
        f_yr_m = min(rem_m, 12 - s_dt.month + 1)
        fy_months[c_yr] = f_yr_m
        rem_m -= f_yr_m
        while rem_m > 0:
            c_yr += 1
            mns = min(rem_m, 12)
            fy_months[c_yr] = mns
            rem_m -= mns
        return fy_months

    # 5. Fitout Amortization & Imputed Interest
    fitouts = list(params.get("Fitout Cost Breakdown", []))
    fitout_lifes = list(params.get("Fitout Useful Lives", []))
    active_months_breakdown = list(params.get("Fitout Active Months Breakdown", []))

    if not fitouts:
        fitouts = [14000000.0, 5000000.0, 15000000.0]
        fitout_lifes = [72, 30, 48]
    if not active_months_breakdown:
        active_months_breakdown = []
        for life in fitout_lifes:
            active_months_breakdown.append(
                calculate_fy_months_local(start_date, term_months, life)
            )

    years = sorted(list(set(yr for phase in active_months_breakdown for yr in phase.keys())))
    if not years:
        years = list(range(start_date.year, start_date.year + 10))

    # In Excel, fitouts are depreciated over total_active_months
    used_dep_by_year = {}
    for yr in years:
        dep_yr = 0.0
        for k in range(len(fitouts)):
            cost_k = fitouts[k]
            phase_m = active_months_breakdown[k] if k < len(active_months_breakdown) else {}
            total_active_k = sum(phase_m.values())
            m_k = phase_m.get(yr, 0)
            if m_k > 0 and total_active_k > 0:
                dep_yr += (cost_k / total_active_k) * m_k
        used_dep_by_year[yr] = dep_yr

    book_val_begin = {}
    book_val_end = {}
    avg_book_val = {}
    imputed_interest = {}
    current_book_val = sum(fitouts)
    for yr in years:
        book_val_begin[yr] = current_book_val
        used_dep = used_dep_by_year[yr]
        current_book_val -= used_dep
        book_val_end[yr] = current_book_val
        avg_book_val[yr] = (book_val_begin[yr] + book_val_end[yr]) / 2.0
        
        phase1_m = active_months_breakdown[0].get(yr, 0) if active_months_breakdown else 0
        imputed_interest[yr] = avg_book_val[yr] * imputed_rate / 12.0 * phase1_m

    sum_used_dep = sum(used_dep_by_year.values())
    sum_imputed_interest = sum(imputed_interest.values())
    phase1_total_months = sum(active_months_breakdown[0].values()) if active_months_breakdown else term_months
    e30_fitout = (sum_used_dep + sum_imputed_interest) / phase1_total_months
    fitout_sqft_mo = e30_fitout / area_sqft

    # 6. Capex Amortization & Imputed Interest
    capex_sched = params.get("Capex Schedule", {})
    capex_lives = params.get("Capex Useful Lives", {})

    def get_capex_tranche_months(s_dt, t_m, tranche_idx, useful_life):
        tranche_yr = s_dt.year + tranche_idx - 1
        if tranche_idx == 1:
            return calculate_fy_months_local(s_dt, t_m, useful_life)
        else:
            first_year_months = 12 - s_dt.month + 1
            elapsed = first_year_months + 12 * (tranche_idx - 2)
            remaining = max(0, t_m - elapsed)
            if remaining <= 0:
                return {}
            return calculate_fy_months_local(datetime.date(tranche_yr, 1, 1), remaining, useful_life)

    start_year = start_date.year
    tranches = []
    for i in range(1, 9):
        yr = start_year + i - 1
        cost = capex_sched.get(yr, 0.0)
        if i == 1:
            default_life = term_months
        else:
            first_year_months = 12 - start_date.month + 1
            elapsed = first_year_months + 12 * (i - 2)
            default_life = max(0, term_months - elapsed)
        life = capex_lives.get(yr, default_life)
        tranches.append({"cost": cost, "life": life, "start_year": yr})

    capex_used_dep = {}
    for yr in years:
        dep_yr = 0.0
        for idx, t in enumerate(tranches):
            if yr >= t["start_year"] and t["life"] > 0:
                active_m = get_capex_tranche_months(start_date, term_months, idx + 1, t["life"]).get(yr, 0)
                if active_m > 0:
                    dep_yr += (t["cost"] / t["life"]) * active_m
        capex_used_dep[yr] = dep_yr

    capex_book_begin = {}
    capex_book_end = {}
    capex_avg_book = {}
    capex_interest = {}
    current_capex_book = sum(t["cost"] for t in tranches)

    for yr in years:
        capex_book_begin[yr] = current_capex_book
        dep_yr = capex_used_dep[yr]
        current_capex_book -= dep_yr
        capex_book_end[yr] = current_capex_book
        capex_avg_book[yr] = (capex_book_begin[yr] + capex_book_end[yr]) / 2.0
        
        active_term = calculate_fy_months_local(start_date, term_months, term_months).get(yr, 0)
        capex_interest[yr] = capex_avg_book[yr] * imputed_rate / 12.0 * active_term

    sum_capex_dep = sum(capex_used_dep.values())
    sum_capex_interest = sum(capex_interest.values())

    e90_capex = (sum_capex_dep + sum_capex_interest) / term_months if term_months > 0 else 0.0
    capex_sqft_mo = e90_capex / area_sqft

    # 7. PM per sqft/mo
    pm_cost_total = params.get("PM Cost Over Lease", 0.0)
    pm_sqft_mo = pm_cost_total / (area_sqft * term_months) if term_months > 0 else 0.0

    # 8. ARO per sqft/mo
    aro_rate = 82.6 if area_sqft < 50000 else 62.72
    aro_sqft_mo = aro_rate / period_months

    # 9. Net Rent II
    net_rent_2 = fitout_sqft_mo + capex_sqft_mo + pm_sqft_mo + aro_sqft_mo

    # 10. Property Management Fees (Net Rent I)
    eff_pm_fees = (avg_rent + avg_parking_per_sqft + carrying_cost_per_sqft + avg_cam + eff_stamp_duty + net_rent_2) * 0.07

    # 11. Net Rent I
    net_rent_1 = avg_rent + avg_parking_per_sqft + carrying_cost_per_sqft + avg_cam + eff_stamp_duty + eff_pm_fees

    # 12. Opex
    opex_others = params.get("Opex Others Per Month", 63.0)
    opex_1 = opex_others * 1.07
    opex_2 = params.get("Opex II Per Month", 0.0)

    # 13. Total
    total_outflow = net_rent_1 + net_rent_2 + opex_1 + opex_2

    return {
        "Base Rental Rate (sqft/mo)": rent_rate,
        "Common Area Maintenance (CAM)": cam_rate,
        "Security Deposit Imputed Carrying Cost (sqft/mo)": carrying_cost_per_sqft,
        "Stamp Duty & Registration (sqft/mo)": eff_stamp_duty,
        "Property Management Fees (sqft/mo)": eff_pm_fees,
        "Net Rent I (Standard)": net_rent_1,
        "Net Rent II (Refinancing)": net_rent_2,
        "Total Net Rental Rate (sqft/mo)": net_rent_1 + net_rent_2,
        "Estimated Stamp Duty & Registration Amount": total_stamp_duty,
        "Opex Others Per Month": opex_1,
        "Opex II Per Month": opex_2,
        "Total OpEx Per Month": opex_1 + opex_2,
        "Total Occupancy Cost": total_outflow
    }

