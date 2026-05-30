import pandas as pd
import numpy as np

def calculate_fy_months(start_date, term_months, fitout_term_months):
    """
    Dynamically distributes term_months across fiscal years (Calendar Years).
    """
    fy_months = {}
    current_year = start_date.year
    remaining_months = min(term_months, fitout_term_months)
    
    # First year months (from start_date month to Dec 31)
    first_year_months = min(remaining_months, 12 - start_date.month + 1)
    fy_months[current_year] = first_year_months
    remaining_months -= first_year_months
    
    # Subsequent years
    while remaining_months > 0:
        current_year += 1
        months = min(remaining_months, 12)
        fy_months[current_year] = months
        remaining_months -= months
        
    return fy_months

def inject(ws, params):
    """
    Injects parameters and schedules into 'CAPEX and PM' worksheet.
    """
    # Write fitout costs (Splitting if needed, otherwise write single fitout)
    fitouts = params.get("Fitout Cost Breakdown", [params["Fitout Cost"], 0, 0])
    # Make sure we have at least 3 elements
    while len(fitouts) < 3:
        fitouts.append(0)
        
    ws["B6"] = fitouts[0]
    ws["B12"] = fitouts[1]
    ws["B18"] = fitouts[2]
    
    # Write Capex schedule by FY
    capex_sched = params.get("Capex Schedule", {2026: 14000000.0, 2027: 12000000.0, 2028: 5000000.0, 2029: 6000000.0, 2030: 6000000.0})
    # Capex headers start at col F (col 6) on row 36
    for c in range(6, 16):
        fy = ws.cell(row=36, column=c).value
        if isinstance(fy, (int, float)):
            fy = int(fy)
            ws.cell(row=38, column=c, value=capex_sched.get(fy, 0.0))
            
    # Write Maintenance (PM) schedule by FY
    pm_sched = params.get("PM Schedule", {2026: 1500000.0, 2027: 200000.0, 2028: 200000.0, 2029: 200000.0, 2030: 200000.0, 2031: 200000.0})
    # PM headers are F95 to K95
    for c in range(6, 16):
        fy = ws.cell(row=2, column=c).value  # Fetch year from row 2
        if isinstance(fy, (int, float)):
            fy = int(fy)
            ws.cell(row=95, column=c, value=pm_sched.get(fy, 0.0))

    # Dynamically populate "Month needed" rows based on start date and term
    start_date = params["Agreement Start Date"]
    term_months = params.get("Lease Term Months", 72)
    if pd.isna(term_months) or term_months is None:
        term_months = 72
        
    f1_months = calculate_fy_months(start_date, term_months, 72)
    f2_months = calculate_fy_months(start_date, term_months, 30)
    f3_months = calculate_fy_months(start_date, term_months, 48)
    
    # Write months needed for Fitout 1, 2, 3
    for c in range(6, 16):
        fy = ws.cell(row=2, column=c).value
        if isinstance(fy, (int, float)):
            fy = int(fy)
            # Fitout 1 (Row 3)
            ws.cell(row=3, column=c, value=f1_months.get(fy, 0))
            # Fitout 2 (Row 9)
            ws.cell(row=9, column=c, value=f2_months.get(fy, 0))
            # Fitout 3 (Row 15)
            ws.cell(row=15, column=c, value=f3_months.get(fy, 0))

def simulate(params):
    """
    Simulates Capex & PM amortization schedules for UI preview.
    """
    start_date = params["Agreement Start Date"]
    imputed_rate = params["Imputed Interest Rate"]
    
    fitouts = params.get("Fitout Cost Breakdown", [params["Fitout Cost"], 0, 0])
    while len(fitouts) < 3:
        fitouts.append(0)
        
    capex_sched = params.get("Capex Schedule", {2026: 14000000.0, 2027: 12000000.0, 2028: 5000000.0, 2029: 6000000.0, 2030: 6000000.0})
    pm_sched = params.get("PM Schedule", {2026: 1500000.0, 2027: 200000.0, 2028: 200000.0, 2029: 200000.0, 2030: 200000.0, 2031: 200000.0})
    
    years = list(range(2026, 2036))
    
    # Fitout 1, 2, 3 calculations
    term_months = params.get("Lease Term Months", 72)
    if pd.isna(term_months) or term_months is None:
        term_months = 72
        
    f1_months = calculate_fy_months(start_date, term_months, 72)
    f2_months = calculate_fy_months(start_date, term_months, 30)
    f3_months = calculate_fy_months(start_date, term_months, 48)
    
    # Amortization variables
    f1_cost, f2_cost, f3_cost = fitouts[0], fitouts[1], fitouts[2]
    total_fitout_cost = sum(fitouts)
    
    # Capex section calculations
    # Capex is split into 5 tranches (FY26-FY30)
    # Tranche 1: 14M (72 mo), Tranche 2: 12M (60 mo), Tranche 3: 5M (48 mo), Tranche 4: 6M (36 mo), Tranche 5: 6M (24 mo)
    tranches = [
        {"cost": capex_sched.get(2026, 0.0), "life": 72, "start_year": 2026},
        {"cost": capex_sched.get(2027, 0.0), "life": 60, "start_year": 2027},
        {"cost": capex_sched.get(2028, 0.0), "life": 48, "start_year": 2028},
        {"cost": capex_sched.get(2029, 0.0), "life": 36, "start_year": 2029},
        {"cost": capex_sched.get(2030, 0.0), "life": 24, "start_year": 2030}
    ]
    
    rows = []
    
    # Keep track of book values for Capex
    current_capex_book = sum([t["cost"] for t in tranches])
    
    for yr in years:
        m1 = f1_months.get(yr, 0)
        m2 = f2_months.get(yr, 0)
        m3 = f3_months.get(yr, 0)
        
        # Used Depreciation Fitouts
        dep1 = (f1_cost / (72 / 12) / 12) * m1 if m1 > 0 else 0.0
        dep2 = (f2_cost / (30 / 12) / 12) * m2 if m2 > 0 else 0.0
        dep3 = (f3_cost / (48 / 12) / 12) * m3 if m3 > 0 else 0.0
        total_fitout_dep = dep1 + dep2 + dep3
        
        # Capex Depreciation
        # For each Capex tranche, calculate months needed and used depreciation
        capex_dep = 0.0
        for t in tranches:
            if yr >= t["start_year"]:
                # Calculate active months in yr
                active_months = calculate_fy_months(datetime.date(t["start_year"], 1, 1), term_months, t["life"]).get(yr, 0)
                if active_months > 0:
                    capex_dep += (t["cost"] / t["life"]) * active_months
                    
        # Capex Imputed Interest
        # In specimen: Average net book value * WACC / 12 * active_months
        # Let's approximate the average book value for Capex in yr
        active_term = calculate_fy_months(start_date, term_months, 72).get(yr, 0)
        capex_imputed_interest = 0.0
        if active_term > 0:
            # Simple linear approximation for preview
            capex_imputed_interest = (current_capex_book - capex_dep / 2.0) * imputed_rate / 12.0 * active_term
            current_capex_book -= capex_dep
            
        rows.append({
            "Fiscal Year": yr,
            "Fitout Months Active": m1,
            "Fitout Amortization": round(total_fitout_dep, 2),
            "Capex Injected": round(capex_sched.get(yr, 0.0), 2),
            "Capex Amortization": round(capex_dep, 2),
            "Capex Imputed Interest": round(capex_imputed_interest, 2),
            "Maintenance Cost": round(pm_sched.get(yr, 0.0), 2)
        })
        
    return pd.DataFrame(rows)
