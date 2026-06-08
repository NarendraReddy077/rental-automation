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
    Simulates year-by-year cash outflows and NPV.
    """
    area_sqft = params["Chargeable Area Sqft"]
    rent_rate = params["Rent Per Sqft"]
    cam_rate = params["Quoted CAM"]
    wacc = params.get("Cost of Capital", 0.105)
    
    years = list(range(1, 11))
    rows = []
    
    esc_freq_years = params.get("Escalation Frequency Months", 36) // 12
    if esc_freq_years <= 0:
        esc_freq_years = 3
    esc_pct = params.get("Escalation %", 0.15)
    term_months = params.get("Lease Term Months", 72)
    
    for yr in years:
        # Calculate active months of lease in this year
        active_months = max(0, min(12, term_months - 12 * (yr - 1)))
        
        # Rent escalation
        escalations = (yr - 1) // esc_freq_years
        rent_factor = (1.0 + esc_pct) ** escalations
            
        # CAM escalation
        cam_factor = (1.0 + params.get("CAM Escalation %", 0.05)) ** (yr - 1)
        
        yearly_rent = rent_rate * rent_factor * area_sqft * active_months
        yearly_cam = cam_rate * cam_factor * area_sqft * active_months
        
        # Add capex and maintenance if data is available from capex_pm simulation
        capex_cost = 0.0
        pm_cost = 0.0
        if active_months > 0 and capex_pm_df is not None and not capex_pm_df.empty:
            match_row = capex_pm_df[capex_pm_df["Fiscal Year"] == (params["Agreement Start Date"].year + yr - 1)]
            if not match_row.empty:
                capex_cost = match_row.iloc[0]["Capex Injected"]
                pm_cost = match_row.iloc[0]["Maintenance Cost"]
                
        # Total nominal cash outflow
        total_nominal = yearly_rent + yearly_cam + capex_cost + pm_cost
        
        # Discount factor
        discount_factor = 1.0 / ((1.0 + wacc) ** yr)
        pv_outflow = total_nominal * discount_factor
        
        rows.append({
            "Year": f"Year {yr}",
            "Rent Cost": round(yearly_rent, 2),
            "CAM Cost": round(yearly_cam, 2),
            "Capex": round(capex_cost, 2),
            "Maintenance": round(pm_cost, 2),
            "Total Outflow": round(total_nominal, 2),
            "PV Factor": round(discount_factor, 4),
            "Present Value Outflow": round(pv_outflow, 2)
        })
        
    df = pd.DataFrame(rows)
    # Calculate cumulative NPV
    npv_value = df["Present Value Outflow"].sum()
    
    return df, npv_value
