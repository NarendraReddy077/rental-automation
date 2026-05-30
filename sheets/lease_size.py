import pandas as pd
import numpy as np

def inject(ws, params):
    """
    Descriptive injection for 'Lease Size' worksheet.
    Since it is fully formula-driven in Excel, no hardcoded cell updates are necessary.
    """
    pass

def simulate(params, capex_pm_df=None):
    """
    Simulates year-by-year cash outflows and NPV.
    """
    area_sqft = params["Chargeable Area Sqft"]
    rent_rate = params["Rent Per Sqft"]
    cam_rate = params["Quoted CAM"]
    wacc = params.get("Cost of Capital", 0.105)
    
    # Calculate years 1-10 cash outflows
    # In specimen, rent escalates by 15% every 3 years.
    # Year 1-3: Rent = Base Rent
    # Year 4-6: Rent = Base Rent * 1.15
    # Year 7-9: Rent = Base Rent * 1.15 * 1.15
    # Year 10: Rent = Base Rent * 1.15 * 1.15 * 1.15
    
    # CAM escalates by 5% every year.
    
    years = list(range(1, 11))
    rows = []
    
    for yr in years:
        # Rent escalation
        if yr <= 3:
            rent_factor = 1.0
        elif yr <= 6:
            rent_factor = 1.0 + params.get("Escalation %", 0.15)
        elif yr <= 9:
            rent_factor = (1.0 + params.get("Escalation %", 0.15)) ** 2
        else:
            rent_factor = (1.0 + params.get("Escalation %", 0.15)) ** 3
            
        # CAM escalation
        cam_factor = (1.0 + params.get("CAM Escalation %", 0.05)) ** (yr - 1)
        
        yearly_rent = rent_rate * rent_factor * area_sqft * 12
        yearly_cam = cam_rate * cam_factor * area_sqft * 12
        
        # Add capex and maintenance if data is available from capex_pm simulation
        capex_cost = 0.0
        pm_cost = 0.0
        if capex_pm_df is not None and not capex_pm_df.empty:
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
