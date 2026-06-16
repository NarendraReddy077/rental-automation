import pandas as pd

def inject(ws, params):
    """
    Injects inputs into the 'Security Deposit' sheet.
    """
    ws["B9"] = 0.0  # Fitout Deposit
    ws["B10"] = params["Security Deposit Amount"]
    ws["B11"] = "='Rent Calculation'!F71"  # Additional deposit
    ws["B12"] = params["Addnl.Deposit -energy(Refundable)"]
    
    # Write the user incremental ARO rate to cell B23 if overridden
    aro_rate = params.get("Incremental Restoration Cost Sqft")
    if aro_rate is not None and not pd.isna(aro_rate):
        ws["B23"] = aro_rate

def simulate(params):
    """
    Simulates Security Deposit carrying costs and ARO parameters.
    """
    currency = params["Currency"]
    area_sqft = params["Chargeable Area Sqft"]
    capital_rate = params.get("Cost of Capital", 0.105)
    term_months = params.get("Lease Term Months", 72)
    if pd.isna(term_months) or term_months is None:
        term_months = 72
        
    sd_amt = params["Security Deposit Amount"]
    energy_dep = params["Addnl.Deposit -energy(Refundable)"]
    
    # Calculate CAM Security Deposit (Additional deposit)
    cam_rate = params.get("Quoted CAM", 15.48) or 15.48
    maint_sd_months = params.get("Maintenance Security Deposit Months", 0.0) or 0.0
    cam_sd_amount = cam_rate * (area_sqft or 0.0) * maint_sd_months
    
    # Interest carrying costs calculations
    # carrying interest cost = Deposit * WACC / 12 * term_months
    sd_carrying_cost = sd_amt * capital_rate / 12 * term_months
    cam_sd_carrying_cost = cam_sd_amount * capital_rate / 12 * term_months
    energy_carrying_cost = energy_dep * capital_rate / 12 * 36  # In specimen, energy deposit term is 36 months!
    
    total_deposit = sd_amt + cam_sd_amount + energy_dep
    total_carrying_cost = sd_carrying_cost + cam_sd_carrying_cost + energy_carrying_cost
    carrying_cost_rate_sqft = total_carrying_cost / (area_sqft * term_months) if (area_sqft and term_months) else 0.0
    
    # ARO (Asset Retirement Obligation) calculations
    aro_rate = params.get("Incremental Restoration Cost Sqft")
    if aro_rate is None or pd.isna(aro_rate):
        aro_rate = 82.6 if (area_sqft is not None and area_sqft < 50000) else 62.72
    elif aro_rate == 82.6 and area_sqft is not None and area_sqft >= 50000:
        aro_rate = 62.72
        
    total_aro_cost = (area_sqft or 0.0) * aro_rate
    aro_per_month = total_aro_cost / term_months if term_months else 0.0
    aro_conversion_factor = aro_per_month / area_sqft if area_sqft else 0.0
    
    return {
        "Fitout Deposit": {"Amount": 0.0, "Carrying Cost": 0.0},
        "Security Deposit": {"Amount": sd_amt, "Carrying Cost": sd_carrying_cost},
        "Additional Deposit": {"Amount": cam_sd_amount, "Carrying Cost": cam_sd_carrying_cost},
        "Energy Deposit": {"Amount": energy_dep, "Carrying Cost": energy_carrying_cost},
        "Total Deposits": {"Amount": total_deposit, "Carrying Cost": total_carrying_cost},
        "Carrying Cost per Sqft/mo": carrying_cost_rate_sqft,
        "ARO Rate per Sqft (as of 2017)": aro_rate,
        "Total ARO Capital Cost Asset": total_aro_cost,
        "Monthly ARO Amortization Cost": aro_per_month,
        "ARO Conversion Factor (sqft/pm)": aro_conversion_factor
    }
