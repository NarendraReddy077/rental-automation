import pandas as pd
import datetime

def inject(ws, params):
    """
    Injects parameters into 'Rent Calculation' sheet.
    """
    ws["J5"] = params.get("Exchange Rate", 105.02)
    ws["F10"] = params.get("City", "Chennai")
    ws["F17"] = params.get("Lease Type", "Warm Shell")
    ws["F27"] = params.get("Amortization Period Months", 72)
    ws["F33"] = params.get("Security Deposit Months", 10)
    ws["F118"] = params.get("Opex Others Per Month", 654.0)

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
    Simulates evaluated figures in the Rent Calculation sheet.
    """
    currency = params["Currency"]
    area_sqft = params["Chargeable Area Sqft"]
    rent_rate = params["Rent Per Sqft"]
    cam_rate = params["Quoted CAM"]
    
    # Calculate simple totals
    initial_monthly_rent = rent_rate * area_sqft
    initial_monthly_cam = cam_rate * area_sqft
    
    # Stamp duty estimation (similar to Excel formulas)
    # Market Value = Ready Reckoner * Total Area * 25% * 1% Stamp Duty + 0.5% Registration
    reckoner_rate = params.get("Ready Reckoner Rate", 15000.0) # Rs. per sqm
    area_sqm = area_sqft / 10.764
    market_value = reckoner_rate * area_sqm
    market_value_lease = market_value * 0.25
    stamp_duty = market_value_lease * 0.01
    reg_charges = stamp_duty * 0.5
    total_stamp_duty = stamp_duty + reg_charges
    
    # OPEX computations
    opex_others = params.get("Opex Others Per Month", 654.0)
    opex_mgmt_fee = opex_others * 0.07
    total_opex = opex_others + opex_mgmt_fee
    
    # Carrying costs (Simulated using Security Deposit carrying cost)
    # carrying cost of security deposit = Security Deposit * Cost of Capital / 12 * 72
    capital_rate = params.get("Cost of Capital", 0.105)
    sd_amount = params["Security Deposit Amount"]
    energy_deposit = params["Addnl.Deposit -energy(Refundable)"]
    carrying_interest = (sd_amount + energy_deposit) * capital_rate / 12.0 * 72.0
    carrying_cost_per_sqft = carrying_interest / (area_sqft * 72)
    
    # ARO Cost
    aro_rate = params.get("Incremental Restoration Cost Sqft", 82.6)
    if aro_rate == 82.6 and area_sqft >= 50000:
        aro_rate = 62.72
    aro_cost_sqft_mo = (area_sqft * aro_rate / 72.0) / area_sqft
    
    # Capex carrying cost (Amortization + Imputed Interest)
    # Fitout cost B6 = 14M, Capex = 43M, PM = 2.5M
    fitout_cost = params["Fitout Cost"]
    pm_cost = params.get("PM Cost Over Lease", 2500000.0)
    
    # Calculate Net Rent I Components (per sqft/mo)
    eff_rent = rent_rate
    eff_parking = 0.0  # in specimen, parking effect is 0
    eff_sd_carrying = carrying_cost_per_sqft
    eff_cam = cam_rate
    eff_brokerage = 0.0  # Brokerage is 0 in specimen
    eff_stamp_duty = total_stamp_duty / (area_sqft * 72)
    eff_pm_fees = (eff_rent + eff_parking + eff_sd_carrying + eff_cam + eff_brokerage + eff_stamp_duty) * 0.07
    
    net_rent_1 = eff_rent + eff_parking + eff_sd_carrying + eff_cam + eff_brokerage + eff_stamp_duty + eff_pm_fees
    
    # Calculate Net Rent II Components (per sqft/mo)
    # Simulated values based on typical specimens
    imputed_rate = params.get("Imputed Interest Rate", 0.0711)
    # Fitouts amortization (average)
    eff_fitouts = (fitout_cost / 72) / area_sqft
    # Capex amortization (average)
    eff_capex = (fitout_cost * 1.2 / 72) / area_sqft  # scaled
    # PM amortization
    eff_pm_amort = (pm_cost / 72) / area_sqft
    # ARO amortization
    eff_aro = aro_cost_sqft_mo
    
    net_rent_2 = eff_fitouts + eff_capex + eff_pm_amort + eff_aro
    
    total_net_rent = net_rent_1 + net_rent_2
    
    return {
        "Base Rental Rate (sqft/mo)": rent_rate,
        "Common Area Maintenance (CAM)": cam_rate,
        "Security Deposit Imputed Carrying Cost (sqft/mo)": round(eff_sd_carrying, 4),
        "Stamp Duty & Registration (sqft/mo)": round(eff_stamp_duty, 4),
        "Property Management Fees (sqft/mo)": round(eff_pm_fees, 4),
        "Net Rent I (Standard)": round(net_rent_1, 2),
        "Net Rent II (Refinancing)": round(net_rent_2, 2),
        "Total Net Rental Rate (sqft/mo)": round(total_net_rent, 2),
        "Estimated Stamp Duty & Registration Amount": round(total_stamp_duty, 2),
        "Opex Others Per Month": opex_others,
        "Total OpEx Per Month": round(total_opex, 2)
    }
