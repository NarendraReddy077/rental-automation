import io
import datetime
import pandas as pd
from sheets.input_parser.utils import get_str, get_float, get_int, parse_date, get_value_by_aliases

def parse_main_sheet(raw_data):
    """Parses the 'Main' sheet of the Excel workbook and returns a dictionary of raw parameters."""
    df = pd.read_excel(io.BytesIO(raw_data), sheet_name="Main")
    df.columns = [str(c).strip() for c in df.columns]
    
    # Dynamically determine Field Name and Sample Value columns
    field_col = None
    value_col = None
    for col in df.columns:
        col_lower = col.lower()
        if "field" in col_lower or "parameter" in col_lower or "name" in col_lower:
            field_col = col
        elif "value" in col_lower or "sample" in col_lower:
            value_col = col
    
    if field_col is None:
        field_col = df.columns[0] if len(df.columns) > 0 else "Field Name"
    if value_col is None:
        value_col = df.columns[1] if len(df.columns) > 1 else "Sample Value"

    # Build dictionary from Name-Value pairs
    raw_params = {}
    current_section = "rent"
    if field_col in df.columns and value_col in df.columns:
        for idx, row in df.iterrows():
            name = str(row[field_col]).strip()
            val = row[value_col]
            if pd.isna(val):
                val = None
            
            # Update current section based on keywords in name
            name_lower = name.lower()
            if "cam" in name_lower:
                current_section = "cam"
            elif "parking" in name_lower or "wheeler" in name_lower:
                current_section = "parking"
            elif "fitout" in name_lower or "capex" in name_lower:
                current_section = "capex"
            elif "pm" in name_lower or "maintenance" in name_lower:
                current_section = "pm"
            
            # If the key is a generic escalation key, qualify it with the section name
            key_to_store = name
            if name_lower in ["escalation %", "escalation frequency months", "escalation frequency"]:
                if current_section == "cam":
                    key_to_store = f"CAM {name}"
                elif current_section == "parking":
                    key_to_store = f"Parking {name}"
                elif current_section == "rent":
                    key_to_store = f"Rent {name}"
            
            raw_params[key_to_store] = val
            
    return raw_params

def extract_main_parameters(raw_params):
    """Extracts and normalizes base parameters from raw_params."""
    params = {}
    params["Lease ID"] = get_str(raw_params, "Lease ID", ["LeaseID", "Lease_ID"], "LEASE-001")
    params["REU Name"] = get_str(raw_params, "REU Name", ["REUName", "REU_Name", "REU"], "IN-CHEK2")
    params["Lease Type"] = get_str(raw_params, "Lease Type", ["LeaseType", "Lease_Type"], "Office")
    params["Building Name"] = get_str(raw_params, "Building Name", ["BuildingName", "Building_Name"], "SKCL Tech Park")
    params["City"] = get_str(raw_params, "City", [], "Chennai")
    params["Country"] = get_str(raw_params, "Country", [], "India")
    params["Currency"] = get_str(raw_params, "Currency", [], "INR")
    
    params["Chargeable Area Sqft"] = get_float(raw_params, "Chargeable Area Sqft", ["Chargeable Area Sq.ft.", "Chargeable Area Sq.ft", "Chargeable Area (Sqft)", "Chargeable Area", "Area Sqft", "Area", "Chargeable Office Area"], 9515.0)
    params["Parking Slots"] = get_int(raw_params, "Parking Slots", ["ParkingSlots", "Parking Slots Count"], 20)
    params["4 Wheeler Slots"] = get_int(raw_params, "4 Wheeler Slots", ["No. of 4 wheelers", "4-wheeler slots", "4 Wheeler Parking Slots", "No of 4 wheelers"], 80)
    params["4 Wheeler Rate"] = get_float(raw_params, "4 Wheeler Rate", ["Rate per 4 wheeler", "4-wheeler rate", "4 Wheeler Parking Rate"], 1500.0)
    params["2 Wheeler Slots"] = get_int(raw_params, "2 Wheeler Slots", ["No. of 2 wheelers", "2-wheeler slots", "2 Wheeler Parking Slots", "No of 2 wheelers"], 50)
    params["2 Wheeler Rate"] = get_float(raw_params, "2 Wheeler Rate", ["Rate per 2 wheeler", "2-wheeler rate", "2 Wheeler Parking Rate"], 1000.0)
    
    start_date_val = get_value_by_aliases(raw_params, ["Agreement Start Date", "Start Date", "Lease Start Date", "Commencement Date"])
    start_date = parse_date(start_date_val, datetime.date(2026, 4, 1))
    
    end_date_val = get_value_by_aliases(raw_params, ["Agreement End Date", "End Date", "Lease End Date", "Expiry Date"])
    end_date = parse_date(end_date_val, datetime.date(2031, 3, 31))
    
    params["Agreement Start Date"] = start_date
    params["Agreement End Date"] = end_date
    
    rent_start_date_val = get_value_by_aliases(raw_params, ["Rent Start Date", "Rent Commencement Date"])
    params["Rent Start Date"] = parse_date(rent_start_date_val, start_date)
    
    # Lease Term calculations
    term_months = get_value_by_aliases(raw_params, ["Lease Term Months", "Lease Term", "Term Months", "Duration Months"])
    if term_months is None or pd.isna(term_months):
        term_months = round((end_date - start_date).days / 30.4167)
    else:
        try:
            term_months = int(float(term_months))
        except Exception:
            term_months = round((end_date - start_date).days / 30.4167)
    params["Lease Term Months"] = term_months
    
    params["Rent Per Sqft"] = get_float(raw_params, "Rent Per Sqft", ["Rent Per Sq.ft.", "Rent per Sqft", "Rent Per Sqft/month", "Base Rent", "Quoted Rentals", "Quoted Rentals "], 120.0)
    params["Quoted CAM"] = get_float(raw_params, "Quoted CAM", ["CAM", "CAM Per Sqft", "Quoted CAM (per sq ft/month)", "CAM per Sq ft", "CAM per Sq.ft.", "CAM per Sqft", "CAM Rate"], 15.48)
    
    esc_val = get_value_by_aliases(raw_params, ["Escalation %", "Rent Escalation %", "Escalation Percentage", "Rent Escalation Pct"])
    params["Escalation %"] = float(esc_val) if esc_val is not None and not pd.isna(esc_val) else 0.15
    
    # Robust check for frequency e.g. 0.36 -> 36 months
    freq_val = get_value_by_aliases(raw_params, ["Escalation Frequency Months", "Rent Escalation Frequency Months", "Rent Escalation Frequency", "Escalation Freq"])
    if freq_val is not None and not pd.isna(freq_val):
        try:
            freq_val = float(freq_val)
            if freq_val < 1.0:
                freq_val = int(round(freq_val * 100))
            else:
                freq_val = int(freq_val)
        except Exception:
            freq_val = 36
    else:
        freq_val = 36
    params["Escalation Frequency Months"] = freq_val
    
    cam_esc_val = get_value_by_aliases(raw_params, ["CAM Escalation %", "CAM Escalation Percentage", "CAM Escalation Pct", "CAM Escalation", "CAM escalation %"])
    params["CAM Escalation %"] = float(cam_esc_val) if cam_esc_val is not None and not pd.isna(cam_esc_val) else 0.05
    
    cam_freq_val = get_value_by_aliases(raw_params, ["CAM Escalation Frequency Months", "CAM Escalation Frequency", "CAM Escalation Freq", "CAM Escalation Period", "CAM Escalation Period Months", "escalation period", "CAM escalation period", "CAM escalation frequency"])
    if cam_freq_val is not None and not pd.isna(cam_freq_val):
        try:
            cam_freq_val = float(cam_freq_val)
            if cam_freq_val < 1.0:
                cam_freq_val = int(round(cam_freq_val * 100))
            else:
                cam_freq_val = int(cam_freq_val)
        except Exception:
            cam_freq_val = 12
    else:
        cam_freq_val = 12
    params["CAM Escalation Frequency Months"] = cam_freq_val
    
    # Parking escalation parameters (fallback to Rent Escalation if not specified)
    park_esc_val = get_value_by_aliases(raw_params, ["Parking Escalation %", "Parking Escalation Percentage", "Parking Escalation Pct"])
    if park_esc_val is not None and not pd.isna(park_esc_val):
        params["Parking Escalation %"] = float(park_esc_val)
    else:
        params["Parking Escalation %"] = params["Escalation %"]
        
    park_freq_val = get_value_by_aliases(raw_params, ["Parking Escalation Frequency Months", "Parking Escalation Frequency", "Parking Escalation Freq"])
    if park_freq_val is not None and not pd.isna(park_freq_val):
        try:
            park_freq_val = float(park_freq_val)
            if park_freq_val < 1.0:
                park_freq_val = int(round(park_freq_val * 100))
            else:
                park_freq_val = int(park_freq_val)
        except Exception:
            park_freq_val = params["Escalation Frequency Months"]
    else:
        park_freq_val = params["Escalation Frequency Months"]
    params["Parking Escalation Frequency Months"] = park_freq_val

    params["Billing Frequency"] = get_str(raw_params, "Billing Frequency", ["BillingFrequency"], "Monthly")
    params["Security Deposit Months"] = get_float(raw_params, "Security Deposit Months", ["Security Deposit (number of months)", "Security Deposit (Months)"], 6.0)
    params["Security Deposit Amount"] = get_float(raw_params, "Security Deposit Amount", ["Security Deposit Amount (INR)", "Security Deposit"], 11418000.0)
    params["Refundable Deposit"] = get_str(raw_params, "Refundable Deposit", ["Refundable"], "Yes")
    
    # Capex & Cost factors
    total_fitout = get_float(raw_params, "Fitout Cost", ["Total Fitout Cost", "Fitout Cost Amount", "Fitouts", "Capex"], 34000000.0)
    params["Fitout Cost"] = total_fitout
    params["Useful Life Years"] = get_float(raw_params, "Useful Life Years", ["Useful Life", "Useful Life (Years)"], 5.0)
    params["Residual Value"] = get_float(raw_params, "Residual Value", ["ResidualValue"], 0.0)
    params["Discount Rate"] = get_float(raw_params, "Discount Rate", ["Discounting Rate", "Discount Rate %"], 0.08)
    params["Incremental Borrowing Rate"] = get_float(raw_params, "Incremental Borrowing Rate", ["IBR", "Incremental Borrowing Rate %", "IBR %"], 0.08)
    params["Cost of Capital"] = get_float(raw_params, "Cost of Capital", ["WACC", "Cost of Capital (WACC) %", "WACC %"], 0.105)
    params["Addnl.Deposit -energy(Refundable)"] = get_float(raw_params, "Addnl.Deposit -energy(Refundable)", ["Energy Deposit", "Energy Security Deposit", "Energy Deposit Amount"], 500000.0)
    
    # Add new parameters missing in original but parsed if they exist
    params["Imputed Interest Rate"] = get_float(raw_params, "Imputed Interest Rate", ["Imputed Interest Rate %", "Imputed Interest", "Imputed Rate"], 0.0711)
    params["Ready Reckoner Rate"] = get_float(raw_params, "Ready Reckoner Rate", ["Ready Reckoner Rate (INR/sq m)", "Ready Reckoner"], 15000.0)
    params["Exchange Rate"] = get_float(raw_params, "Exchange Rate", ["Forex Rate", "Exchange Rate (INR/Euro)", "Forex"], 105.02)
    params["Incremental Restoration Cost Sqft"] = get_float(raw_params, "Incremental Restoration Cost Sqft", ["Restoration Cost per Sq ft (ARO)", "Restoration Cost per Sqft", "ARO Rate", "ARO Cost per Sqft", "ARO rate per sq ft"], None)
    
    # OpEx I and OpEx II parameters
    params["Opex Others Per Month"] = get_float(raw_params, "Opex Others Per Month", ["Opex Others", "OpEx Others Rs../ month (At Actuals)", "Opex Others Rs./ month", "Opex Others Rs/ month", "Opex I"], 654.0)
    params["Opex II Per Month"] = get_float(raw_params, "Opex II Per Month", ["Opex II", "OpEx II Rs../ month", "Opex II Rs./ month", "Opex II Rs/ month", "Opex II - OpEx Add-on"], 0.0)
    
    total_pm = get_float(raw_params, "PM Cost Over Lease", ["Preventive Maintenance Cost", "PM Cost", "Maintenance Cost over Lease", "PM cost Over lease period", "PM cost Over lease"], 2500000.0)
    params["PM Cost Over Lease"] = total_pm
    
    return params
