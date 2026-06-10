import datetime

def inject(ws, params):
    """
    Injects parameters into the 'Main' worksheet of the openpyxl workbook.
    """
    ws["B2"] = params["REU Name"]
    ws["B3"] = params["Chargeable Area Sqft"]
    ws["C3"] = "=B3/10.764"  # Area in sq.m.
    
    # Parse and write dates
    start_date = params["Agreement Start Date"]
    end_date = params["Agreement End Date"]
    ws["B4"] = start_date
    ws["B5"] = end_date
    ws["B6"] = "=(B5-B4)/365"
    
    ws["B7"] = params["Rent Per Sqft"]
    ws["B8"] = params["Quoted CAM"]
    ws["B9"] = params["Security Deposit Amount"]
    ws["B10"] = params["Fitout Cost"]
    ws["B11"] = params["PM Cost Over Lease"]
    
    # Imputed interest
    ws["B13"] = params["Imputed Interest Rate"]
    
    # Text clauses
    rent_esc_years = params["Escalation Frequency Months"] // 12
    ws["B14"] = f"{int(params['Escalation %'] * 100)}% every {rent_esc_years} years"
    cam_esc_freq = params.get("CAM Escalation Frequency Months", 12)
    if cam_esc_freq % 12 == 0:
        cam_esc_years = cam_esc_freq // 12
        if cam_esc_years == 1:
            ws["B15"] = f"{int(params['CAM Escalation %'] * 100)}% every year"
        else:
            ws["B15"] = f"{int(params['CAM Escalation %'] * 100)}% every {cam_esc_years} years"
    else:
        ws["B15"] = f"{int(params['CAM Escalation %'] * 100)}% every {cam_esc_freq} months"
    
    ws["B12"] = params.get("Cost of Capital", 0.105)
    ws["B16"] = params.get("4 Wheeler Slots", 80)
    ws["B17"] = params.get("4 Wheeler Rate", 1500.0)
    ws["B18"] = params.get("2 Wheeler Slots", 50)
    ws["B19"] = params.get("2 Wheeler Rate", 1000.0)

def simulate(params):
    """
    Simulates the Main sheet values for the Streamlit UI preview.
    Returns a dictionary of key-value pairs formatted nicely.
    """
    start_date = params["Agreement Start Date"]
    end_date = params["Agreement End Date"]
    duration_yrs = (end_date - start_date).days / 365.0
    
    return {
        "REU Name": params["REU Name"],
        "Chargeable Area (Sq.ft.)": f"{int(params['Chargeable Area Sqft']):,}",
        "Chargeable Area (Sq.m.)": f"{params['Chargeable Area Sqft'] / 10.764:,.2f}",
        "Lease Commencement Date": start_date.strftime("%d-%b-%Y"),
        "Lease Expiry Date": end_date.strftime("%d-%b-%Y"),
        "Lease Term (Years)": f"{duration_yrs:.2f}",
        "Quoted Rentals (per Sqft/mo)": f"{params['Currency']} {params['Rent Per Sqft']:.2f}",
        "Quoted CAM (per Sqft/mo)": f"{params['Currency']} {params['Quoted CAM']:.2f}",
        "Security Deposit Amount": f"{params['Currency']} {params['Security Deposit Amount']:,.2f}",
        "Capex Cost": f"{params['Currency']} {params['Fitout Cost']:,.2f}",
        "PM Cost Over Lease Period": f"{params['Currency']} {params['PM Cost Over Lease']:,.2f}",
        "Imputed Interest Rate": f"{params['Imputed Interest Rate']*100:.2f}%",
        "Rent Escalation Clause": f"{int(params['Escalation %']*100)}% every {params['Escalation Frequency Months']//12} years",
        "CAM Escalation Clause": (
            f"{int(params['CAM Escalation %']*100)}% every year" if params.get("CAM Escalation Frequency Months", 12) == 12 else
            (f"{int(params['CAM Escalation %']*100)}% every {params.get('CAM Escalation Frequency Months', 12)//12} years" if params.get("CAM Escalation Frequency Months", 12) % 12 == 0 else
             f"{int(params['CAM Escalation %']*100)}% every {params.get('CAM Escalation Frequency Months', 12)} months")
        ),
        "Energy Deposit": f"{params['Currency']} {params['Addnl.Deposit -energy(Refundable)']:,.2f}",
        "Incremental Restoration Cost / sqft": f"{params['Currency']} {params['Incremental Restoration Cost Sqft']:.2f}" if params.get("Incremental Restoration Cost Sqft") is not None else ""
    }
