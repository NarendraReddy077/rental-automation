import datetime
from sheets.capex_pm import calculate_fy_months

def apply_schedules_and_fallbacks(
    params,
    fitout_breakdown,
    fitout_lifes,
    fitout_active_months_breakdown,
    capex_sched,
    capex_lives,
    capex_active_months_breakdown,
    pm_sched,
    was_breakdown_parsed,
    was_pm_parsed
):
    """Applies fallbacks and scales values for Fitouts, Capex, and PM, updating params in-place."""
    start_date = params["Agreement Start Date"]
    start_year = start_date.year
    term_months = params["Lease Term Months"]
    total_fitout = params["Fitout Cost"]
    total_pm = params["PM Cost Over Lease"]
    
    # If fitout breakdown was not parsed, initialize default 3 phases
    if not fitout_breakdown:
        fitout_breakdown = [14000000.0, 5000000.0, 15000000.0]
        fitout_lifes = [term_months, 30, 48]
        
    if was_breakdown_parsed:
        # Update Fitout Cost to match sum of parsed breakdown
        total_fitout = sum(fitout_breakdown)
        params["Fitout Cost"] = total_fitout
    else:
        # If fitout total from Main doesn't match sum of parsed breakdown, scale it
        if abs(total_fitout - sum(fitout_breakdown)) > 1.0:
            if sum(fitout_breakdown) > 0:
                scale = total_fitout / sum(fitout_breakdown)
                fitout_breakdown = [f * scale for f in fitout_breakdown]
            else:
                num_p = len(fitout_breakdown) if fitout_breakdown else 3
                fitout_breakdown = [total_fitout / num_p] * num_p

    # If fitout_lifes is empty, initialize default lives
    if not fitout_lifes:
        default_lifes = [term_months, 30, 48]
        fitout_lifes = []
        for idx in range(len(fitout_breakdown)):
            if idx < len(default_lifes):
                fitout_lifes.append(default_lifes[idx])
            else:
                fitout_lifes.append(int(params["Useful Life Years"] * 12))

    # Fitout active months fallback
    if not fitout_active_months_breakdown:
        fitout_active_months_breakdown = []
        for life in fitout_lifes:
            fitout_active_months_breakdown.append(
                calculate_fy_months(start_date, max(term_months, 120), life)
            )

    # Capex schedule fallback
    if not capex_sched:
        capex_sched = {
            start_year: 14000000.0,
            start_year + 1: 12000000.0,
            start_year + 2: 5000000.0,
            start_year + 3: 6000000.0,
            start_year + 4: 6000000.0
        }
        # Scale Capex if total Fitout Cost changed
        if abs(total_fitout - 34000000.0) > 1.0:
            scale = total_fitout / 34000000.0
            for k in capex_sched:
                capex_sched[k] *= scale

    # Capex useful lives fallback
    if not capex_lives:
        for yr in sorted(capex_sched.keys()):
            life = max(0, term_months - 12 * (yr - start_year))
            capex_lives[yr] = life

    # Capex active months fallback
    if not capex_active_months_breakdown:
        capex_active_months_breakdown = []
        from sheets.capex_pm import get_capex_tranche_months
        # Filter years to match start_year up to start_year + 9 (maximum 10 tranches)
        for idx in range(10):
            yr = start_year + idx
            life = capex_lives.get(yr, term_months)
            capex_active_months_breakdown.append(
                get_capex_tranche_months(start_date, term_months, idx + 1, life)
            )

    # PM schedule fallback
    if not pm_sched:
        pm_sched = {
            start_year: 1500000.0,
            start_year + 1: 200000.0,
            start_year + 2: 200000.0,
            start_year + 3: 200000.0,
            start_year + 4: 200000.0,
            start_year + 5: 200000.0
        }
        # Scale PM if total PM cost changed
        if abs(total_pm - 2500000.0) > 1.0:
            scale = total_pm / 2500000.0
            for k in pm_sched:
                pm_sched[k] *= scale
                
    if was_pm_parsed:
        total_pm = sum(pm_sched.values())
        params["PM Cost Over Lease"] = total_pm
                
    params["Fitout Cost Breakdown"] = fitout_breakdown
    params["Fitout Useful Lives"] = fitout_lifes
    params["Fitout Active Months Breakdown"] = fitout_active_months_breakdown
    params["Capex Schedule"] = capex_sched
    params["Capex Useful Lives"] = capex_lives
    params["Capex Active Months Breakdown"] = capex_active_months_breakdown
    params["PM Schedule"] = pm_sched

    return params
