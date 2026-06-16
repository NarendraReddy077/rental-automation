from sheets.input_parser.utils import resolve_file_bytes
from sheets.input_parser.main_parser import parse_main_sheet, extract_main_parameters
from sheets.input_parser.sheet_parser import parse_extra_sheets, parse_schedules_from_main_fallback
from sheets.input_parser.fallbacks import apply_schedules_and_fallbacks

def parse_input_xlsx(file_bytes):
    """Parses user-uploaded input Excel sheet with extensive error resilience."""
    try:
        raw_data = resolve_file_bytes(file_bytes)
        raw_params = parse_main_sheet(raw_data)
        params = extract_main_parameters(raw_params)
        
        try:
            (
                fitout_breakdown,
                fitout_lifes,
                fitout_active_months_breakdown,
                capex_sched,
                capex_lives,
                capex_active_months_breakdown,
                pm_sched,
                has_capex_pm_sheet
            ) = parse_extra_sheets(raw_data, params)
            
            start_year = params["Agreement Start Date"].year if "Agreement Start Date" in params and params["Agreement Start Date"] else 2026
            was_breakdown_parsed = (len(fitout_breakdown) > 0)
            was_pm_parsed = (len(pm_sched) > 0)
            
            if not has_capex_pm_sheet:
                # Fallback parsing if CAPEX and PM sheet is missing
                f_breakdown, c_sched, p_sched = parse_schedules_from_main_fallback(raw_params, start_year)
                fitout_breakdown = f_breakdown
                capex_sched = c_sched
                pm_sched = p_sched
                was_breakdown_parsed = (len(fitout_breakdown) > 0)
                was_pm_parsed = (len(pm_sched) > 0)
        except Exception:
            # If openpyxl parsing fails, fallback to regex parsing from main
            start_year = params.get("Agreement Start Date").year if params.get("Agreement Start Date") else 2026
            fitout_breakdown, capex_sched, pm_sched = parse_schedules_from_main_fallback(raw_params, start_year)
            fitout_lifes = []
            fitout_active_months_breakdown = []
            capex_lives = {}
            capex_active_months_breakdown = []
            was_breakdown_parsed = (len(fitout_breakdown) > 0)
            was_pm_parsed = (len(pm_sched) > 0)

        # Apply fallback defaults and scaling
        params = apply_schedules_and_fallbacks(
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
        )
        
        return params, None
    except Exception as e:
        return {}, str(e)
