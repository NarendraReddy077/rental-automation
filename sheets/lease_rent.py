import openpyxl
import math

def inject(ws, params):
    """
    Dynamically injects rental and CAM escalation formulas/values into 'Lease Rent' worksheet.
    Also inserts rows if term_months > 72, and updates cross-sheet references in 'Rent Calculation'.
    """
    rent_rate = params.get("Rent Per Sqft", 120.0)
    cam_rate = params.get("Quoted CAM", 15.48)
    rent_esc_pct = params.get("Escalation %", 0.15)
    rent_esc_freq = params.get("Escalation Frequency Months", 36)
    cam_esc_pct = params.get("CAM Escalation %", 0.05)
    cam_esc_freq = params.get("CAM Escalation Frequency Months", 12)
    term_months = params.get("Lease Term Months", 72)
    
    # Calculate how many months we actually need to write
    # If the user-defined lease is longer than 72 months, we dynamically insert rows.
    # The default specimen template has 72 monthly rows (ending at row 102).
    # Month 1 is row 31. Month 72 is row 102.
    # If term_months > 72, we insert (term_months - 72) rows at Row 103 (above original summary row).
    max_months = max(72, term_months)
    if term_months > 72:
        inserted = term_months - 72
        ws.insert_rows(103, amount=inserted)

    # Update starting date and calendar year dynamically
    ws["B31"] = "=Main!B4"
    ws["A31"] = "=YEAR(B31)"

    # Identify the year boundary rows
    # Year 1 ends at row 38 (8 months).
    # Year 2, 3, etc. have 12 months each.
    # Boundary row is when a new year begins.
    # Month 1 to 8: rows 31 to 38 (Year 1). Row 39 starts Year 2. So Row 39 contains Year 1 summary in N/O.
    # Month 9 to 20: rows 39 to 50 (Year 2). Row 51 starts Year 3. So Row 51 contains Year 2 summary in N/O.
    # In general:
    # Year y (for y >= 2) starts at row: start_r = 39 + 12 * (y - 2)
    # Year y ends at row: end_r = start_r + 11
    # Boundary row (which starts Year y+1 and has Year y summary) is: sum_r = end_r + 1
    
    # Let's map summary rows to their respective years:
    # Year 1 summary is at Row 39 (N39 = SUM(K31:K38))
    # Year y summary is at Row 39 + 12 * (y - 1)
    
    summary_rows = {}
    
    # Loop through all months and write formulas
    for m in range(1, max_months + 1):
        r = 30 + m
        
        # Beyond the lease term, set rates to 0.0
        if m > term_months:
            ws.cell(row=r, column=1, value="")   # Column A: Year
            ws.cell(row=r, column=2, value="")   # Column B: From Date
            ws.cell(row=r, column=3, value="")   # Column C: To Date
            ws.cell(row=r, column=4, value="")   # Column D: Months
            ws.cell(row=r, column=5, value="")   # Column E: Area SQF
            ws.cell(row=r, column=6, value="")   # Column F: Area SQM
            ws.cell(row=r, column=7, value="")   # Column G: Escalation
            ws.cell(row=r, column=8, value=0.0)  # Column H: Rental rate
            ws.cell(row=r, column=9, value="")   # Column I: Car Parks
            ws.cell(row=r, column=10, value=0.0) # Column J: CAM rate
            continue
            
        # Write Area columns
        if m == 1:
            ws.cell(row=r, column=5, value="='Rent Calculation'!$F$23")
            ws.cell(row=r, column=6, value="='Rent Calculation'!G23")
            ws.cell(row=r, column=8, value="=Main!B7")
            ws.cell(row=r, column=10, value="=Main!B8")
        else:
            ws.cell(row=r, column=5, value=f"=E{r-1}")
            ws.cell(row=r, column=6, value=f"=F{r-1}")
            
            # Rent Escalation
            if (m - 1) % rent_esc_freq == 0:
                ws.cell(row=r, column=8, value=f"=H{r-1}*(1+{rent_esc_pct})")
            else:
                ws.cell(row=r, column=8, value=f"=H{r-1}")
                
            # CAM Escalation
            if (m - 1) % cam_esc_freq == 0:
                ws.cell(row=r, column=10, value=f"=J{r-1}*(1+{cam_esc_pct})")
            else:
                ws.cell(row=r, column=10, value=f"=J{r-1}")
                
        # Determine if this row is a Year boundary (for years 1 to 10)
        # Year 1 summary is written at row 39 (Month 9)
        # Year 2 summary at row 51 (Month 21)
        # etc.
        # So boundary row is when m = 8 + 12 * (y - 1) + 1 = 9 + 12 * (y - 1)
        if m > 8 and (m - 9) % 12 == 0:
            y = (m - 9) // 12 + 1  # This is the year that just ended
            if y <= 10:
                if y == 1:
                    start_sum = 31
                    end_sum = 38
                else:
                    start_sum = 39 + 12 * (y - 2)
                    end_sum = start_sum + 11
                ws.cell(row=r, column=14, value=f"=SUM(K{start_sum}:K{end_sum})") # Column N: Rentals
                ws.cell(row=r, column=15, value=f"=SUM(L{start_sum}:L{end_sum})") # Column O: CAM
                summary_rows[y] = r

    # Calculate final year number and the final summary row index
    # If term_months <= 8, final year is Year 1.
    if term_months <= 8:
        final_year = 1
        start_sum = 31
        end_sum = 30 + term_months
    else:
        final_year = math.ceil((term_months - 8) / 12) + 1
        if final_year == 1:
            start_sum = 31
        else:
            start_sum = 39 + 12 * (final_year - 2)
        end_sum = 30 + term_months
        
    final_sum_row = 30 + term_months + 1
    
    # Write the final year summary row
    ws.cell(row=final_sum_row, column=1, value="")
    ws.cell(row=final_sum_row, column=2, value="")
    ws.cell(row=final_sum_row, column=14, value=f"=SUM(K{start_sum}:K{end_sum})") # Column N
    ws.cell(row=final_sum_row, column=15, value=f"=SUM(L{start_sum}:L{end_sum})") # Column O
    summary_rows[final_year] = final_sum_row
    
    # Write the grand total row at final_sum_row + 1
    total_row = final_sum_row + 1
    ws.cell(row=total_row, column=11, value=f"=SUM(K31:K{30 + term_months})") # Column K
    ws.cell(row=total_row, column=12, value=f"=SUM(L31:L{30 + term_months})") # Column L
    ws.cell(row=total_row, column=13, value=f"=SUM(M31:M{30 + term_months})") # Column M
    ws.cell(row=total_row, column=14, value=f"=SUM(N31:N{final_sum_row})") # Column N
    ws.cell(row=total_row, column=15, value=f"=SUM(O31:O{final_sum_row})") # Column O

    # Update Rent Calculation sheet references dynamically using ws.parent
    wb = ws.parent
    if "Rent Calculation" in wb.sheetnames:
        rc_ws = wb["Rent Calculation"]
        
        # Years 1 to 10 correspond to Rent Calculation rows 211 to 220
        for y in range(1, 11):
            r_rc = 210 + y
            if y in summary_rows:
                sum_row = summary_rows[y]
                rc_ws.cell(row=r_rc, column=3, value=f"='Lease Rent'!N{sum_row}") # Column C: Rent (INR)
                rc_ws.cell(row=r_rc, column=9, value=f"='Lease Rent'!O{sum_row}") # Column I: CAM (INR)
            else:
                # Inactive year: set to 0.0
                rc_ws.cell(row=r_rc, column=3, value=0.0)
                rc_ws.cell(row=r_rc, column=9, value=0.0)
