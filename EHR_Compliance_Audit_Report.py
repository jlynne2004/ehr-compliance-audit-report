import pandas as pd
import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows

# ==========================================
# 1. LOAD AND AUDIT MESSY EHR DATA
# ==========================================
# Load the messy CSV data
df = pd.read_csv("messy_ehr_export.csv")

# Convert date columns to datetime objects for mathematical comparison
df['admission_date'] = pd.to_datetime(df['admission_date'], errors='coerce')
df['discharge_date'] = pd.to_datetime(df['discharge_date'], errors='coerce')

# Initialize lists to hold our audit tally findings
audit_tally = []

# Loop through records to perform row-by-row data quality checks
for idx, row in df.iterrows():
    patient = row['patient_id']
    
    # Check 1: Missing or Malformed NPI (Must be exactly 10 digits)
    npi = str(row['provider_npi']).strip().split('.')[0] # Clean float strings if any
    if pd.isna(row['provider_npi']) or len(npi) != 10 or not npi.isdigit():
        audit_tally.append({
            "Patient ID": patient, "Field": "provider_npi", 
            "Issue Type": "Invalid/Missing NPI", "Current Value": row['provider_npi'],
            "Action Required": "Verify provider NPI profile and update registry."
        })
        
    # Check 2: Date Logic Error (Discharge before Admission)
    if pd.notna(row['admission_date']) and pd.notna(row['discharge_date']):
        if row['discharge_date'] < row['admission_date']:
            audit_tally.append({
                "Patient ID": patient, "Field": "discharge_date", 
                "Issue Type": "Timeline Discrepancy", 
                "Current Value": f"Adm: {row['admission_date'].strftime('%Y-%m-%d')} | Dis: {row['discharge_date'].strftime('%Y-%m-%d')}",
                "Action Required": "Review chart timeline; correct clinical dates."
            })
            
    # Check 3: Missing Required Clinical Coding for Compliance
    if pd.isna(row['icd_10_code']) or str(row['icd_10_code']).strip() == "":
        audit_tally.append({
            "Patient ID": patient, "Field": "icd_10_code", 
            "Issue Type": "Missing Billing Code", "Current Value": "BLANK",
            "Action Required": "Route back to medical coding team for chart review."
        })

# Convert audit findings into a clean dataframe
df_tally = pd.DataFrame(audit_tally)

# ==========================================
# 2. GENERATE POLISHED EXCEL WORKBOOK
# ==========================================
wb = Workbook()

# Styling Definitions
font_title = Font(name="Calibri", size=16, bold=True, color="1B365D")
font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
font_body = Font(name="Calibri", size=11, bold=False)
font_kpi_num = Font(name="Calibri", size=20, bold=True, color="1B365D")

fill_header = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
fill_zebra = PatternFill(start_color="F4F7FA", end_color="F4F7FA", fill_type="solid")
fill_error = PatternFill(start_color="FADBD8", end_color="FADBD8", fill_type="solid") # Soft Red

thin_border = Border(
    left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
)

# ------------------------------------------
# TAB 1: EXECUTIVE DASHBOARD
# ------------------------------------------
ws1 = wb.active
ws1.title = "Executive Summary"
ws1.views.sheetView[0].showGridLines = False # Remove default gridlines

# Title Block
ws1["A1"] = "EHR Compliance & Data Quality Executive Report"
ws1["A1"].font = font_title
ws1["A2"] = f"Report Generated: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}"
ws1["A2"].font = Font(name="Calibri", size=10, italic=True, color="595959")

# Construct KPI Blocks
ws1["A4"] = "Total Records Audited"
ws1["A5"] = len(df)
ws1["B4"] = "Total Issues Detected"
ws1["B5"] = len(df_tally)

for col in ["A", "B"]:
    ws1[f"{col}4"].font = Font(name="Calibri", size=10, bold=True, color="595959")
    ws1[f"{col}5"].font = font_kpi_num
    ws1[f"{col}4"].alignment = Alignment(horizontal="center")
    ws1[f"{col}5"].alignment = Alignment(horizontal="center")
    ws1[f"{col}4"].fill = fill_zebra
    ws1[f"{col}5"].fill = fill_zebra

# ------------------------------------------
# TAB 2: THE ACTIONABLE TALLY SHEET
# ------------------------------------------
ws2 = wb.create_sheet(title="Data Quality Tally")
ws2.views.sheetView[0].showGridLines = True # Keep gridlines on the detailed sheet

# Append Headers to Tab 2
headers = list(df_tally.columns)
ws2.append(headers)

# Apply header styles
for col_num, header in enumerate(headers, 1):
    cell = ws2.cell(row=1, column=col_num)
    cell.font = font_header
    cell.fill = fill_header
    cell.alignment = Alignment(horizontal="left", vertical="center")

# Append row data and apply professional styling formatting
for row_idx, row_data in enumerate(dataframe_to_rows(df_tally, index=False, header=False), 2):
    ws2.append(row_data)
    for col_num in range(1, len(row_data) + 1):
        cell = ws2.cell(row=row_idx, column=col_num)
        cell.font = font_body
        cell.border = thin_border
        
        # Soft red fill for the specific error column to draw focus
        if col_num == 3: # Issue Type Column
            cell.fill = fill_error
        elif row_idx % 2 == 0: # Soft zebra striping for readability
            cell.fill = fill_zebra

# Freeze headers on data log so they scroll nicely
ws2.freeze_panes = "A2"

# Auto-fit columns to avoid truncated text or ### errors
for ws in [ws1, ws2]:
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = col[0].column_letter
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

# Save the final automated file
wb.save("EHR_Compliance_Audit_Report.xlsx")
print("Process complete. Polished Excel report generated.")
