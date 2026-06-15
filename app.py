import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

st.set_page_config(page_title="OLT Inventory & Tracker Studio", page_icon="📊", layout="wide")

# App Header
st.title("📊 OLT Inventory & Tracker Studio")
st.markdown("Upload your existing data files below to generate an optimized, duplicate-free Master Tracker and automated OLT Inventory Summary.")

# -----------------------------------------------------------------------------
# File Uploaders
# -----------------------------------------------------------------------------
col1, col2 = st.columns(2)
with col1:
    tracker_file = st.file_uploader("📂 Upload Existing Tracker (CSV or Excel)", type=["csv", "xlsx"])
with col2:
    add_file = st.file_uploader("➕ Upload 'Data to Add to Tracker' (CSV or Excel)", type=["csv", "xlsx"])

def load_data(uploaded_file, header_row=0):
    if uploaded_file is not None:
        if uploaded_file.name.endswith('.csv'):
            return pd.read_csv(uploaded_file, header=header_row)
        else:
            return pd.read_excel(uploaded_file, header=header_row)
    return None

if tracker_file and add_file:
    # Load files (Handles multi-row metadata offsets common in engineering sheets)
    try:
        # Looking at tracker structure, row index 1 contains the headers
        df_tracker = load_data(tracker_file, header_row=1)
        df_add = load_data(add_file, header_row=2)
    except Exception as e:
        df_tracker = load_data(tracker_file, header_row=0)
        df_add = load_data(add_file, header_row=0)

    # Data Cleaning: Strip whitespaces from column names
    df_tracker.columns = df_tracker.columns.astype(str).str.strip()
    df_add.columns = df_add.columns.astype(str).str.strip()

    # -----------------------------------------------------------------------------
    # Processing Engine: Tracker Update Logic & Discrepancy Matching
    # -----------------------------------------------------------------------------
    st.subheader("⚙️ Processing Engine Status")
    
    # Harmonize naming conventions between sheets if they slightly differ
    if 'PLAID' in df_add.columns and 'PLAID' in df_tracker.columns:
        df_tracker['PLAID'] = df_tracker['PLAID'].astype(str).str.strip()
        df_add['PLAID'] = df_add['PLAID'].astype(str).str.strip()
        df_tracker['Site Name'] = df_tracker['Site Name'].astype(str).str.strip()
        df_add['Site Name'] = df_add['Site Name'].astype(str).str.strip()
        
        # Track highlighting labels
        df_tracker['Row_Origin_Color'] = 'Original'
        df_add['Row_Origin_Color'] = 'New_Update'
        
        # Standardize matching keys
        tracker_keys = set(zip(df_tracker['PLAID'], df_tracker['Site Name']))
        
        cleaned_add_rows = []
        for _, row in df_add.iterrows():
            key = (str(row.get('PLAID', '')).strip(), str(row.get('Site Name', '')).strip())
            
            if key in tracker_keys:
                # Find matching row in existing tracker to test for 100% identity match
                match_existing = df_tracker[(df_tracker['PLAID'] == key[0]) & (df_tracker['Site Name'] == key[1])]
                
                # Check discrepancy vs perfect duplicate
                is_identical = False
                for _, ext_row in match_existing.iterrows():
                    # Compare columns in common
                    common_cols = [c for c in df_add.columns if c in df_tracker.columns and c not in ['Row_Origin_Color', 'SN']]
                    if row[common_cols].equals(ext_row[common_cols]):
                        is_identical = True
                        break
                
                if is_identical:
                    continue  # Skip 100% duplicate entries
                else:
                    row['Row_Origin_Color'] = 'Discrepancy_Blue'
                    # Mark parent row in tracker to be highlighted green
                    df_tracker.loc[(df_tracker['PLAID'] == key[0]) & (df_tracker['Site Name'] == key[1]), 'Row_Origin_Color'] = 'Original_Green'
                    cleaned_add_rows.append(row)
            else:
                cleaned_add_rows.append(row)
                
        if cleaned_add_rows:
            df_add_filtered = pd.DataFrame(cleaned_add_rows)
            df_master_tracker = pd.concat([df_tracker, df_add_filtered], ignore_index=True)
        else:
            df_master_tracker = df_tracker.copy()
            
        st.success("✅ Deduplication & discrepancy checking successfully compiled!")
    else:
        st.error("❌ Key identifier column 'PLAID' or 'Site Name' missing from datasets.")
        df_master_tracker = df_tracker.copy()

    # -----------------------------------------------------------------------------
    # Dynamic Column Dynamic Mapping Engine for OLT Summary
    # -----------------------------------------------------------------------------
    olt_columns_requested = [
        "Region", "Project Type", "Site Name", "PLAID", "Equipment Type",
        "No. of Chassis", "No. of Cards", "Site Status", "Site Survey Actual Date",
        "Installation Done Actual Date", "Powertapping Done Actual Date", 
        "Integration Done Actual Date", "PAT Done Actual Date", 
        "PAC Approval Actual Date", "FAC Approval Actual Date"
    ]
    
    # Mapping variants found across input forms to guarantee seamless joins
    mapping_dictionary = {
        "Region": ["Region"],
        "Project Type": ["Project Type", "PROJECT or PROGRAM"],
        "Site Name": ["Site Name"],
        "PLAID": ["PLAID"],
        "Equipment Type": ["Equipment Type", "Electronics Equipment", "OLT Scope"],
        "No. of Chassis": ["No. of Chassis", "No. of chassis"],
        "No. of Cards": ["No. of Cards", "Number of Cards", "Number \nof Cards"],
        "Site Status": ["Site Status", "Scope Status", "Status"],
        "Site Survey Actual Date": ["Site Survey", "Survey Date", "Site Survey Actual date"],
        "Installation Done Actual Date": ["Installation done", "Installed date", "Installation done actual date"],
        "Powertapping Done Actual Date": ["Powertapping done", "POWERTAPPED DATE", "Powertapping done actual date"],
        "Integration Done Actual Date": ["Integration done", "Integrated Date", "Integration done actual date"],
        "PAT Done Actual Date": ["PAT Done", "PAT", "PAT'ed", "Pat done actual date"],
        "PAC Approval Actual Date": ["PAC Approval", "PAC", "PAC'ed", "PAC approval actual date"],
        "FAC Approval Actual Date": ["FAC Approval", "FAC", "FAC'ed", "FAC approval actual date"]
    }
    
    df_olt_summary = pd.DataFrame()
    
    for targeted_col, aliases in mapping_dictionary.items():
        matched = False
        for alias in aliases:
            if alias in df_master_tracker.columns:
                df_olt_summary[targeted_col] = df_master_tracker[alias]
                matched = True
                break
        if not matched:
            df_olt_summary[targeted_col] = "" # Fallback structural fill if missing
            
    # Clean up Index columns
    df_olt_summary.insert(0, 'No.', range(1, 1 + len(df_olt_summary)))

    # Preview Tabs
    tab1, tab2 = st.tabs(["📋 Updated Master Tracker Preview", "🖥️ OLT Inventory Summary Preview"])
    with tab1:
        st.dataframe(df_master_tracker.drop(columns=['Row_Origin_Color'], errors='ignore').head(100))
    with tab2:
        st.dataframe(df_olt_summary.head(100))

    # -----------------------------------------------------------------------------
    # Advanced Workbook Generation (OpenPyXL Styler)
    # -----------------------------------------------------------------------------
    output_buffer = io.BytesIO()
    
    # Constructing stylized workbook instance
    wb = Workbook()
    
    # Sheet 1: OLT Inventory Summary
    ws_olt = wb.active
    ws_olt.title = "OLT Inventory Summary"
    ws_olt.views.sheetView[0].showGridLines = True
    
    # Sheet 2: Master Tracker
    ws_track = wb.create_sheet(title="Master Tracker")
    ws_track.views.sheetView[0].showGridLines = True
    
    # Theme Design Definitions (Classic Corporate Navy & Soft Alerts)
    navy_header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    white_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    regular_font = Font(name="Segoe UI", size=10)
    
    discrepancy_blue_fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid") # Soft Blue
    original_green_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")   # Soft Green
    
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    
    # Build Sheet 1: OLT Summary
    ws_olt.append(list(df_olt_summary.columns))
    for col_num, header in enumerate(df_olt_summary.columns, 1):
        cell = ws_olt.cell(row=1, column=col_num)
        cell.fill = navy_header_fill
        cell.font = white_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    
    for _, row in df_olt_summary.iterrows():
        ws_olt.append(list(row))
        
    for r_idx in range(2, ws_olt.max_row + 1):
        for c_idx in range(1, ws_olt.max_column + 1):
            cell = ws_olt.cell(row=r_idx, column=c_idx)
            cell.font = regular_font
            cell.border = thin_border
            if c_idx in [1, 4, 5, 7, 8]: # Numeric or clean short codes center-aligned
                cell.alignment = Alignment(horizontal="center")

    # Build Sheet 2: Master Tracker
    track_export_df = df_master_tracker.copy()
    color_series = track_export_df['Row_Origin_Color'].fillna('')
    track_export_df = track_export_df.drop(columns=['Row_Origin_Color'], errors='ignore')
    
    ws_track.append(list(track_export_df.columns))
    for col_num, header in enumerate(track_export_df.columns, 1):
        cell = ws_track.cell(row=1, column=col_num)
        cell.fill = navy_header_fill
        cell.font = white_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        
    for _, row in track_export_df.iterrows():
        # Sanitize NaN/Null for neat spreadsheet formatting
        ws_track.append([str(v) if pd.notna(v) else "" for v in row])
        
    for r_idx, color_type in enumerate(color_series, start=2):
        for c_idx in range(1, ws_track.max_column + 1):
            cell = ws_track.cell(row=r_idx, column=c_idx)
            cell.font = regular_font
            cell.border = thin_border
            
            # Apply color strategy mapping rules directly to row cells
            if color_type == 'Discrepancy_Blue':
                cell.fill = discrepancy_blue_fill
            elif color_type == 'Original_Green':
                cell.fill = original_green_fill

    # Dynamic Column Width Adjustments
    for ws in [ws_olt, ws_track]:
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = col[0].column_letter
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    wb.save(output_buffer)
    
    # -----------------------------------------------------------------------------
    # Download Presentation Tier
    # -----------------------------------------------------------------------------
    st.markdown("---")
    st.subheader("📥 Download Generated Sheets")
    st.download_button(
        label="🚀 Download Consolidated Report Workbook (.xlsx)",
        data=output_buffer.getvalue(),
        file_name="Master_OLT_Inventory_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )