import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

st.set_page_config(page_title="OLT Inventory & Tracker Studio", page_icon="📊", layout="wide")

st.title("📊 OLT Inventory & Tracker Studio")
st.markdown("Upload your datasets below to generate an optimized, duplicate-free Master Tracker and automated OLT Inventory Summary.")

# -----------------------------------------------------------------------------
# File Uploaders
# -----------------------------------------------------------------------------
col1, col2 = st.columns(2)
with col1:
    tracker_file = st.file_uploader("📂 Upload Existing Tracker", type=["csv", "xlsx"])
with col2:
    add_file = st.file_uploader("➕ Upload 'Data to Add to Tracker'", type=["csv", "xlsx"])

def load_data(uploaded_file):
    if uploaded_file is not None:
        filename = uploaded_file.name.lower()
        
        # Scenario A: Excel Formats (Binary, don't use string encodings)
        if filename.endswith('.xlsx') or filename.endswith('.xls'):
            df_raw = pd.read_excel(uploaded_file, header=None)
            
            header_row_idx = 0
            for idx, row in df_raw.iterrows():
                row_str = [str(x).strip().lower() for x in row.values]
                if 'plaid' in row_str or 'project tagging' in row_str or 'project or program' in row_str:
                    header_row_idx = idx
                    break
                    
            uploaded_file.seek(0)
            return pd.read_excel(uploaded_file, header=header_row_idx)
            
        # Scenario B: CSV Formats (Text based, requires fallback encodings)
        else:
            encodings_to_try = ['utf-8', 'cp1252', 'latin1', 'utf-8-sig']
            df_raw = None
            successful_encoding = None
            
            for enc in encodings_to_try:
                try:
                    uploaded_file.seek(0)
                    df_raw = pd.read_csv(uploaded_file, header=None, encoding=enc)
                    successful_encoding = enc
                    break
                except (UnicodeDecodeError, LookupError):
                    continue
            
            if df_raw is None:
                raise UnicodeDecodeError("Could not decode the CSV file using standard character sets (UTF-8, CP1252, Latin1).")
                
            header_row_idx = 0
            for idx, row in df_raw.iterrows():
                row_str = [str(x).strip().lower() for x in row.values]
                if 'plaid' in row_str or 'project tagging' in row_str or 'project or program' in row_str:
                    header_row_idx = idx
                    break
            
            uploaded_file.seek(0)
            return pd.read_csv(uploaded_file, header=header_row_idx, encoding=successful_encoding)
            
    return None

if tracker_file and add_file:
    with st.spinner("Processing files and adjusting data encodings safely..."):
        try:
            df_tracker = load_data(tracker_file)
            df_add = load_data(add_file)
        except Exception as e:
            st.error(f"❌ Read Error: {str(e)}")
            st.stop()

    # Data Cleaning: Strip whitespaces from column names
    df_tracker.columns = df_tracker.columns.astype(str).str.strip()
    df_add.columns = df_add.columns.astype(str).str.strip()

    # Standardize names to match key targets
    for df in [df_tracker, df_add]:
        if 'Site name' in df.columns and 'Site Name' not in df.columns:
            df.rename(columns={'Site name': 'Site Name'}, inplace=True)

    # -----------------------------------------------------------------------------
    # Processing Engine: Tracker Update Logic & Discrepancy Matching
    # -----------------------------------------------------------------------------
    st.subheader("⚙️ Processing Engine Status")
    
    if 'PLAID' in df_add.columns and 'PLAID' in df_tracker.columns and 'Site Name' in df_tracker.columns and 'Site Name' in df_add.columns:
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
                match_existing = df_tracker[(df_tracker['PLAID'] == key[0]) & (df_tracker['Site Name'] == key[1])]
                
                is_identical = False
                for _, ext_row in match_existing.iterrows():
                    common_cols = [c for c in df_add.columns if c in df_tracker.columns and c not in ['Row_Origin_Color', 'SN']]
                    if row[common_cols].equals(ext_row[common_cols]):
                        is_identical = True
                        break
                
                if is_identical:
                    continue  # Skip 100% duplicate entries
                else:
                    row['Row_Origin_Color'] = 'Discrepancy_Blue'
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
        st.error("❌ Key identifier column 'PLAID' or 'Site Name' missing from datasets. Please check headers.")
        df_master_tracker = df_tracker.copy()

    # -----------------------------------------------------------------------------
    # Dynamic Column Dynamic Mapping Engine for OLT Summary
    # -----------------------------------------------------------------------------
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
            df_olt_summary[targeted_col] = "" 
            
    df_olt_summary.insert(0, 'No.', range(1, 1 + len(df_olt_summary)))

    # Preview Tabs
    tab1, tab2 = st.tabs(["📋 Updated Master Tracker Preview", "🖥️ OLT Inventory Summary Preview"])
    with tab1:
        st.dataframe(df_master_tracker.drop(columns=['Row_Origin_Color'], errors='ignore').head(50))
    with tab2:
        st.dataframe(df_olt_summary.head(50))

    # -----------------------------------------------------------------------------
    # Advanced Workbook Generation (OpenPyXL Styler)
    # -----------------------------------------------------------------------------
    output_buffer = io.BytesIO()
    wb = Workbook()
    
    # Sheet 1: OLT Inventory Summary
    ws_olt = wb.active
    ws_olt.title = "OLT Summary"
    ws_olt.views.sheetView[0].showGridLines = True
    
    # Sheet 2: Master Tracker
    ws_track = wb.create_sheet(title="Tracker")
    ws_track.views.sheetView[0].showGridLines = True
    
    # Styling Parameters
    navy_header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    white_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    regular_font = Font(name="Segoe UI", size=10)
    
    discrepancy_blue_fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid") 
    original_green_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")   
    
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
    )
    
    # Build Sheet 1: OLT Summary
    ws_olt.append(list(df_olt_summary.columns))
    for col_num, header in enumerate(df_olt_summary.columns, 1):
        cell = ws_olt.cell(row=1, column=col_num)
        cell.fill = navy_header_fill
        cell.font = white_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    
    for _, row in df_olt_summary.iterrows():
        ws_olt.append([str(v) if pd.notna(v) else "" for v in row])
        
    for r_idx in range(2, ws_olt.max_row + 1):
        for c_idx in range(1, ws_olt.max_column + 1):
            cell = ws_olt.cell(row=r_idx, column=c_idx)
            cell.font = regular_font
            cell.border = thin_border
            if c_idx in [1, 4, 5, 7, 8]: 
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
        ws_track.append([str(v) if pd.notna(v) else "" for v in row])
        
    for r_idx, color_type in enumerate(color_series, start=2):
        for c_idx in range(1, ws_track.max_column + 1):
            cell = ws_track.cell(row=r_idx, column=c_idx)
            cell.font = regular_font
            cell.border = thin_border
            
            if color_type == 'Discrepancy_Blue':
                cell.fill = discrepancy_blue_fill
            elif color_type == 'Original_Green':
                cell.fill = original_green_fill

    # Auto-adjust column sizing
    for ws in [ws_olt, ws_track]:
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = col[0].column_letter
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    wb.save(output_buffer)
    
    # Download Presentation Tier
    st.markdown("---")
    st.subheader("📥 Download Generated Sheets")
    st.download_button(
        label="🚀 Download Consolidated Report Workbook (.xlsx)",
        data=output_buffer.getvalue(),
        file_name="Master_OLT_Inventory_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )