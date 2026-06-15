import streamlit as st
import pandas as pd
import string
import io
import openpyxl

st.set_page_config(page_title="OLT Tracker Tool", layout="wide")

st.title("📊 Master Tracker → Nokia OLT Rollout Tool")

# -----------------------------
# ✅ Core Helper Functions
# -----------------------------

def clean_string_normalization(val) -> str:
    """Normalize a string to single-spaced lower-case without punctuation."""
    if pd.isna(val):
        return ""
    text = str(val).strip().replace("\n", " ").replace("\xa0", " ")
    text = text.translate(str.maketrans('', '', string.punctuation))
    return " ".join(text.lower().split())

def find_dynamic_header_row(xls: pd.ExcelFile, sheet_name: str, lookahead_rows: int = 50) -> int:
    """Scans the beginning of an Excel sheet to guess the real structural header row."""
    preview_df = xls.parse(sheet_name, nrows=lookahead_rows, header=None)
    structural_anchors = {"plaid", "site name", "project", "build year", "equipment type", "sn"}
    
    for idx, row in preview_df.iterrows():
        row_values = [clean_string_normalization(val) for val in row.values]
        if any(anchor in row_values for anchor in structural_anchors):
            return int(idx)
    return 0 

def clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Normalizes and safely deduplicates DataFrame column names dynamically."""
    cleaned = []
    seen_counts = {}
    
    for col in df.columns:
        col_str = str(col).strip().replace("\n", " ").replace("\xa0", " ")
        col_str = col_str.translate(str.maketrans('', '', string.punctuation))
        final_name = " ".join(col_str.split())
        
        if final_name in seen_counts:
            seen_counts[final_name] += 1
            final_name = f"{final_name}_{seen_counts[final_name]}"
        else:
            seen_counts[final_name] = 0
            
        cleaned.append(final_name)
        
    df.columns = cleaned
    return df

def find_column(columns, keywords):
    for col in columns:
        col_str = str(col).lower()
        for key in keywords:
            if key.lower() in col_str:
                return col
    return None

def detect_sheet(xls: pd.ExcelFile, keywords: list) -> str:
    for sheet in xls.sheet_names:
        name = sheet.lower()
        if any(k in name for k in keywords):
            return sheet
    return xls.sheet_names[0]

# -----------------------------
# ✅ File Upload Blocks
# -----------------------------

master_file = st.file_uploader("Upload Master Tracker (Data File)", type=["xlsx"])
olt_file = st.file_uploader("Upload Nokia OLT Tracker (Rollout)", type=["xlsx"])

if master_file and olt_file:
    master_bytes = master_file.read()
    olt_bytes = olt_file.read()
    
    master_xls = pd.ExcelFile(io.BytesIO(master_bytes))
    olt_xls = pd.ExcelFile(io.BytesIO(olt_bytes))

    # -----------------------------
    # 🛠️ Sidebar Configuration & Controls
    # -----------------------------
    st.sidebar.header("🛠️ Configuration Controls")
    
    auto_master_sheet = detect_sheet(master_xls, ["master", "luzon", "file"])
    auto_olt_sheet = detect_sheet(olt_xls, ["rollout", "nokia", "olt", "summary", "inventory"])
    
    selected_master_sheet = st.sidebar.selectbox("Master Sheet Name", master_xls.sheet_names, index=master_xls.sheet_names.index(auto_master_sheet))
    selected_olt_sheet = st.sidebar.selectbox("OLT Sheet Name", olt_xls.sheet_names, index=olt_xls.sheet_names.index(auto_olt_sheet))

    auto_master_idx = find_dynamic_header_row(master_xls, selected_master_sheet)
    auto_olt_idx = find_dynamic_header_row(olt_xls, selected_olt_sheet)
    
    master_header_idx = st.sidebar.number_input("Master Header Row Index (1-based)", min_value=1, value=auto_master_idx + 1) - 1
    olt_header_idx = st.sidebar.number_input("OLT Header Row Index (1-based)", min_value=1, value=auto_olt_idx + 1) - 1

    # Parse structural frames
    master_df = master_xls.parse(selected_master_sheet, header=master_header_idx)
    olt_df = olt_xls.parse(selected_olt_sheet, header=olt_header_idx)

    # Save absolute original raw headers for rebuilding the output file later
    orig_master_cols = list(master_df.columns)
    orig_olt_cols = list(olt_df.columns)

    # Cleaned schemas for background mapping evaluation
    master_df_cleaned = clean_columns(master_df.copy())
    olt_df_cleaned = clean_columns(olt_df.copy())

    # Find the critical linking anchors
    clean_m_plaid = find_column(master_df_cleaned.columns, ["plaid"])
    clean_o_plaid = find_column(olt_df_cleaned.columns, ["plaid"])

    def get_index_fallback(clean_target, original_list, clean_list):
        if clean_target in clean_list:
            return clean_list.index(clean_target)
        return 0

    st.sidebar.subheader("🎯 Primary Identifier Validation")
    m_plaid_idx = get_index_fallback(clean_m_plaid, orig_master_cols, list(master_df_cleaned.columns))
    chosen_master_plaid_raw = st.sidebar.selectbox("Master PLAID Column Identifier", orig_master_cols, index=m_plaid_idx)
    
    o_plaid_idx = get_index_fallback(clean_o_plaid, orig_olt_cols, list(olt_df_cleaned.columns))
    chosen_olt_plaid_raw = st.sidebar.selectbox("OLT PLAID Column Identifier", orig_olt_cols, index=o_plaid_idx)

    # Map selected user boundaries down to operational strings
    master_plaid_col_clean = list(master_df_cleaned.columns)[orig_master_cols.index(chosen_master_plaid_raw)]
    olt_plaid_col_clean = list(olt_df_cleaned.columns)[orig_olt_cols.index(chosen_olt_plaid_raw)]

    # -----------------------------
    # 📉 Missing Records Analysis Block
    # -----------------------------

    st.write(f"📂 **Active Master Sheet:** `{selected_master_sheet}` (Header Row: {master_header_idx + 1})")
    st.write(f"📂 **Active OLT Sheet:** `{selected_olt_sheet}` (Header Row: {olt_header_idx + 1})")

    # Set cleaned column names onto dataframes for processing
    master_df.columns = master_df_cleaned.columns
    olt_df.columns = olt_df_cleaned.columns

    # Standardize primary linking frames
    master_df[master_plaid_col_clean] = master_df[master_plaid_col_clean].astype(str).str.strip()
    olt_df[olt_plaid_col_clean] = olt_df[olt_plaid_col_clean].astype(str).str.strip()

    # Isolate valid rows
    master_clean_df = master_df[master_df[master_plaid_col_clean].str.lower() != "nan"].copy()
    missing_mask = ~master_clean_df[master_plaid_col_clean].isin(olt_df[olt_plaid_col_clean])
    missing_records = master_clean_df[missing_mask].copy()

    st.subheader("❌ Unmapped Raw Master Entries")
    st.write(f"Total Missing Rows Isolated: **{len(missing_records)}**")
    st.dataframe(missing_records.head(50), use_container_width=True)

    # -----------------------------
    # 🔄 Auto-Intersection Mapping Sequence
    # -----------------------------
    st.subheader("🔄 Intersecting Column Matrix Map")
    
    # Initialize compilation dataframe matching the exact original structure layout of Nokia OLT
    append_df = pd.DataFrame(columns=orig_olt_cols)
    
    # Track which mappings were discovered automatically for logging purposes
    mapped_columns_log = []

    # Map indices back and forth automatically based on match logic
    for orig_olt_col in orig_olt_cols:
        clean_olt_name = clean_string_normalization(orig_olt_col)
        
        # Look for this exact normalized text header inside the Master Tracker dataframe schema
        matched_master_col = None
        for clean_m_col in master_df.columns:
            if clean_string_normalization(clean_m_col) == clean_olt_name and clean_olt_name != "":
                matched_master_col = clean_m_col
                break
        
        # If found, migrate all data records cleanly across tables
        if matched_master_col:
            append_df[orig_olt_col] = missing_records[matched_master_col].tolist()
            mapped_columns_log.append(f"🔗 Appending **'{orig_olt_col}'** from Master field *'{matched_master_col}'*")
        else:
            # Check for standard telecom schema aliases if exact name didn't match
            if "plaid" in clean_olt_name:
                append_df[orig_olt_col] = missing_records[master_plaid_col_clean].tolist()
                mapped_columns_log.append(f"🔗 Appending **'{orig_olt_col}'** from Master primary identifier")
            elif "build year" in clean_olt_name and "year" in master_df.columns:
                append_df[orig_olt_col] = missing_records["year"].tolist()
                mapped_columns_log.append(f"🔗 Appending **'{orig_olt_col}'** from Master field *'YEAR'*")
            elif "cards" in clean_olt_name and "expansion" not in clean_olt_name and "number of cards" in master_df.columns:
                append_df[orig_olt_col] = missing_records["number of cards"].tolist()
                mapped_columns_log.append(f"🔗 Appending **'{orig_olt_col}'** from Master field *'Number of Cards'*")
            elif "equipment type" in clean_olt_name and "expansion" not in clean_olt_name and "electronics equipment" in master_df.columns:
                append_df[orig_olt_col] = missing_records["electronics equipment"].tolist()
                mapped_columns_log.append(f"🔗 Appending **'{orig_olt_col}'** from Master field *'Electronics Equipment'*")
            elif "site status" in clean_olt_name and "status" in master_df.columns:
                append_df[orig_olt_col] = missing_records["status"].tolist()
                mapped_columns_log.append(f"🔗 Appending **'{orig_olt_col}'** from Master field *'Status'*")
            else:
                # No data link match found: populate blank spacer column values
                append_df[orig_olt_col] = [""] * len(missing_records)

    with st.expander("👀 View auto-detected column connection mapping"):
        for log in mapped_columns_log:
            st.markdown(log)

    st.subheader("📋 Output Blueprint (Ready to append into Nokia OLT)")
    st.dataframe(append_df.head(100), use_container_width=True)

    # -----------------------------
    # 💾 In-Memory Workbook Appending Engine
    # -----------------------------
    if len(append_df) > 0:
        if st.button("🚀 Merge and Append into OLT Spreadsheet"):
            try:
                # Read structural layout tracking and load directly with openpyxl 
                wb = openpyxl.load_workbook(io.BytesIO(olt_bytes))
                ws = wb[selected_olt_sheet]
                
                # Append rows safely right after the last occupied row index
                start_row = ws.max_row + 1
                
                for r_idx, row_data in enumerate(append_df.values, start=start_row):
                    for c_idx, value in enumerate(row_data, start=1):
                        # Avoid rendering raw 'nan' values as text in the spreadsheet
                        if pd.isna(value) or str(value).lower() == "nan":
                            ws.cell(row=r_idx, column=c_idx, value="")
                        else:
                            ws.cell(row=r_idx, column=c_idx, value=value)
                
                # Output memory byte stream buffer
                out_buffer = io.BytesIO()
                wb.save(out_buffer)
                out_buffer.seek(0)
                
                st.success(f"🎉 Successfully matched, formatted, and appended {len(append_df)} records directly to the rollout file!")
                
                st.download_button(
                    label="⬇️ Download Updated Nokia OLT Tracker File",
                    data=out_buffer.getvalue(),
                    file_name=f"Updated_{olt_file.name}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as err:
                st.error(f"Failed to append entries inside workbook structure: {err}")
    else:
        st.info("ℹ️ All tracking datasets are completely synchronized. No missing fields to append.")