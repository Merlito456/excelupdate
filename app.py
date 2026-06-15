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
    # Read bytes to enable multiple parsing operations safely
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

    # Parse initial structures
    master_df = master_xls.parse(selected_master_sheet, header=master_header_idx)
    olt_df = olt_xls.parse(selected_olt_sheet, header=olt_header_idx)

    orig_master_cols = list(master_df.columns)
    orig_olt_cols = list(olt_df.columns)

    master_df = clean_columns(master_df)
    olt_df = clean_columns(olt_df)

    # Column Mapping Selectors
    auto_m_plaid = find_column(master_df.columns, ["plaid"])
    auto_m_site = find_column(master_df.columns, ["site name"])
    auto_o_plaid = find_column(olt_df.columns, ["plaid"])

    def get_index_fallback(clean_target, original_list, clean_list):
        if clean_target in clean_list:
            return clean_list.index(clean_target)
        return 0

    st.sidebar.subheader("🎯 Target Column Mapping")
    
    m_plaid_idx = get_index_fallback(auto_m_plaid, orig_master_cols, list(master_df.columns))
    chosen_master_plaid_raw = st.sidebar.selectbox("Master PLAID Column", orig_master_cols, index=m_plaid_idx)
    
    m_site_idx = get_index_fallback(auto_m_site, orig_master_cols, list(master_df.columns))
    chosen_master_site_raw = st.sidebar.selectbox("Master Site Name Column", orig_master_cols, index=m_site_idx)

    o_plaid_idx = get_index_fallback(auto_o_plaid, orig_olt_cols, list(olt_df.columns))
    chosen_olt_plaid_raw = st.sidebar.selectbox("OLT PLAID Column", orig_olt_cols, index=o_plaid_idx)

    master_plaid = list(master_df.columns)[orig_master_cols.index(chosen_master_plaid_raw)]
    master_site = list(master_df.columns)[orig_master_cols.index(chosen_master_site_raw)]
    olt_plaid = list(olt_df.columns)[orig_olt_cols.index(chosen_olt_plaid_raw)]

    # -----------------------------
    # 📉 Core Extraction Logic
    # -----------------------------

    st.write(f"📂 **Active Master Sheet:** `{selected_master_sheet}` (Header Row: {master_header_idx + 1})")
    st.write(f"📂 **Active OLT Sheet:** `{selected_olt_sheet}` (Header Row: {olt_header_idx + 1})")

    master_df[master_plaid] = master_df[master_plaid].astype(str).str.strip()
    olt_df[olt_plaid] = olt_df[olt_plaid].astype(str).str.strip()

    master_clean_df = master_df[master_df[master_plaid].str.lower() != "nan"]
    missing_mask = ~master_clean_df[master_plaid].isin(olt_df[olt_plaid])
    missing_df = master_clean_df[missing_mask].copy()

    st.subheader("❌ Isolated Missing Entries")
    st.write(f"Total Missing Rows Found: **{len(missing_df)}**")
    st.dataframe(missing_df.head(100), use_container_width=True)

    # -----------------------------
    # 🔄 Build Appending Rows Dataframe
    # -----------------------------
    # This acts as a blueprint structured exactly like your targeted OLT tracker schema
    append_df = pd.DataFrame(columns=orig_olt_cols)
    
    # Extract data safely from Master to assign to the specific OLT layout
    # Use fallback mapping to match your structural targets
    m_site_col = chosen_master_site_raw
    m_plaid_col = chosen_master_plaid_raw

    # Prepare specific data lists
    site_names = missing_df[master_site].tolist() if master_site else [""] * len(missing_df)
    plaids = missing_df[master_plaid].tolist()
    
    regions = missing_df["Region"].tolist() if "Region" in missing_df.columns else [""] * len(missing_df)
    
    if "Build Year" in missing_df.columns:
        years = missing_df["Build Year"].tolist()
    elif "YEAR" in missing_df.columns:
        years = missing_df["YEAR"].tolist()
    else:
        years = [""] * len(missing_df)

    cards = missing_df["Number of Cards"].tolist() if "Number of Cards" in missing_df.columns else [""] * len(missing_df)
    eq_type = missing_df["Electronics Equipment"].tolist() if "Electronics Equipment" in missing_df.columns else [""] * len(missing_df)
    status = missing_df["Status"].tolist() if "Status" in missing_df.columns else [""] * len(missing_df)

    # Insert lists dynamically back to OLT column spaces using exact name strings from raw headers
    for col in append_df.columns:
        clean_name = clean_string_normalization(col)
        if "site name" in clean_name:
            append_df[col] = site_names
        elif "plaid" in clean_name:
            append_df[col] = plaids
        elif "region" in clean_name:
            append_df[col] = regions
        elif "build year" in clean_name:
            append_df[col] = years
        elif "cards" in clean_name and "expansion" not in clean_name:
            append_df[col] = cards
        elif "equipment type" in clean_name and "expansion" not in clean_name:
            append_df[col] = eq_type
        elif "status" in clean_name:
            append_df[col] = status
        else:
            append_df[col] = "" # Blank spacer values for remaining rollout metrics

    st.subheader("📋 Rows to be Appended (Nokia Layout Format)")
    st.dataframe(append_df.head(100), use_container_width=True)

    # -----------------------------
    # 💾 In-Memory Appending Work Engine 
    # -----------------------------
    if len(append_df) > 0:
        if st.button("🚀 Merge and Append into OLT Spreadsheet"):
            try:
                # Load existing workbook with openpyxl to maintain historical styling blocks
                wb = openpyxl.load_workbook(io.BytesIO(olt_bytes))
                ws = wb[selected_olt_sheet]
                
                # Determine append entry start bounds using the 1-based index setup
                start_row = ws.max_row + 1
                
                # Write data records sequentially row by row
                for r_idx, row_data in enumerate(append_df.values, start=start_row):
                    for c_idx, value in enumerate(row_data, start=1):
                        ws.cell(row=r_idx, column=c_idx, value=value)
                
                # Output memory byte stream buffer
                out_buffer = io.BytesIO()
                wb.save(out_buffer)
                out_buffer.seek(0)
                
                st.success(f"🎉 Successfully merged {len(append_df)} rows directly to the base of tracking dataset!")
                
                st.download_button(
                    label="⬇️ Download Updated Nokia OLT Tracker File",
                    data=out_buffer.getvalue(),
                    file_name=f"Updated_{olt_file.name}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as err:
                st.error(f"Failed to append entries inside workbook structure: {err}")
    else:
        st.info("ℹ️ All target systems match perfectly. No unique entries found to update.")