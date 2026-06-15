import streamlit as st
import pandas as pd
import string
import io

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
        
        # Deduplication tracking sequence
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
    # 🛠️ Sidebar Configuration & Controls (Override Panels)
    # -----------------------------
    st.sidebar.header("🛠️ Configuration Controls")
    
    # Sheet Selection
    auto_master_sheet = detect_sheet(master_xls, ["master", "luzon", "file"])
    auto_olt_sheet = detect_sheet(olt_xls, ["rollout", "nokia", "olt", "summary", "inventory"])
    
    selected_master_sheet = st.sidebar.selectbox("Master Sheet Name", master_xls.sheet_names, index=master_xls.sheet_names.index(auto_master_sheet))
    selected_olt_sheet = st.sidebar.selectbox("OLT Sheet Name", olt_xls.sheet_names, index=olt_xls.sheet_names.index(auto_olt_sheet))

    # Header Row Index Selection
    auto_master_idx = find_dynamic_header_row(master_xls, selected_master_sheet)
    auto_olt_idx = find_dynamic_header_row(olt_xls, selected_olt_sheet)
    
    master_header_idx = st.sidebar.number_input("Master Header Row Index (1-based)", min_value=1, value=auto_master_idx + 1) - 1
    olt_header_idx = st.sidebar.number_input("OLT Header Row Index (1-based)", min_value=1, value=auto_olt_idx + 1) - 1

    # -----------------------------
    # 🏃 Execution Engine
    # -----------------------------
    
    # Parse DataFrames using targeted structural offsets
    master_df = master_xls.parse(selected_master_sheet, header=master_header_idx)
    olt_df = olt_xls.parse(selected_olt_sheet, header=olt_header_idx)

    # Backup original columns safely for select-boxes
    orig_master_cols = list(master_df.columns)
    orig_olt_cols = list(olt_df.columns)

    # Standardize schemas and safely remove duplicates
    master_df = clean_columns(master_df)
    olt_df = clean_columns(olt_df)

    # Attempt Auto Mapping
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

    # Re-map clean pointers based on dropdown parameters
    master_plaid = list(master_df.columns)[orig_master_cols.index(chosen_master_plaid_raw)]
    master_site = list(master_df.columns)[orig_master_cols.index(chosen_master_site_raw)]
    olt_plaid = list(olt_df.columns)[orig_olt_cols.index(chosen_olt_plaid_raw)]

    # -----------------------------
    # 📉 Core Business Automation Flow
    # -----------------------------

    st.write(f"📂 **Active Master Sheet:** `{selected_master_sheet}` (Header Row: {master_header_idx + 1})")
    st.write(f"📂 **Active OLT Sheet:** `{selected_olt_sheet}` (Header Row: {olt_header_idx + 1})")

    # Vectorized String Normalization
    master_df[master_plaid] = master_df[master_plaid].astype(str).str.strip()
    olt_df[olt_plaid] = olt_df[olt_plaid].astype(str).str.strip()

    # Discrepancy Tracking (Exclude missing / nan entries)
    master_clean_df = master_df[master_df[master_plaid].str.lower() != "nan"]
    missing_mask = ~master_clean_df[master_plaid].isin(olt_df[olt_plaid])
    missing_df = master_clean_df[missing_mask].copy()

    st.subheader("❌ Missing Entries (Data → Rollout)")
    st.write(f"Total Missing Rows Isolated: **{len(missing_df)}**")
    st.dataframe(missing_df.head(100), use_container_width=True)

    # Generation & Mapping Matrix
    st.subheader("🔄 Generated Data Structure Mapping")
    mapped = pd.DataFrame(index=master_clean_df.index)

    mapped["Site Name"] = master_clean_df[master_site] if master_site else ""
    mapped["PLAID"] = master_clean_df[master_plaid]
    mapped["Region"] = master_clean_df["Region"] if "Region" in master_clean_df.columns else ""
    
    if "Build Year" in master_clean_df.columns:
        mapped["Build Year"] = master_clean_df["Build Year"]
    elif "YEAR" in master_clean_df.columns:
        mapped["Build Year"] = master_clean_df["YEAR"]
    else:
        mapped["Build Year"] = ""

    mapped["No. of Cards"] = master_clean_df["Number of Cards"] if "Number of Cards" in master_clean_df.columns else ""
    mapped["Equipment Type"] = master_clean_df["Electronics Equipment"] if "Electronics Equipment" in master_clean_df.columns else ""
    mapped["Site Status"] = master_clean_df["Status"] if "Status" in master_clean_df.columns else ""

    new_rows = mapped[missing_mask]
    st.dataframe(new_rows.head(100), use_container_width=True)

    # In-Memory Workbook Packaging & Serialization
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        missing_df.to_excel(writer, sheet_name="Missing_Records", index=False)
        new_rows.to_excel(writer, sheet_name="Formatted_Upload_Rows", index=False)
    
    st.markdown("---")
    st.download_button(
        label="⬇️ Download Discrepancy Reports (.xlsx)",
        data=buffer.getvalue(),
        file_name="OLT_Missing_Entries.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )