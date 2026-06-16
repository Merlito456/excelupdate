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
    auto_olt_sheet = detect_sheet(xls=olt_xls, keywords=["rollout", "nokia", "olt", "summary", "inventory"])
    
    selected_master_sheet = st.sidebar.selectbox("Master Sheet Name", master_xls.sheet_names, index=master_xls.sheet_names.index(auto_master_sheet))
    selected_olt_sheet = st.sidebar.selectbox("OLT Sheet Name", olt_xls.sheet_names, index=olt_xls.sheet_names.index(auto_olt_sheet))

    auto_master_idx = find_dynamic_header_row(master_xls, selected_master_sheet)
    auto_olt_idx = find_dynamic_header_row(olt_xls, selected_olt_sheet)
    
    master_header_idx = st.sidebar.number_input("Master Header Row Index (1-based)", min_value=1, value=auto_master_idx + 1) - 1
    olt_header_idx = st.sidebar.number_input("OLT Header Row Index (1-based)", min_value=1, value=auto_olt_idx + 1) - 1

    # Parse dataframes
    master_df = master_xls.parse(selected_master_sheet, header=master_header_idx)
    olt_df = olt_xls.parse(selected_olt_sheet, header=olt_header_idx)

    # Noise filtration
    olt_df = olt_df.loc[:, ~olt_df.columns.astype(str).str.startswith('Unnamed:')]
    olt_df = olt_df.loc[:, olt_df.columns.notna() & (olt_df.columns != "")]
    
    # Save original headers
    orig_master_cols = list(master_df.columns)
    orig_olt_cols = list(olt_df.columns)

    # Standard clean background frames
    master_df_cleaned = clean_columns(master_df.copy())
    olt_df_cleaned = clean_columns(olt_df.copy())

    clean_m_plaid = find_column(master_df_cleaned.columns, ["plaid"])
    clean_o_plaid = find_column(olt_df_cleaned.columns, ["plaid"])

    def get_index_fallback(clean_target, original_list, clean_list):
        if clean_target in clean_list:
            return clean_list.index(clean_target)
        return 0

    st.sidebar.subheader("🎯 Primary Identifier Validation")
    m_plaid_idx = get_index_fallback(clean_m_plaid, orig_master_cols, list(master_df_cleaned.columns))
    chosen_master_plaid_raw = st.sidebar.selectbox("🔑 Master PLAID Column (Unique Key)", orig_master_cols, index=m_plaid_idx)
    
    o_plaid_idx = get_index_fallback(clean_o_plaid, orig_olt_cols, list(olt_df_cleaned.columns))
    chosen_olt_plaid_raw = st.sidebar.selectbox("🔑 OLT PLAID Column", orig_olt_cols, index=o_plaid_idx)

    master_plaid_col_clean = list(master_df_cleaned.columns)[orig_master_cols.index(chosen_master_plaid_raw)]
    olt_plaid_col_clean = list(olt_df_cleaned.columns)[orig_olt_cols.index(chosen_olt_plaid_raw)]

    # -----------------------------
    # 📉 Data Integrity Analysis
    # -----------------------------
    st.write(f"📂 **Active Master Sheet:** `{selected_master_sheet}` | **Active OLT Sheet:** `{selected_olt_sheet}`")

    master_df.columns = master_df_cleaned.columns
    olt_df.columns = olt_df_cleaned.columns

    master_df[master_plaid_col_clean] = master_df[master_plaid_col_clean].astype(str).str.strip()
    olt_df[olt_plaid_col_clean] = olt_df[olt_plaid_col_clean].astype(str).str.strip()

    # Identify duplicates in Master
    master_df['is_duplicate'] = master_df.duplicated(subset=[master_plaid_col_clean], keep=False)
    
    # Identify missing (not in OLT)
    missing_mask = ~master_df[master_plaid_col_clean].isin(olt_df[olt_plaid_col_clean])
    missing_records = master_df[missing_mask].copy()

    st.subheader("📋 Master Data View (Highlighting Duplicates)")
    
    def highlight_duplicates(row):
        return ['background-color: #ffcccb' if row['is_duplicate'] else '' for _ in row]

    st.dataframe(
        master_df.style.apply(highlight_duplicates, axis=1),
        use_container_width=True
    )

    st.write(f"Total Unique Missing Rows to be Added: **{len(missing_records)}**")

    # -----------------------------
    # 🔄 High-Precision Matrix Map Engine
    # -----------------------------
    st.subheader("🔄 Intersecting Column Matrix Map")
    
    append_df = pd.DataFrame(columns=orig_olt_cols)
    mapped_columns_log = []

    alias_map = {
        "project tagging": ["project tagging", "project or program", "project", "program and project tagging", "program project"],
        "site name": ["site name", "sitename", "station name", "site description"],
        "clustering": ["clustering", "territory", "area", "cluster"],
        "province": ["province", "territory", "area", "region"],
        "cards": ["cards", "number of cards", "no of cards", "card count"]
    }

    for c_idx, orig_olt_col in enumerate(orig_olt_cols):
        clean_olt_name = clean_string_normalization(orig_olt_col)
        matched_master_col = None

        if clean_olt_name == "" or "solution track" in clean_olt_name or clean_olt_name == "track":
            continue

        if "clustering" in clean_olt_name or c_idx == 3:
            append_df[orig_olt_col] = [""] * len(missing_records)
            mapped_columns_log.append(f"🌍 **Position Linked**: Nokia Column D ('{orig_olt_col}') ← [INTENTIONALLY BLANK]")
            continue

        # [Mappings logic omitted for brevity, same as your original code]
        # ... (Include all your Override 1-11 logic here) ...

        if matched_master_col:
            append_df[orig_olt_col] = missing_records[matched_master_col].tolist()
        else:
            append_df[orig_olt_col] = [""] * len(missing_records)

    append_df = append_df.loc[:, ~append_df.columns.astype(str).str.contains('track', case=False)]

    with st.expander("👀 View automated column connection mapping audit trail"):
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
                base_olt_df = olt_xls.parse(selected_olt_sheet, header=olt_header_idx)
                final_combined_df = pd.concat([base_olt_df, append_df], ignore_index=True)
                
                out_buffer = io.BytesIO()
                with pd.ExcelWriter(out_buffer, engine='openpyxl') as writer:
                    final_combined_df.to_excel(writer, sheet_name=selected_olt_sheet, index=False)
                
                out_buffer.seek(0)
                st.success(f"🎉 Successfully merged and processed {len(append_df)} records!")
                st.download_button(
                    label="⬇️ Download Updated Nokia OLT Tracker File",
                    data=out_buffer.getvalue(),
                    file_name=f"Updated_{olt_file.name}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as err:
                st.error(f"Failed to process file: {err}")