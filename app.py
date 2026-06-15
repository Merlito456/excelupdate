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

    # Parse initial DataFrames
    master_df = master_xls.parse(selected_master_sheet, header=master_header_idx)
    olt_df = olt_xls.parse(selected_olt_sheet, header=olt_header_idx)

    # Cache original layout headers
    orig_master_cols = list(master_df.columns)
    orig_olt_cols = list(olt_df.columns)

    # Standard clean background frames
    master_df_cleaned = clean_columns(master_df.copy())
    olt_df_cleaned = clean_columns(olt_df.copy())

    # Smart fallback guesses
    clean_m_plaid = find_column(master_df_cleaned.columns, ["plaid"])
    clean_o_plaid = find_column(olt_df_cleaned.columns, ["plaid"])
    clean_m_year = find_column(master_df_cleaned.columns, ["year"])
    clean_m_type = find_column(master_df_cleaned.columns, ["scope", "type", "nature"])

    def get_index_fallback(clean_target, original_list, clean_list):
        if clean_target in clean_list:
            return clean_list.index(clean_target)
        return 0

    # 🎯 CONTROL PANEL: Manual Mapping Dropdowns Overrides
    st.sidebar.subheader("🎯 Direct Column Mapping Controls")
    
    m_plaid_idx = get_index_fallback(clean_m_plaid, orig_master_cols, list(master_df_cleaned.columns))
    chosen_master_plaid_raw = st.sidebar.selectbox("🔑 Master PLAID Column (Unique Key)", orig_master_cols, index=m_plaid_idx)
    
    o_plaid_idx = get_index_fallback(clean_o_plaid, orig_olt_cols, list(olt_df_cleaned.columns))
    chosen_olt_plaid_raw = st.sidebar.selectbox("🔑 OLT PLAID Column", orig_olt_cols, index=o_plaid_idx)
    
    m_year_idx = get_index_fallback(clean_m_year, orig_master_cols, list(master_df_cleaned.columns))
    if m_year_idx == 0 and len(orig_master_cols) >= 5: m_year_idx = 4 # Default to Column E if year isn't matched
    chosen_master_year_raw = st.sidebar.selectbox("📅 Master Year Source column (e.g. 2021)", orig_master_cols, index=m_year_idx)

    m_type_idx = get_index_fallback(clean_m_type, orig_master_cols, list(master_df_cleaned.columns))
    chosen_master_type_raw = st.sidebar.selectbox("🛠️ Master Project Type/Scope Source Column", orig_master_cols, index=m_type_idx)

    # Convert user selection down to background programmatic strings
    master_plaid_col_clean = list(master_df_cleaned.columns)[orig_master_cols.index(chosen_master_plaid_raw)]
    olt_plaid_col_clean = list(olt_df_cleaned.columns)[orig_olt_cols.index(chosen_olt_plaid_raw)]
    master_year_col_clean = list(master_df_cleaned.columns)[orig_master_cols.index(chosen_master_year_raw)]
    master_type_col_clean = list(master_df_cleaned.columns)[orig_master_cols.index(chosen_master_type_raw)]

    # -----------------------------
    # 📉 Missing Records Analysis
    # -----------------------------
    st.write(f"📂 **Active Master Sheet:** `{selected_master_sheet}` | **Active OLT Sheet:** `{selected_olt_sheet}`")

    master_df.columns = master_df_cleaned.columns
    olt_df.columns = olt_df_cleaned.columns

    master_df[master_plaid_col_clean] = master_df[master_plaid_col_clean].astype(str).str.strip()
    olt_df[olt_plaid_col_clean] = olt_df[olt_plaid_col_clean].astype(str).str.strip()

    master_clean_df = master_df[master_df[master_plaid_col_clean].str.lower() != "nan"].copy()
    missing_mask = ~master_clean_df[master_plaid_col_clean].isin(olt_df[olt_plaid_col_clean])
    missing_records = master_clean_df[missing_mask].copy()

    st.subheader("❌ Unmapped Raw Master Entries")
    st.write(f"Total Missing Rows Isolated: **{len(missing_records)}**")
    st.dataframe(missing_records.head(20), use_container_width=True)

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
        "cards": ["cards", "number of cards", "no of cards", "card count"],
        "equipment type": ["equipment type", "electronics equipment", "equipment model", "chassis type"],
        "status": ["status", "site status", "rollout status", "state"]
    }

    for orig_olt_col in orig_olt_cols:
        clean_olt_name = clean_string_normalization(orig_olt_col)
        matched_master_col = None

        # 🚨 OVERRIDE 1: Unique Keys
        if "plaid" in clean_olt_name:
            matched_master_col = master_plaid_col_clean
        
        # 🚨 OVERRIDE 2: Build Year Output (appends " build" text)
        elif "build year" in clean_olt_name:
            matched_master_col = master_year_col_clean
            raw_values = missing_records[matched_master_col].tolist()
            formatted_years = [f"{str(val).split('.')[0].strip()} build" if pd.notna(val) and str(val).strip() != "" and str(val).lower() != "nan" else "" for val in raw_values]
            append_df[orig_olt_col] = formatted_years
            mapped_columns_log.append(f"📅 **Manual Rule Override**: Nokia '{orig_olt_col}' ← Master '{matched_master_col}' + ' build'")
            continue

        # 🚨 OVERRIDE 3: Project Type Output (Pulls 100% raw unedited data)
        elif "project type" in clean_olt_name:
            matched_master_col = master_type_col_clean
            append_df[orig_olt_col] = missing_records[matched_master_col].tolist()
            mapped_columns_log.append(f"📋 **Direct Raw Copy**: Nokia '{orig_olt_col}' ← Master Selector Custom Source '{matched_master_col}'")
            continue

        # Fallback automated mapping checks
        if not matched_master_col:
            for clean_m_col in master_df.columns:
                if clean_string_normalization(clean_m_col) == clean_olt_name and clean_olt_name != "":
                    matched_master_col = clean_m_col
                    break
            
            if not matched_master_col and clean_olt_name != "":
                for base_key, variations in alias_map.items():
                    if any(v in clean_olt_name for v in variations):
                        for clean_m_col in master_df.columns:
                            clean_m_norm = clean_string_normalization(clean_m_col)
                            if any(v == clean_m_norm for v in variations):
                                matched_master_col = clean_m_col
                                break
                    if matched_master_col:
                        break

        # Map findings into structural columns
        if matched_master_col:
            append_df[orig_olt_col] = missing_records[matched_master_col].tolist()
            mapped_columns_log.append(f"🔗 Linked OLT **'{orig_olt_col}'** ← Master *'{matched_master_col}'*")
        else:
            append_df[orig_olt_col] = [""] * len(missing_records)

    with st.expander("👀 View automated column connection mapping mapping audit trail"):
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
                wb = openpyxl.load_workbook(io.BytesIO(olt_bytes))
                ws = wb[selected_olt_sheet]
                
                start_row = ws.max_row + 1
                for r_idx, row_data in enumerate(append_df.values, start=start_row):
                    for c_idx, value in enumerate(row_data, start=1):
                        if pd.isna(value) or str(value).lower() == "nan":
                            ws.cell(row=r_idx, column=c_idx, value="")
                        else:
                            ws.cell(row=r_idx, column=c_idx, value=value)
                
                out_buffer = io.BytesIO()
                wb.save(out_buffer)
                out_buffer.seek(0)
                
                st.success(f"🎉 Successfully mapped, formatted, and appended {len(append_df)} records directly into the Nokia OLT Tracker sheet!")
                st.download_button(
                    label="⬇️ Download Updated Nokia OLT Tracker File",
                    data=out_buffer.getvalue(),
                    file_name=f"Updated_{olt_file.name}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as err:
                st.error(f"Failed to append entries inside workbook structure: {err}")