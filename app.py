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

    # Noise filtration for target ghost columns
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
        "cards": ["cards", "number of cards", "no of cards", "card count"]
    }

    for c_idx, orig_olt_col in enumerate(orig_olt_cols):
        clean_olt_name = clean_string_normalization(orig_olt_col)
        matched_master_col = None

        if clean_olt_name == "" or "solution track" in clean_olt_name or clean_olt_name == "track":
            continue

        # 🚨 OVERRIDE 1: Nokia Column B (Index 1) ← Master Column E (Index 4) [Build Year]
        if c_idx == 1: 
            if len(orig_master_cols) >= 5:
                matched_master_col = master_df.columns[4]
                raw_values = missing_records[matched_master_col].tolist()
                formatted_years = [f"{str(val).split('.')[0].strip()} build" if pd.notna(val) and str(val).strip() != "" and str(val).lower() != "nan" else "" for val in raw_values]
                append_df[orig_olt_col] = formatted_years
                mapped_columns_log.append(f"📅 **Position Linked**: Nokia Column B ('{orig_olt_col}') ← Master Column E ('{matched_master_col}') + ' build'")
                continue

        # 🚨 OVERRIDE 2: Nokia Project Type (Column F / Index 5) ← Master Column M (Index 12) [OLT Scope]
        if "project type" in clean_olt_name or c_idx == 5:
            if len(orig_master_cols) >= 13:
                matched_master_col = master_df.columns[12]
                append_df[orig_olt_col] = missing_records[matched_master_col].tolist()
                mapped_columns_log.append(f"📋 **Position Linked**: Nokia Column F ('{orig_olt_col}') ← Master Column M ('{matched_master_col}') [OLT Scope]")
                continue

        # 🚨 OVERRIDE 3: Nokia Equipment Type (Column N / Index 13) ← Master Column N (Index 13) [Electronics Equipment]
        if "equipment type" in clean_olt_name or c_idx == 13:
            if len(orig_master_cols) >= 14:
                matched_master_col = master_df.columns[13]
                append_df[orig_olt_col] = missing_records[matched_master_col].tolist()
                mapped_columns_log.append(f"⚙️ **Position Linked**: Nokia Column N ('{orig_olt_col}') ← Master Column N ('{matched_master_col}') [Electronics Equipment]")
                continue

        # 🚨 OVERRIDE 4: Nokia Site Status (Column V / Index 21) ← Master Column L (Index 11) [Scope Status]
        if "site status" in clean_olt_name or c_idx == 21:
            if len(orig_master_cols) >= 12:
                matched_master_col = master_df.columns[11]
                append_df[orig_olt_col] = missing_records[matched_master_col].tolist()
                mapped_columns_log.append(f"📡 **Position Linked**: Nokia Column V ('{orig_olt_col}') ← Master Column L ('{matched_master_col}') [Scope status]")
                continue

        # 🚨 OVERRIDE 5: Nokia Site Survey Actual Date (Column AO / Index 40) ← Master Column X (Index 23) [Survey Date]
        if "site survey actual date" in clean_olt_name or c_idx == 40:
            if len(orig_master_cols) >= 24:
                matched_master_col = master_df.columns[23]
                raw_dates = missing_records[matched_master_col].tolist()
                cleaned_dates = [str(d_val).split(" ")[0] if pd.notna(d_val) and str(d_val).lower() != "nan" else "" for d_val in raw_dates]
                append_df[orig_olt_col] = cleaned_dates
                mapped_columns_log.append(f"📆 **Position Linked**: Nokia Column AO ('{orig_olt_col}') ← Master Column X ('{matched_master_col}') [Survey Date]")
                continue

        # 🚨 OVERRIDE 6: Nokia Installation done Actual Date (Column BD / Index 55) ← Master Column AA (Index 26) [Installed date]
        if "installation done actual date" in clean_olt_name or c_idx == 55:
            if len(orig_master_cols) >= 27:
                matched_master_col = master_df.columns[26]
                raw_dates = missing_records[matched_master_col].tolist()
                cleaned_dates = [str(d_val).split(" ")[0] if pd.notna(d_val) and str(d_val).lower() != "nan" else "" for d_val in raw_dates]
                append_df[orig_olt_col] = cleaned_dates
                mapped_columns_log.append(f"🏗️ **Position Linked**: Nokia Column BD ('{orig_olt_col}') ← Master Column AA ('{matched_master_col}') [Installed date]")
                continue

        # 🚨 OVERRIDE 7: Nokia Powertapping done Actual Date (Column BK / Index 62) ← Master Column AC (Index 28) [Powertapped date]
        if "powertapping done actual date" in clean_olt_name or c_idx == 62:
            if len(orig_master_cols) >= 29:
                matched_master_col = master_df.columns[28]
                raw_dates = missing_records[matched_master_col].tolist()
                cleaned_dates = [str(d_val).split(" ")[0] if pd.notna(d_val) and str(d_val).lower() != "nan" else "" for d_val in raw_dates]
                append_df[orig_olt_col] = cleaned_dates
                mapped_columns_log.append(f"⚡ **Position Linked**: Nokia Column BK ('{orig_olt_col}') ← Master Column AC ('{matched_master_col}') [Powertapped date]")
                continue

        # 🚨 OVERRIDE 8: Nokia Integration done Actual Date (Column BS / Index 70) ← Master Column AD (Index 29) [Integrated date]
        if "integration done actual date" in clean_olt_name or c_idx == 70:
            if len(orig_master_cols) >= 30:
                matched_master_col = master_df.columns[29]
                raw_dates = missing_records[matched_master_col].tolist()
                cleaned_dates = [str(d_val).split(" ")[0] if pd.notna(d_val) and str(d_val).lower() != "nan" else "" for d_val in raw_dates]
                append_df[orig_olt_col] = cleaned_dates
                mapped_columns_log.append(f"🌐 **Position Linked**: Nokia Column BS ('{orig_olt_col}') ← Master Column AD ('{matched_master_col}') [Integrated date]")
                continue

        # 🚨 OVERRIDE 9: Nokia PAT done Actual Date (Column CB / Index 79) ← Master Column AF (Index 31) [Pat'ed]
        if "pat done actual date" in clean_olt_name or c_idx == 79:
            if len(orig_master_cols) >= 32:
                matched_master_col = master_df.columns[31]
                raw_dates = missing_records[matched_master_col].tolist()
                cleaned_dates = [str(d_val).split(" ")[0] if pd.notna(d_val) and str(d_val).lower() != "nan" else "" for d_val in raw_dates]
                append_df[orig_olt_col] = cleaned_dates
                mapped_columns_log.append(f"📋 **Position Linked**: Nokia Column CB ('{orig_olt_col}') ← Master Column AF ('{matched_master_col}') [Pat'ed]")
                continue

        # 🚨 OVERRIDE 10: Nokia PAC Approval done Actual Date (Column CM / Index 90) ← Master Column AI (Index 34) [PAC'ed]
        if "pac approval done actual date" in clean_olt_name or c_idx == 90:
            if len(orig_master_cols) >= 35:
                matched_master_col = master_df.columns[34]
                raw_dates = missing_records[matched_master_col].tolist()
                cleaned_dates = [str(d_val).split(" ")[0] if pd.notna(d_val) and str(d_val).lower() != "nan" else "" for d_val in raw_dates]
                append_df[orig_olt_col] = cleaned_dates
                mapped_columns_log.append(f"📜 **Position Linked**: Nokia Column CM ('{orig_olt_col}') ← Master Column AI ('{matched_master_col}') [PAC'ed]")
                continue

        # 🚨 OVERRIDE 11: Nokia FAC Approval done Actual Date (Column DC / Index 106) ← Master Column AJ (Index 35) [FAC'ed]
        if "fac approval done actual date" in clean_olt_name or c_idx == 106:
            if len(orig_master_cols) >= 36:
                matched_master_col = master_df.columns[35] # Column AJ is index 35
                raw_dates = missing_records[matched_master_col].tolist()
                cleaned_dates = [str(d_val).split(" ")[0] if pd.notna(d_val) and str(d_val).lower() != "nan" else "" for d_val in raw_dates]
                append_df[orig_olt_col] = cleaned_dates
                mapped_columns_log.append(f"🏆 **Position Linked (Date Sanitized)**: Nokia Column DC ('{orig_olt_col}') ← Master Column AJ ('{matched_master_col}') [FAC'ed]")
                continue

        # Force key tracking structural link
        if "plaid" in clean_olt_name:
            matched_master_col = master_plaid_col_clean

        # Fallback loop name analyzer
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
        else:
            append_df[orig_olt_col] = [""] * len(missing_records)

    # Exclude formatting column anomalies from data grid preview frames
    append_df = append_df.loc[:, ~append_df.columns.astype(str).str.contains('track', case=False)]

    with st.expander("👀 View automated column connection mapping mapping audit trail"):
        for log in mapped_columns_log:
            st.markdown(log)

    st.subheader("📋 Output Blueprint (Ready to append into Nokia OLT)")
    st.dataframe(append_df.head(100), use_container_width=True)

    # -----------------------------
    # 💾 FIXED: In-Memory Workbook Appending Engine
    # -----------------------------
    if len(append_df) > 0:
        if st.button("🚀 Merge and Append into OLT Spreadsheet"):
            try:
                # 1. Parse existing data directly via pandas (bypassing openpyxl structure errors)
                base_olt_df = olt_xls.parse(selected_olt_sheet, header=olt_header_idx)
                
                # 2. Append new data to the existing dataframe
                final_combined_df = pd.concat([base_olt_df, append_df], ignore_index=True)
                
                # 3. Create output stream using ExcelWriter
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