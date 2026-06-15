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

    # Map target strings to their clean representations 
    m_normalized_cols = list(master_df.columns)

    for c_idx, orig_olt_col in enumerate(orig_olt_cols):
        clean_olt_name = clean_string_normalization(orig_olt_col)
        matched_master_col = None

        if clean_olt_name == "" or "solution track" in clean_olt_name or clean_olt_name == "track":
            continue

        # 🚨 METADATA OVERRIDE 1: Project Tagging ← PROJECT or PROGRAM
        if clean_olt_name == "project tagging":
            if "project or program" in m_normalized_cols:
                matched_master_col = "project or program"
                append_df[orig_olt_col] = missing_records[matched_master_col].tolist()
                mapped_columns_log.append(f"📋 **Schema Linked**: OLT ('{orig_olt_col}') ← Master Tracker ('{matched_master_col}')")
                continue

        # 🚨 METADATA OVERRIDE 2: Build Year ← YEAR
        if clean_olt_name == "build year":
            if "year" in m_normalized_cols:
                matched_master_col = "year"
                raw_values = missing_records[matched_master_col].tolist()
                formatted_years = [f"{str(val).split('.')[0].strip()} build" if pd.notna(val) and str(val).strip() != "" and str(val).lower() != "nan" else "" for val in raw_values]
                append_df[orig_olt_col] = formatted_years
                mapped_columns_log.append(f"📅 **Schema Linked**: OLT ('{orig_olt_col}') ← Master Tracker ('{matched_master_col}') + ' build'")
                continue

        # 🚨 METADATA OVERRIDE 3: Region ← Region
        if clean_olt_name == "region":
            if "region" in m_normalized_cols:
                matched_master_col = "region"
                append_df[orig_olt_col] = missing_records[matched_master_col].tolist()
                mapped_columns_log.append(f"📍 **Schema Linked**: OLT ('{orig_olt_col}') ← Master Tracker ('{matched_master_col}')")
                continue

        # 🚨 METADATA OVERRIDE 4: Clustering ← TERRITORY
        if clean_olt_name == "clustering":
            if "territory" in m_normalized_cols:
                matched_master_col = "territory"
                append_df[orig_olt_col] = missing_records[matched_master_col].tolist()
                mapped_columns_log.append(f"🌍 **Schema Linked**: OLT ('{orig_olt_col}') ← Master Tracker ('{matched_master_col}')")
                continue

        # 🚨 METADATA OVERRIDE 5: Project Type ← OLT Scope
        if clean_olt_name == "project type":
            if "olt scope" in m_normalized_cols:
                matched_master_col = "olt scope"
                append_df[orig_olt_col] = missing_records[matched_master_col].tolist()
                mapped_columns_log.append(f"📋 **Schema Linked**: OLT ('{orig_olt_col}') ← Master Tracker ('{matched_master_col}')")
                continue

        # 🚨 METADATA OVERRIDE 6: Site Name ← Site Name
        if clean_olt_name == "site name":
            if "site name" in m_normalized_cols:
                matched_master_col = "site name"
                append_df[orig_olt_col] = missing_records[matched_master_col].tolist()
                mapped_columns_log.append(f"📡 **Schema Linked**: OLT ('{orig_olt_col}') ← Master Tracker ('{matched_master_col}')")
                continue

        # 🚨 METADATA OVERRIDE 7: Equipment Type ← Electronics Equipment
        if clean_olt_name == "equipment type":
            if "electronics equipment" in m_normalized_cols:
                matched_master_col = "electronics equipment"
                append_df[orig_olt_col] = missing_records[matched_master_col].tolist()
                mapped_columns_log.append(f"⚙️ **Schema Linked**: OLT ('{orig_olt_col}') ← Master Tracker ('{matched_master_col}')")
                continue

        # 🚨 METADATA OVERRIDE 8: No. of Cards / Number of Cards
        if clean_olt_name == "no of cards":
            if "number of cards" in m_normalized_cols:
                matched_master_col = "number of cards"
                append_df[orig_olt_col] = missing_records[matched_master_col].tolist()
                mapped_columns_log.append(f"💳 **Schema Linked**: OLT ('{orig_olt_col}') ← Master Tracker ('{matched_master_col}')")
                continue

        # 🚨 METADATA OVERRIDE 9: Site Status ← Scope Status
        if clean_olt_name == "site status":
            if "scope status" in m_normalized_cols:
                matched_master_col = "scope status"
                append_df[orig_olt_col] = missing_records[matched_master_col].tolist()
                mapped_columns_log.append(f"📊 **Schema Linked**: OLT ('{orig_olt_col}') ← Master Tracker ('{matched_master_col}')")
                continue

        # 🚨 MILESTONE DATE OVERRIDES WITH AUTOMATIC TIMESTAMP STRIPPING
        if clean_olt_name == "site survey actual date":
            if "survey date" in m_normalized_cols:
                matched_master_col = "survey date"
                raw_dates = missing_records[matched_master_col].tolist()
                cleaned_dates = [str(d_val).split(" ")[0] if pd.notna(d_val) and str(d_val).lower() != "nan" else "" for d_val in raw_dates]
                append_df[orig_olt_col] = cleaned_dates
                mapped_columns_log.append(f"📆 **Milestone Linked**: OLT ('{orig_olt_col}') ← Master Tracker ('{matched_master_col}')")
                continue

        if clean_olt_name == "tssr approval actual date" or clean_olt_name == "tssr submission actual date":
            if "tssr approved date" in m_normalized_cols:
                matched_master_col = "tssr approved date"
                raw_dates = missing_records[matched_master_col].tolist()
                cleaned_dates = [str(d_val).split(" ")[0] if pd.notna(d_val) and str(d_val).lower() != "nan" else "" for d_val in raw_dates]
                append_df[orig_olt_col] = cleaned_dates
                mapped_columns_log.append(f"📝 **Milestone Linked**: OLT ('{orig_olt_col}') ← Master Tracker ('{matched_master_col}')")
                continue

        if clean_olt_name == "installation done actual date":
            if "installed date" in m_normalized_cols:
                matched_master_col = "installed date"
                raw_dates = missing_records[matched_master_col].tolist()
                cleaned_dates = [str(d_val).split(" ")[0] if pd.notna(d_val) and str(d_val).lower() != "nan" else "" for d_val in raw_dates]
                append_df[orig_olt_col] = cleaned_dates
                mapped_columns_log.append(f"🏗️ **Milestone Linked**: OLT ('{orig_olt_col}') ← Master Tracker ('{matched_master_col}')")
                continue

        if clean_olt_name == "powertapping done actual date":
            if "powertapped date" in m_normalized_cols:
                matched_master_col = "powertapped date"
                raw_dates = missing_records[matched_master_col].tolist()
                cleaned_dates = [str(d_val).split(" ")[0] if pd.notna(d_val) and str(d_val).lower() != "nan" else "" for d_val in raw_dates]
                append_df[orig_olt_col] = cleaned_dates
                mapped_columns_log.append(f"⚡ **Milestone Linked**: OLT ('{orig_olt_col}') ← Master Tracker ('{matched_master_col}')")
                continue

        if clean_olt_name == "integration done actual date":
            if "integrated date" in m_normalized_cols:
                matched_master_col = "integrated date"
                raw_dates = missing_records[matched_master_col].tolist()
                cleaned_dates = [str(d_val).split(" ")[0] if pd.notna(d_val) and str(d_val).lower() != "nan" else "" for d_val in raw_dates]
                append_df[orig_olt_col] = cleaned_dates
                mapped_columns_log.append(f"🌐 **Milestone Linked**: OLT ('{orig_olt_col}') ← Master Tracker ('{matched_master_col}')")
                continue

        if clean_olt_name == "pat done actual date":
            if "pated" in m_normalized_cols:
                matched_master_col = "pated"
                raw_dates = missing_records[matched_master_col].tolist()
                cleaned_dates = [str(d_val).split(" ")[0] if pd.notna(d_val) and str(d_val).lower() != "nan" else "" for d_val in raw_dates]
                append_df[orig_olt_col] = cleaned_dates
                mapped_columns_log.append(f"📋 **Milestone Linked**: OLT ('{orig_olt_col}') ← Master Tracker ('{matched_master_col}')")
                continue

        if clean_olt_name == "pac approval actual date" or clean_olt_name == "pac submission actual date":
            if "paced" in m_normalized_cols:
                matched_master_col = "paced"
                raw_dates = missing_records[matched_master_col].tolist()
                cleaned_dates = [str(d_val).split(" ")[0] if pd.notna(d_val) and str(d_val).lower() != "nan" else "" for d_val in raw_dates]
                append_df[orig_olt_col] = cleaned_dates
                mapped_columns_log.append(f"📜 **Milestone Linked**: OLT ('{orig_olt_col}') ← Master Tracker ('{matched_master_col}')")
                continue

        if clean_olt_name == "fac approval actual date" or clean_olt_name == "fac submission actual date":
            if "faced" in m_normalized_cols:
                matched_master_col = "faced"
                raw_dates = missing_records[matched_master_col].tolist()
                cleaned_dates = [str(d_val).split(" ")[0] if pd.notna(d_val) and str(d_val).lower() != "nan" else "" for d_val in raw_dates]
                append_df[orig_olt_col] = cleaned_dates
                mapped_columns_log.append(f"🏆 **Milestone Linked**: OLT ('{orig_olt_col}') ← Master Tracker ('{matched_master_col}')")
                continue

        # Force key tracking structural link
        if "plaid" in clean_olt_name:
            matched_master_col = master_plaid_col_clean

        # Fallback automated logic for anything else
        if not matched_master_col:
            for clean_m_col in master_df.columns:
                if clean_string_normalization(clean_m_col) == clean_olt_name and clean_olt_name != "":
                    matched_master_col = clean_m_col
                    break

        # Record into structural blueprint
        if matched_master_col:
            append_df[orig_olt_col] = missing_records[matched_master_col].tolist()
            if f"'{orig_olt_col}'" not in str(mapped_columns_log):
                mapped_columns_log.append(f"🔗 Linked OLT **'{orig_olt_col}'** ← Master *'{matched_master_col}'*")
        else:
            append_df[orig_olt_col] = [""] * len(missing_records)

    # Exclude layout noise
    append_df = append_df.loc[:, ~append_df.columns.astype(str).str.contains('track', case=False)]

    with st.expander("👀 View automated column connection mapping mapping audit trail"):
        for log in mapped_columns_log:
            st.markdown(log)

    st.subheader("📋 Output Blueprint (Ready to append into Nokia OLT)")
    st.dataframe(append_df.head(100), use_container_width=True)

    # -----------------------------
    # 💾 Safe Appending Engine Pipeline
    # -----------------------------
    if len(append_df) > 0:
        if st.button("🚀 Merge and Append into OLT Spreadsheet"):
            append_success = False
            out_buffer = io.BytesIO()
            
            # --- METHOD 1: Clean openpyxl with node-tree error bypass flags ---
            try:
                wb = openpyxl.load_workbook(io.BytesIO(olt_bytes), data_only=False, keep_links=False)
                ws = wb[selected_olt_sheet]
                
                start_row = ws.max_row + 1
                for r_idx, row_data in enumerate(append_df.values, start=start_row):
                    for c_idx, value in enumerate(row_data, start=1):
                        if pd.isna(value) or str(value).lower() == "nan":
                            ws.cell(row=r_idx, column=c_idx, value="")
                        else:
                            ws.cell(row=r_idx, column=c_idx, value=value)
                
                wb.save(out_buffer)
                out_buffer.seek(0)
                append_success = True
                st.success("🎉 Successfully processed using Primary OpenPyXL Engine layout preservation rules!")
            except Exception as e:
                st.warning(f"Primary workbook structure protected or containing nested groups. Engaging safe fallback writer... (Error: {e})")
            
            # --- METHOD 2: Robust engine fallback via Pandas context writer if openpyxl core fails ---
            if not append_success:
                try:
                    out_buffer = io.BytesIO()
                    # Reload the pure data array to bypass broken layout tree definitions completely
                    base_olt_df = olt_xls.parse(selected_olt_sheet, header=olt_header_idx)
                    base_olt_df = base_olt_df.loc[:, ~base_olt_df.columns.astype(str).str.startswith('Unnamed:')]
                    base_olt_df = base_olt_df.loc[:, base_olt_df.columns.notna() & (base_olt_df.columns != "")]
                    
                    # Ensure same columns are mapped exactly
                    append_df.columns = base_olt_df.columns
                    final_combined_df = pd.concat([base_olt_df, append_df], ignore_index=True)
                    
                    with pd.ExcelWriter(out_buffer, engine='openpyxl') as writer:
                        final_combined_df.to_excel(writer, sheet_name=selected_olt_sheet, index=False)
                    
                    out_buffer.seek(0)
                    append_success = True
                    st.success("🎉 Successfully processed using Fallback Data-Stream Matrix Engine!")
                except Exception as err_fallback:
                    st.error(f"Failed to append entries inside workbook structure: {err_fallback}")

            if append_success:
                st.download_button(
                    label="⬇️ Download Updated Nokia OLT Tracker File",
                    data=out_buffer.getvalue(),
                    file_name=f"Updated_{olt_file.name}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )