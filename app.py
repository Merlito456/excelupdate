import streamlit as st
import pandas as pd
import string
import io
import openpyxl
from datetime import datetime
import re

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

def highlight_duplicates(df, master_plaid_col):
    """Add a column to highlight potential duplicates"""
    df_with_highlight = df.copy()
    
    duplicate_mask = df_with_highlight[master_plaid_col].duplicated(keep=False)
    
    df_with_highlight['DUPLICATE_STATUS'] = ''
    df_with_highlight.loc[duplicate_mask, 'DUPLICATE_STATUS'] = '⚠️ POTENTIAL DUPLICATE'
    
    duplicate_counts = df_with_highlight[master_plaid_col].value_counts()
    df_with_highlight['DUPLICATE_COUNT'] = df_with_highlight[master_plaid_col].map(duplicate_counts)
    df_with_highlight.loc[~duplicate_mask, 'DUPLICATE_COUNT'] = 1
    
    def highlight_rows(row):
        if row['DUPLICATE_STATUS'] == '⚠️ POTENTIAL DUPLICATE':
            return ['background-color: #FFE5E5'] * len(row)
        return [''] * len(row)
    
    styled_df = df_with_highlight.style.apply(highlight_rows, axis=1)
    
    return df_with_highlight, styled_df

# -----------------------------
# 📊 DATA VALIDATION FUNCTIONS
# -----------------------------

def validate_data_type(series, expected_type):
    """
    Validates if the data in a series matches the expected type
    Returns: (is_valid, sample_values, description)
    """
    if len(series) == 0:
        return False, [], "Empty series"
    
    # Get non-null values for sampling
    sample = series.dropna().head(100)
    if len(sample) == 0:
        return False, [], "All values are null"
    
    sample_str = sample.astype(str).str.strip()
    sample_str = sample_str[sample_str != '']
    sample_str = sample_str[sample_str != 'nan']
    
    if len(sample_str) == 0:
        return False, [], "No valid values"
    
    # Check data patterns
    date_pattern = r'^\d{1,2}[/-]\d{1,2}[/-]\d{2,4}$|^\d{4}[/-]\d{1,2}[/-]\d{1,2}$|^\d{1,2}-[A-Za-z]{3}-\d{2,4}$'
    year_pattern = r'^\d{4}$'
    text_pattern = r'^[A-Za-z\s\-\.]+$'
    numeric_pattern = r'^\d+\.?\d*$'
    status_pattern = r'^(done|completed|pending|in progress|ongoing|on-going|cancelled|on hold)$'
    
    # Sample analysis
    date_matches = sample_str.str.contains(date_pattern, case=False, na=False).sum()
    year_matches = sample_str.str.contains(year_pattern, case=False, na=False).sum()
    text_matches = sample_str.str.contains(text_pattern, case=False, na=False).sum()
    numeric_matches = sample_str.str.contains(numeric_pattern, case=False, na=False).sum()
    status_matches = sample_str.str.contains(status_pattern, case=False, na=False).sum()
    
    total_valid = len(sample_str)
    
    # Determine the predominant type
    if date_matches / total_valid > 0.7:
        return True, sample_str.head(5).tolist(), "Dates"
    elif year_matches / total_valid > 0.7:
        return True, sample_str.head(5).tolist(), "Years"
    elif status_matches / total_valid > 0.7:
        return True, sample_str.head(5).tolist(), "Status values"
    elif numeric_matches / total_valid > 0.7:
        return True, sample_str.head(5).tolist(), "Numbers"
    elif text_matches / total_valid > 0.7:
        return True, sample_str.head(5).tolist(), "Text"
    else:
        return True, sample_str.head(5).tolist(), "Mixed data"

def verify_data_similarity(master_series, olt_series, threshold=0.6):
    """
    Verifies if two columns contain similar data patterns
    Returns: (similarity_score, is_similar, samples)
    """
    if len(master_series) == 0 or len(olt_series) == 0:
        return 0, False, []
    
    # Clean and prepare samples
    master_sample = master_series.dropna().astype(str).str.strip().head(50)
    olt_sample = olt_series.dropna().astype(str).str.strip().head(50)
    
    master_sample = master_sample[master_sample != '']
    master_sample = master_sample[master_sample != 'nan']
    olt_sample = olt_sample[olt_sample != '']
    olt_sample = olt_sample[olt_sample != 'nan']
    
    if len(master_sample) == 0 or len(olt_sample) == 0:
        return 0, False, []
    
    # Check for common patterns
    common_words_master = set(' '.join(master_sample).lower().split())
    common_words_olt = set(' '.join(olt_sample).lower().split())
    
    # Calculate similarity based on common words
    if len(common_words_master) > 0 and len(common_words_olt) > 0:
        intersection = len(common_words_master.intersection(common_words_olt))
        union = len(common_words_master.union(common_words_olt))
        word_similarity = intersection / union if union > 0 else 0
    else:
        word_similarity = 0
    
    # Check for value overlap (exact matches)
    common_values = set(master_sample).intersection(set(olt_sample))
    overlap_ratio = len(common_values) / max(len(set(master_sample)), len(set(olt_sample))) if len(set(master_sample)) > 0 and len(set(olt_sample)) > 0 else 0
    
    # Combined similarity score
    similarity_score = max(word_similarity, overlap_ratio * 0.8)
    is_similar = similarity_score >= threshold
    
    # Get sample common values
    sample_common = list(common_values)[:5] if common_values else []
    
    return similarity_score, is_similar, sample_common

# -----------------------------
# 📊 HEADER MAPPING WITH DATA VALIDATION
# -----------------------------

def get_header_mapping():
    """
    Defines the accurate mapping between Master Tracker and Nokia OLT Tracker headers
    Based on: Master Tracker Luzon_v3 - 02MARCH2026.xlsx and Nokia OLT Tracker v4.xlsx
    """
    
    mapping = {
        # Core Identifiers
        'PLAID': {
            'master_col': 'PLAID',
            'olt_col': 'PLAID',
            'type': 'key',
            'description': 'Primary Key Identifier',
            'required': True
        },
        
        # Site Information
        'SITE NAME': {
            'master_col': 'SITE NAME',
            'olt_col': 'SITE NAME',
            'type': 'text',
            'description': 'Site Name',
            'required': True
        },
        'SITE CODE': {
            'master_col': 'SITECODE',
            'olt_col': 'SITE CODE',
            'type': 'text',
            'description': 'Site Code',
            'required': False
        },
        
        # Project Information
        'BUILD YEAR': {
            'master_col': 'BUILD YEAR',
            'olt_col': 'BUILD YEAR',
            'type': 'date_text',
            'description': 'Build Year (formatted as "YYYY build")',
            'format': lambda x: f"{str(x).split('.')[0].strip()} build" if pd.notna(x) and str(x).strip() != "" and str(x).lower() != "nan" else "",
            'required': True
        },
        'PROJECT': {
            'master_col': 'PROJECT',
            'olt_col': 'PROJECT TAGGING',
            'type': 'text',
            'description': 'Project Tagging',
            'required': True
        },
        'PROJECT TYPE': {
            'master_col': 'OLT SCOPE',
            'olt_col': 'PROJECT TYPE',
            'type': 'text',
            'description': 'Project Type (from OLT Scope)',
            'required': True
        },
        
        # Location Information
        'CLUSTER': {
            'master_col': 'CLUSTER',
            'olt_col': 'CLUSTERING',
            'type': 'text',
            'description': 'Clustering/Cluster',
            'required': True
        },
        'PROVINCE': {
            'master_col': 'PROVINCE',
            'olt_col': 'PROVINCE',
            'type': 'text',
            'description': 'Province',
            'required': True
        },
        'REGION': {
            'master_col': 'REGION',
            'olt_col': 'REGION',
            'type': 'text',
            'description': 'Region',
            'required': False
        },
        
        # Equipment & Classification
        'EQUIPMENT TYPE': {
            'master_col': 'ELECTRONICS EQUIPMENT',
            'olt_col': 'EQUIPMENT TYPE',
            'type': 'text',
            'description': 'Equipment Type (from Electronics Equipment)',
            'required': True
        },
        'SITE CLASSIFICATION': {
            'master_col': 'SITE CLASSIFICATION',
            'olt_col': 'SITE CLASSIFICATION',
            'type': 'text',
            'description': 'Site Classification',
            'required': False
        },
        
        # Status Fields
        'SITE STATUS': {
            'master_col': 'SITE STATUS',
            'olt_col': 'SITE STATUS',
            'type': 'text',
            'description': 'Site Status',
            'required': True
        },
        'SCOPE STATUS': {
            'master_col': 'SCOPE STATUS',
            'olt_col': 'SCOPE STATUS',
            'type': 'text',
            'description': 'Scope Status',
            'required': False
        },
        
        # Milestone Dates
        'SURVEY DATE': {
            'master_col': 'SURVEY DATE',
            'olt_col': 'SITE SURVEY ACTUAL DATE',
            'type': 'date',
            'description': 'Site Survey Actual Date',
            'format': lambda x: str(x).split(" ")[0] if pd.notna(x) and str(x).lower() != "nan" else "",
            'required': False
        },
        'INSTALLED DATE': {
            'master_col': 'INSTALLED DATE',
            'olt_col': 'INSTALLATION DONE ACTUAL DATE',
            'type': 'date',
            'description': 'Installation Done Actual Date',
            'format': lambda x: str(x).split(" ")[0] if pd.notna(x) and str(x).lower() != "nan" else "",
            'required': False
        },
        'POWER TAPPED DATE': {
            'master_col': 'POWER TAPPED DATE',
            'olt_col': 'POWER TAPPING DONE ACTUAL DATE',
            'type': 'date',
            'description': 'Power Tapping Done Actual Date',
            'format': lambda x: str(x).split(" ")[0] if pd.notna(x) and str(x).lower() != "nan" else "",
            'required': False
        },
        'INTEGRATED DATE': {
            'master_col': 'INTEGRATED DATE',
            'olt_col': 'INTEGRATION DONE ACTUAL DATE',
            'type': 'date',
            'description': 'Integration Done Actual Date',
            'format': lambda x: str(x).split(" ")[0] if pd.notna(x) and str(x).lower() != "nan" else "",
            'required': False
        },
        'PAT DATE': {
            'master_col': 'PAT DATE',
            'olt_col': 'PAT DONE ACTUAL DATE',
            'type': 'date',
            'description': 'PAT Done Actual Date',
            'format': lambda x: str(x).split(" ")[0] if pd.notna(x) and str(x).lower() != "nan" else "",
            'required': False
        },
        'PAC DATE': {
            'master_col': "PAC'ED",
            'olt_col': 'PAC APPROVAL DONE ACTUAL DATE',
            'type': 'date',
            'description': 'PAC Approval Done Actual Date',
            'format': lambda x: str(x).split(" ")[0] if pd.notna(x) and str(x).lower() != "nan" else "",
            'required': False
        },
        'FAC DATE': {
            'master_col': "FAC'ED",
            'olt_col': 'FAC APPROVAL DONE ACTUAL DATE',
            'type': 'date',
            'description': 'FAC Approval Done Actual Date',
            'format': lambda x: str(x).split(" ")[0] if pd.notna(x) and str(x).lower() != "nan" else "",
            'required': False
        },
        
        # Additional Fields
        'HANDOVER DATE': {
            'master_col': 'HANDOVER DATE',
            'olt_col': 'HANDOVER DATE',
            'type': 'date',
            'description': 'Handover Date',
            'format': lambda x: str(x).split(" ")[0] if pd.notna(x) and str(x).lower() != "nan" else "",
            'required': False
        },
        'TARGET DATE': {
            'master_col': 'TARGET DATE',
            'olt_col': 'TARGET DATE',
            'type': 'date',
            'description': 'Target Date',
            'format': lambda x: str(x).split(" ")[0] if pd.notna(x) and str(x).lower() != "nan" else "",
            'required': False
        },
        'REMARKS': {
            'master_col': 'REMARKS',
            'olt_col': 'REMARKS',
            'type': 'text',
            'description': 'Remarks',
            'required': False
        }
    }
    
    return mapping

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
    # 📊 Header Mapping with Data Validation
    # -----------------------------
    st.subheader("🔍 Header Mapping with Data Validation")
    
    # Get mapping configuration
    header_mapping = get_header_mapping()
    
    # Perform data validation for each mapping
    validation_results = []
    
    for field, mapping in header_mapping.items():
        master_header = mapping['master_col']
        olt_header = mapping['olt_col']
        
        # Find actual columns
        master_col = None
        olt_col = None
        
        for col in master_df.columns:
            if clean_string_normalization(col) == clean_string_normalization(master_header):
                master_col = col
                break
        
        for col in olt_df.columns:
            if clean_string_normalization(col) == clean_string_normalization(olt_header):
                olt_col = col
                break
        
        result = {
            'Field': field,
            'Master Header': master_header,
            'OLT Header': olt_header,
            'Master Found': master_col is not None,
            'OLT Found': olt_col is not None,
            'Status': '❌ Missing',
            'Data Match': 'N/A',
            'Match Score': 0,
            'Master Sample': '',
            'OLT Sample': '',
            'Common Values': ''
        }
        
        if master_col and olt_col:
            # Validate data in both columns
            master_series = master_df[master_col]
            olt_series = olt_df[olt_col]
            
            # Check data types
            master_valid, master_sample, master_type = validate_data_type(master_series, mapping['type'])
            olt_valid, olt_sample, olt_type = validate_data_type(olt_series, mapping['type'])
            
            # Verify data similarity
            similarity_score, is_similar, common_values = verify_data_similarity(master_series, olt_series)
            
            result['Status'] = '✅ Matched' if is_similar else '⚠️ Check Data'
            result['Data Match'] = f"{master_type} ↔ {olt_type}"
            result['Match Score'] = f"{similarity_score:.2%}"
            result['Master Sample'] = ', '.join(master_sample[:3]) if master_sample else ''
            result['OLT Sample'] = ', '.join(olt_sample[:3]) if olt_sample else ''
            result['Common Values'] = ', '.join(common_values[:3]) if common_values else ''
        elif master_col:
            result['Status'] = '⚠️ Master Only'
        elif olt_col:
            result['Status'] = '⚠️ OLT Only'
        
        validation_results.append(result)
    
    # Display validation results
    validation_df = pd.DataFrame(validation_results)
    
    # Color coding for status
    def color_status(val):
        if '✅' in str(val):
            return 'background-color: #90EE90'
        elif '⚠️' in str(val):
            return 'background-color: #FFD700'
        elif '❌' in str(val):
            return 'background-color: #FF6B6B'
        return ''
    
    st.dataframe(validation_df.style.applymap(color_status, subset=['Status']), use_container_width=True)
    
    # Show detailed validation warnings
    warnings = validation_df[validation_df['Status'].str.contains('⚠️|❌', na=False)]
    if len(warnings) > 0:
        st.warning(f"⚠️ Found {len(warnings)} columns with potential mapping issues. Please review the data samples above.")
        
        with st.expander("📋 View Data Sample Comparison"):
            for idx, row in warnings.iterrows():
                st.write(f"**{row['Field']}**")
                col1, col2 = st.columns(2)
                with col1:
                    st.write(f"Master Sample: {row['Master Sample']}")
                with col2:
                    st.write(f"OLT Sample: {row['OLT Sample']}")
                if row['Common Values']:
                    st.write(f"Common values: {row['Common Values']}")
                st.write("---")

    # -----------------------------
    # 📊 All Entries with Duplicate Highlighting
    # -----------------------------
    st.write(f"📂 **Active Master Sheet:** `{selected_master_sheet}` | **Active OLT Sheet:** `{selected_olt_sheet}`")

    master_df.columns = master_df_cleaned.columns
    olt_df.columns = olt_df_cleaned.columns

    master_df[master_plaid_col_clean] = master_df[master_plaid_col_clean].astype(str).str.strip()
    olt_df[olt_plaid_col_clean] = olt_df[olt_plaid_col_clean].astype(str).str.strip()

    # Show all master entries with duplicate highlighting
    st.subheader("📊 All Master Entries with Duplicate Highlighting")
    
    master_with_dup, master_styled = highlight_duplicates(master_df, master_plaid_col_clean)
    
    total_entries = len(master_df)
    duplicate_count = master_df[master_plaid_col_clean].duplicated(keep=False).sum()
    unique_entries = master_df[master_plaid_col_clean].nunique()
    
    col1, col2, col3 = st.columns(3)
    col1.metric("Total Entries", total_entries)
    col2.metric("Unique Entries", unique_entries)
    col3.metric("Potential Duplicates", duplicate_count)
    
    st.dataframe(master_styled, use_container_width=True)
    
    if duplicate_count > 0:
        with st.expander("🔍 View Duplicate Details"):
            duplicates_only = master_with_dup[master_with_dup['DUPLICATE_STATUS'] == '⚠️ POTENTIAL DUPLICATE'].sort_values(master_plaid_col_clean)
            st.write(f"Found **{len(duplicates_only)}** rows with potential duplicates")
            st.dataframe(duplicates_only, use_container_width=True)

    # -----------------------------
    # 📉 Missing Records Analysis
    # -----------------------------
    master_clean_df = master_df[master_df[master_plaid_col_clean].str.lower() != "nan"].copy()
    missing_mask = ~master_clean_df[master_plaid_col_clean].isin(olt_df[olt_plaid_col_clean])
    missing_records = master_clean_df[missing_mask].copy()

    st.subheader("❌ Unmapped Raw Master Entries (Missing from OLT)")
    st.write(f"Total Missing Rows Isolated: **{len(missing_records)}**")
    
    if len(missing_records) > 0:
        missing_with_dup, missing_styled = highlight_duplicates(missing_records, master_plaid_col_clean)
        st.dataframe(missing_styled, use_container_width=True)

    # -----------------------------
    # 🔄 Automated Mapping with Data Verification
    # -----------------------------
    st.subheader("🔄 Automated Column Mapping with Data Verification")
    
    # Create append DataFrame with OLT columns
    append_df = pd.DataFrame(columns=orig_olt_cols)
    mapped_columns_log = []
    mapping_issues = []
    
    # Process each OLT column based on the mapping
    for c_idx, orig_olt_col in enumerate(orig_olt_cols):
        clean_olt_name = clean_string_normalization(orig_olt_col)
        
        # Skip empty or track columns
        if clean_olt_name == "" or "solution track" in clean_olt_name or clean_olt_name == "track":
            continue
        
        matched_master_col = None
        formatting_func = None
        match_confidence = "Low"
        verification_note = ""
        
        # First, check if this OLT column is in our mapping
        for field, mapping in header_mapping.items():
            if clean_string_normalization(mapping['olt_col']) == clean_olt_name:
                # Found a match in our mapping
                master_header = mapping['master_col']
                # Find the actual column in master_df
                for clean_m_col in master_df.columns:
                    if clean_string_normalization(clean_m_col) == clean_string_normalization(master_header):
                        matched_master_col = clean_m_col
                        formatting_func = mapping.get('format', None)
                        
                        # Verify data similarity
                        master_series = master_df[matched_master_col]
                        olt_series = olt_df[orig_olt_col]
                        similarity_score, is_similar, common_values = verify_data_similarity(master_series, olt_series)
                        
                        match_confidence = "High" if similarity_score > 0.8 else "Medium" if similarity_score > 0.6 else "Low"
                        
                        if is_similar:
                            verification_note = f"✅ Data verified (Score: {similarity_score:.2%})"
                            # Check required fields
                            if mapping.get('required', False) and not is_similar:
                                mapping_issues.append(f"⚠️ Required field '{field}' has low data similarity ({similarity_score:.2%})")
                        else:
                            verification_note = f"⚠️ Data mismatch (Score: {similarity_score:.2%}) - Check values"
                            mapping_issues.append(f"⚠️ Field '{field}' data doesn't match between files (Score: {similarity_score:.2%})")
                        
                        break
                if matched_master_col:
                    mapped_columns_log.append({
                        'olt_col': orig_olt_col,
                        'master_col': matched_master_col,
                        'field': field,
                        'confidence': match_confidence,
                        'verification': verification_note,
                        'description': mapping['description']
                    })
                break
        
        # If not in mapping, try to match by column name with data verification
        if not matched_master_col:
            best_match = None
            best_score = 0
            
            for clean_m_col in master_df.columns:
                # Check name similarity
                if clean_string_normalization(clean_m_col) == clean_olt_name:
                    best_match = clean_m_col
                    best_score = 1.0
                    break
                
                # Check partial match
                if clean_olt_name in clean_string_normalization(clean_m_col) or clean_string_normalization(clean_m_col) in clean_olt_name:
                    # Verify data similarity for this potential match
                    master_series = master_df[clean_m_col]
                    olt_series = olt_df[orig_olt_col]
                    similarity_score, is_similar, common_values = verify_data_similarity(master_series, olt_series)
                    
                    if similarity_score > best_score:
                        best_score = similarity_score
                        best_match = clean_m_col
            
            if best_match and best_score > 0.4:
                matched_master_col = best_match
                match_confidence = "High" if best_score > 0.8 else "Medium" if best_score > 0.6 else "Low"
                verification_note = f"🔀 Auto-matched (Score: {best_score:.2%})"
                mapped_columns_log.append({
                    'olt_col': orig_olt_col,
                    'master_col': matched_master_col,
                    'field': 'Auto-matched',
                    'confidence': match_confidence,
                    'verification': verification_note,
                    'description': 'Auto-detected by name & data similarity'
                })
        
        # Handle the PLAID column specially
        if clean_olt_name == clean_string_normalization('PLAID'):
            matched_master_col = master_plaid_col_clean
            match_confidence = "High"
            verification_note = "🔑 Key identifier verified"
            mapped_columns_log.append({
                'olt_col': orig_olt_col,
                'master_col': matched_master_col,
                'field': 'PLAID',
                'confidence': match_confidence,
                'verification': verification_note,
                'description': 'Primary Key Identifier'
            })
        
        # Apply the mapping
        if matched_master_col:
            # Apply any formatting function if defined
            if formatting_func:
                raw_values = missing_records[matched_master_col].tolist()
                append_df[orig_olt_col] = [formatting_func(val) for val in raw_values]
            else:
                append_df[orig_olt_col] = missing_records[matched_master_col].tolist()
        else:
            append_df[orig_olt_col] = [""] * len(missing_records)
            mapped_columns_log.append({
                'olt_col': orig_olt_col,
                'master_col': 'No match found',
                'field': 'Unmapped',
                'confidence': 'None',
                'verification': '❌ Column will be left blank',
                'description': 'No matching column found'
            })
    
    # Display mapping warnings
    if mapping_issues:
        st.warning("⚠️ Data Verification Issues Found")
        for issue in mapping_issues:
            st.write(f"- {issue}")
    
    # Display mapping audit trail
    with st.expander("👀 View Detailed Column Mapping Audit Trail"):
        mapping_df = pd.DataFrame(mapped_columns_log)
        # Color code confidence levels
        def color_confidence(val):
            if val == 'High':
                return 'background-color: #90EE90'
            elif val == 'Medium':
                return 'background-color: #FFD700'
            elif val == 'Low':
                return 'background-color: #FFA500'
            return ''
        
        st.dataframe(mapping_df.style.applymap(color_confidence, subset=['confidence']), use_container_width=True)
    
    # Exclude formatting column anomalies
    append_df = append_df.loc[:, ~append_df.columns.astype(str).str.contains('track', case=False)]
    
    st.subheader("📋 Output Blueprint (Ready to append into Nokia OLT)")
    st.dataframe(append_df.head(100), use_container_width=True)
    
    # Show preview of filled data with verification
    if len(append_df) > 0:
        st.subheader("📊 Sample Filled Data Preview with Verification")
        
        # Show which columns have data
        filled_cols = []
        for col in append_df.columns:
            non_null_count = append_df[col].notna().sum()
            if non_null_count > 0:
                filled_cols.append({
                    'Column': col,
                    'Filled Rows': non_null_count,
                    'Fill Rate': f"{non_null_count/len(append_df)*100:.1f}%"
                })
        
        if filled_cols:
            filled_df = pd.DataFrame(filled_cols)
            st.dataframe(filled_df, use_container_width=True)
        
        # Preview data
        preview_cols = [col for col in append_df.columns if append_df[col].notna().any()][:10]
        if preview_cols:
            st.dataframe(append_df[preview_cols].head(10), use_container_width=True)

    # -----------------------------
    # 💾 Download Merged File
    # -----------------------------
    if len(append_df) > 0:
        if st.button("🚀 Merge and Append into OLT Spreadsheet"):
            try:
                base_olt_df = olt_xls.parse(selected_olt_sheet, header=olt_header_idx)
                
                # Remove any rows with matching PLAID to avoid duplicates
                base_plaid_col = chosen_olt_plaid_raw
                append_plaid_col = chosen_olt_plaid_raw
                
                # Filter out rows that already exist in OLT
                existing_plaids = set(base_olt_df[base_plaid_col].astype(str).str.strip())
                append_plaid_values = append_df[append_plaid_col].astype(str).str.strip()
                new_records_mask = ~append_plaid_values.isin(existing_plaids)
                new_records = append_df[new_records_mask]
                
                # Also filter out records with empty PLAID
                new_records = new_records[new_records[append_plaid_col].astype(str).str.strip() != '']
                new_records = new_records[new_records[append_plaid_col].astype(str).str.strip() != 'nan']
                
                if len(new_records) > 0:
                    # Data verification summary
                    st.info(f"📊 **Data Summary:**\n"
                           f"- Total missing records: {len(append_df)}\n"
                           f"- New records to add: {len(new_records)}\n"
                           f"- Skipped existing records: {len(append_df) - len(new_records)}")
                    
                    final_combined_df = pd.concat([base_olt_df, new_records], ignore_index=True)
                    
                    out_buffer = io.BytesIO()
                    with pd.ExcelWriter(out_buffer, engine='openpyxl') as writer:
                        final_combined_df.to_excel(writer, sheet_name=selected_olt_sheet, index=False)
                    
                    out_buffer.seek(0)
                    
                    st.success(f"🎉 Successfully merged and processed {len(new_records)} new records!")
                    st.download_button(
                        label="⬇️ Download Updated Nokia OLT Tracker File",
                        data=out_buffer.getvalue(),
                        file_name=f"Updated_{olt_file.name}",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.warning("⚠️ All records already exist in the OLT tracker or have invalid PLAID values. No new records to append.")
                    
            except Exception as err:
                st.error(f"Failed to process file: {err}")