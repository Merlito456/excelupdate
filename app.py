import streamlit as st
import pandas as pd
import string
import io
import openpyxl
from datetime import datetime
import re
from collections import Counter

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
# 📊 DATA CATEGORY DETECTION FUNCTIONS
# -----------------------------

def detect_data_category(series):
    """
    Detects the category/type of data in a series without looking at specific values
    Returns: (category, confidence, sample_values)
    """
    if len(series) == 0:
        return "Empty", 0, []
    
    # Get non-null values for sampling
    sample = series.dropna().head(100)
    if len(sample) == 0:
        return "Empty", 0, []
    
    sample_str = sample.astype(str).str.strip()
    sample_str = sample_str[sample_str != '']
    sample_str = sample_str[sample_str != 'nan']
    
    if len(sample_str) == 0:
        return "Empty", 0, []
    
    # Define patterns for different data categories
    patterns = {
        'PLAID/ID': r'^[A-Z0-9\-_]+$',
        'Date': r'^\d{1,2}[/-]\d{1,2}[/-]\d{2,4}$|^\d{4}[/-]\d{1,2}[/-]\d{1,2}$|^\d{1,2}-[A-Za-z]{3}-\d{2,4}$|^\d{4}-\d{2}-\d{2}$',
        'Year': r'^\d{4}$',
        'Text/Name': r'^[A-Za-z\s\-\.]+$',
        'Numeric': r'^\d+\.?\d*$',
        'Status': r'^(done|completed|pending|in progress|ongoing|on-going|cancelled|on hold|active|inactive|new|open|closed)$',
        'Project Type': r'^(OLT|MSAG|FTTH|FTTX|GPON|NGN|RAN|MW|SDH|DWDM)$',
        'Equipment Type': r'^(OLT|MSAG|MDU|SFU|HGU|ONU|ONT|Switch|Router|Gateway)$',
        'Cluster/Area': r'^[A-Z]{2,4}-?\d*$|^[A-Za-z]+\s+[0-9]+$',
        'Province/Region': r'^[A-Z][a-z]+(\s+[A-Z][a-z]+)*$',
        'Site Name': r'^[A-Za-z\s\-\.]+$',
        'Build Year': r'^20\d{2}$'
    }
    
    # Count matches for each category
    category_scores = {}
    for category, pattern in patterns.items():
        matches = sample_str.str.contains(pattern, case=False, na=False).sum()
        if len(sample_str) > 0:
            score = matches / len(sample_str)
            if score > 0.5:  # Only consider if more than 50% match
                category_scores[category] = score
    
    # If no category matches well, try to detect based on data characteristics
    if not category_scores:
        # Check if it's mixed text with numbers
        has_letters = sample_str.str.contains('[A-Za-z]', na=False).sum() > 0
        has_numbers = sample_str.str.contains('\d', na=False).sum() > 0
        
        if has_letters and has_numbers:
            return "Mixed (Text+Numbers)", 0.3, sample_str.head(5).tolist()
        elif has_letters:
            return "Text", 0.3, sample_str.head(5).tolist()
        elif has_numbers:
            return "Numeric", 0.3, sample_str.head(5).tolist()
        else:
            return "Unknown", 0, sample_str.head(5).tolist()
    
    # Get the best matching category
    best_category = max(category_scores, key=category_scores.get)
    best_score = category_scores[best_category]
    
    return best_category, best_score, sample_str.head(5).tolist()

def verify_category_match(master_category, master_score, olt_category, olt_score):
    """
    Verifies if two data categories match
    Returns: (is_match, confidence, description)
    """
    # Define category groups that are compatible
    compatible_groups = {
        'Identifiers': ['PLAID/ID'],
        'Dates': ['Date'],
        'Years': ['Year', 'Build Year'],
        'Text': ['Text/Name', 'Site Name', 'Text'],
        'Status': ['Status'],
        'Locations': ['Province/Region', 'Cluster/Area'],
        'Equipment': ['Equipment Type'],
        'Project': ['Project Type'],
        'Numbers': ['Numeric']
    }
    
    # Check if categories are in the same group
    master_group = None
    olt_group = None
    
    for group, categories in compatible_groups.items():
        if master_category in categories:
            master_group = group
        if olt_category in categories:
            olt_group = group
    
    # If both categories are in the same group, they're compatible
    if master_group and olt_group and master_group == olt_group:
        avg_score = (master_score + olt_score) / 2
        if avg_score >= 0.8:
            return True, "High", f"Both are {master_category} ({avg_score:.0%} confidence)"
        elif avg_score >= 0.6:
            return True, "Medium", f"Both are {master_category} ({avg_score:.0%} confidence)"
        else:
            return True, "Low", f"Both are {master_category} but with low confidence ({avg_score:.0%})"
    
    # Special case: Years and Dates are often compatible
    if (master_category in ['Year', 'Build Year'] and olt_category == 'Date') or \
       (master_category == 'Date' and olt_category in ['Year', 'Build Year']):
        avg_score = (master_score + olt_score) / 2
        return True, "Medium", f"Years and Dates are compatible ({avg_score:.0%} confidence)"
    
    # Check if categories are semantically similar
    semantic_pairs = [
        ('Text/Name', 'Site Name'),
        ('Text/Name', 'Text'),
        ('Province/Region', 'Text/Name'),
        ('Cluster/Area', 'Text/Name'),
        ('Equipment Type', 'Text/Name'),
        ('Project Type', 'Text/Name')
    ]
    
    if (master_category, olt_category) in semantic_pairs or (olt_category, master_category) in semantic_pairs:
        avg_score = (master_score + olt_score) / 2
        return True, "Low", f"Semantically similar categories ({avg_score:.0%} confidence)"
    
    # If categories are completely different
    return False, "None", f"Category mismatch: {master_category} ↔ {olt_category}"

# -----------------------------
# 📊 DATA SIMILARITY & REPEATED VALUES ANALYSIS
# -----------------------------

def analyze_data_patterns(series):
    """
    Analyzes data patterns including repeated values, unique values, and distributions
    Returns: (pattern_summary, repeated_values, uniqueness_score)
    """
    if len(series) == 0:
        return "Empty", {}, 0
    
    # Clean the data
    clean_series = series.dropna().astype(str).str.strip()
    clean_series = clean_series[clean_series != '']
    clean_series = clean_series[clean_series != 'nan']
    
    if len(clean_series) == 0:
        return "All empty", {}, 0
    
    # Get value counts
    value_counts = Counter(clean_series)
    total_values = len(clean_series)
    unique_values = len(value_counts)
    
    # Calculate uniqueness score (0-1, where 1 means all unique)
    uniqueness_score = unique_values / total_values if total_values > 0 else 0
    
    # Find repeated values (appear more than once)
    repeated_values = {k: v for k, v in value_counts.items() if v > 1}
    
    # Get the most common repeated values
    top_repeated = dict(sorted(repeated_values.items(), key=lambda x: x[1], reverse=True)[:10])
    
    # Determine pattern type
    if uniqueness_score > 0.9:
        pattern_type = "Highly Unique (IDs/Keys)"
    elif uniqueness_score > 0.7:
        pattern_type = "Mostly Unique (Mixed)"
    elif uniqueness_score > 0.4:
        pattern_type = "Moderately Repeated (Categories)"
    elif uniqueness_score > 0.1:
        pattern_type = "Highly Repeated (Status/Type)"
    else:
        pattern_type = "Almost Constant (Single Value)"
    
    pattern_summary = {
        'pattern_type': pattern_type,
        'total_values': total_values,
        'unique_values': unique_values,
        'uniqueness_score': uniqueness_score,
        'repeated_count': len(repeated_values),
        'most_repeated': top_repeated
    }
    
    return pattern_summary, top_repeated, uniqueness_score

def verify_data_similarity_with_patterns(master_series, olt_series, threshold=0.6):
    """
    Enhanced verification that checks both category patterns and actual value overlap
    Returns: (similarity_score, is_similar, details, common_values, pattern_match)
    """
    if len(master_series) == 0 or len(olt_series) == 0:
        return 0, False, "Empty series", [], False
    
    # Get pattern analysis for both series
    master_pattern, master_repeated, master_uniqueness = analyze_data_patterns(master_series)
    olt_pattern, olt_repeated, olt_uniqueness = analyze_data_patterns(olt_series)
    
    # Clean data for comparison
    master_clean = master_series.dropna().astype(str).str.strip()
    master_clean = master_clean[master_clean != '']
    master_clean = master_clean[master_clean != 'nan']
    
    olt_clean = olt_series.dropna().astype(str).str.strip()
    olt_clean = olt_clean[olt_clean != '']
    olt_clean = olt_clean[olt_clean != 'nan']
    
    if len(master_clean) == 0 or len(olt_clean) == 0:
        return 0, False, "No valid values to compare", [], False
    
    # 1. Check pattern similarity (uniqueness and repetition patterns)
    uniqueness_diff = abs(master_uniqueness - olt_uniqueness)
    pattern_similarity = 1 - uniqueness_diff if uniqueness_diff <= 1 else 0
    
    # 2. Check value overlap (actual data similarity)
    common_values = set(master_clean.head(100)).intersection(set(olt_clean.head(100)))
    overlap_ratio = len(common_values) / max(len(set(master_clean.head(100))), len(set(olt_clean.head(100)))) if len(set(master_clean.head(100))) > 0 and len(set(olt_clean.head(100))) > 0 else 0
    
    # 3. Check repeated value patterns
    master_top_repeated = Counter(master_clean).most_common(5)
    olt_top_repeated = Counter(olt_clean).most_common(5)
    
    # Check if top repeated values are similar
    repeated_overlap = 0
    for m_val, m_count in master_top_repeated:
        for o_val, o_count in olt_top_repeated:
            if clean_string_normalization(m_val) == clean_string_normalization(o_val):
                repeated_overlap += 1
                break
    
    repeated_similarity = repeated_overlap / max(len(master_top_repeated), len(olt_top_repeated)) if max(len(master_top_repeated), len(olt_top_repeated)) > 0 else 0
    
    # Combined similarity score
    similarity_score = (pattern_similarity * 0.3) + (overlap_ratio * 0.4) + (repeated_similarity * 0.3)
    
    # Determine if similar
    is_similar = similarity_score >= threshold
    
    # Build details
    details = {
        'pattern_similarity': pattern_similarity,
        'overlap_ratio': overlap_ratio,
        'repeated_similarity': repeated_similarity,
        'master_pattern': master_pattern['pattern_type'],
        'olt_pattern': olt_pattern['pattern_type'],
        'master_uniqueness': master_uniqueness,
        'olt_uniqueness': olt_uniqueness,
        'master_repeated_count': len(master_repeated),
        'olt_repeated_count': len(olt_repeated)
    }
    
    pattern_match = (master_pattern['pattern_type'] == olt_pattern['pattern_type']) or \
                   (master_uniqueness > 0.7 and olt_uniqueness > 0.7) or \
                   (master_uniqueness < 0.3 and olt_uniqueness < 0.3)
    
    return similarity_score, is_similar, details, list(common_values)[:10], pattern_match

# -----------------------------
# 📊 HEADER INSPECTION AND RENAMING UI
# -----------------------------

def header_inspection_ui(df, df_name, key_prefix):
    """
    Creates an interactive UI for inspecting and renaming column headers
    Returns: dictionary mapping original to new column names
    """
    st.write(f"### 🔍 Inspect and Rename Headers - {df_name}")
    st.info(f"Total columns: {len(df.columns)}")
    
    # Initialize session state for column renaming
    rename_key = f"{key_prefix}_rename_mapping"
    if rename_key not in st.session_state:
        st.session_state[rename_key] = {col: col for col in df.columns}
    
    # Allow user to select a column to inspect
    selected_col = st.selectbox(
        f"Select a column from {df_name} to inspect:",
        options=df.columns,
        key=f"{key_prefix}_select_col"
    )
    
    if selected_col:
        # Show data preview for selected column
        st.write(f"#### 📊 Data Preview: {selected_col}")
        
        # Get data samples
        sample_data = df[selected_col].dropna().head(20)
        sample_values = sample_data.tolist()
        
        # Show data statistics
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Non-null Count", df[selected_col].count())
        with col2:
            st.metric("Unique Values", df[selected_col].nunique())
        with col3:
            null_count = df[selected_col].isna().sum()
            st.metric("Null Count", null_count)
        with col4:
            # Detect data category
            category, confidence, _ = detect_data_category(df[selected_col])
            st.metric("Data Category", category)
        
        # Show sample data in a table
        st.write("**Sample Values (First 20 non-null):**")
        if len(sample_values) > 0:
            sample_df = pd.DataFrame({
                'Row': range(1, len(sample_values) + 1),
                'Value': sample_values
            })
            st.dataframe(sample_df, use_container_width=True)
        else:
            st.warning("No non-null values found in this column")
        
        # Show value distribution for repeated values
        if df[selected_col].nunique() < 50 and df[selected_col].count() > 0:
            st.write("**Value Distribution (Top 10):**")
            value_counts = df[selected_col].value_counts().head(10)
            dist_df = pd.DataFrame({
                'Value': value_counts.index,
                'Count': value_counts.values,
                'Percentage': (value_counts.values / len(df) * 100).round(1)
            })
            st.dataframe(dist_df, use_container_width=True)
        
        # Allow column renaming
        st.write("#### ✏️ Rename This Column")
        
        col1, col2 = st.columns([2, 1])
        with col1:
            new_name = st.text_input(
                f"New name for '{selected_col}':",
                value=st.session_state[rename_key].get(selected_col, selected_col),
                key=f"{key_prefix}_rename_input_{selected_col}"
            )
        with col2:
            if st.button(f"Apply Rename", key=f"{key_prefix}_apply_rename_{selected_col}"):
                if new_name and new_name != selected_col:
                    # Check if new name already exists
                    if new_name in df.columns and new_name != selected_col:
                        st.warning(f"⚠️ Column '{new_name}' already exists. Please choose a different name.")
                    else:
                        # Update the rename mapping
                        st.session_state[rename_key][selected_col] = new_name
                        st.success(f"✅ Column renamed to '{new_name}'")
                        st.rerun()
                elif new_name == selected_col:
                    st.info("No change made. New name is the same as original.")
                else:
                    st.warning("Please enter a valid column name.")
    
    # Show all column mappings with current names
    st.write("#### 📋 Current Column Name Mapping")
    
    mapping_data = []
    for orig_col in df.columns:
        new_col = st.session_state[rename_key].get(orig_col, orig_col)
        status = "✅ Renamed" if new_col != orig_col else "Original"
        mapping_data.append({
            'Original Name': orig_col,
            'Current Name': new_col,
            'Status': status
        })
    
    mapping_df = pd.DataFrame(mapping_data)
    st.dataframe(mapping_df, use_container_width=True)
    
    # Button to reset all renames
    if st.button(f"🔄 Reset All Column Names for {df_name}", key=f"{key_prefix}_reset"):
        st.session_state[rename_key] = {col: col for col in df.columns}
        st.success("✅ All column names reset to original")
        st.rerun()
    
    # Return the rename mapping
    return st.session_state[rename_key]

# -----------------------------
# 📊 HEADER MAPPING WITH COMPREHENSIVE VERIFICATION
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

    # Apply cleaned column names to dataframes
    master_df.columns = master_df_cleaned.columns
    olt_df.columns = olt_df_cleaned.columns

    # -----------------------------
    # 📊 HEADER INSPECTION AND RENAMING
    # -----------------------------
    st.subheader("🔍 Header Inspection and Renaming")
    st.info("Inspect the data in each column and rename headers if needed. This helps ensure accurate column mapping.")
    
    # Create tabs for Master and OLT header inspection
    tab1, tab2 = st.tabs(["📋 Master Tracker Headers", "📋 OLT Tracker Headers"])
    
    with tab1:
        master_rename_mapping = header_inspection_ui(master_df, "Master Tracker", "master")
        # Apply renaming to master_df
        if master_rename_mapping:
            new_master_columns = []
            for col in master_df.columns:
                new_name = master_rename_mapping.get(col, col)
                new_master_columns.append(new_name)
            master_df.columns = new_master_columns
    
    with tab2:
        olt_rename_mapping = header_inspection_ui(olt_df, "OLT Tracker", "olt")
        # Apply renaming to olt_df
        if olt_rename_mapping:
            new_olt_columns = []
            for col in olt_df.columns:
                new_name = olt_rename_mapping.get(col, col)
                new_olt_columns.append(new_name)
            olt_df.columns = new_olt_columns
    
    # Show updated column lists
    st.write("### 📊 Updated Column Lists")
    col1, col2 = st.columns(2)
    with col1:
        st.write("**Master Tracker Columns:**")
        st.write(", ".join(master_df.columns.tolist()))
    with col2:
        st.write("**OLT Tracker Columns:**")
        st.write(", ".join(olt_df.columns.tolist()))
    
    # Confirm after inspection
    st.write("### ✅ Confirm After Header Inspection")
    confirm_inspection = st.checkbox(
        "I have inspected and renamed headers as needed. Proceed to mapping verification.",
        key="confirm_inspection"
    )
    
    if not confirm_inspection:
        st.info("Please inspect and verify the headers before proceeding.")
        st.stop()

    # -----------------------------
    # 📊 Header Mapping with Comprehensive Verification
    # -----------------------------
    st.subheader("🔍 Comprehensive Column Mapping Verification")
    st.info("This tool verifies column mappings using:")
    st.info("1. **Data Category/Type** - What kind of data is in the column")
    st.info("2. **Data Pattern Analysis** - Uniqueness and repetition patterns")
    st.info("3. **Actual Data Similarity** - Overlap in actual values")

    # Get mapping configuration
    header_mapping = get_header_mapping()
    
    # Create interactive mapping verification
    st.write("### 📋 Comprehensive Column Mapping Verification")
    
    # Initialize session state for mapping verification
    if 'mapping_confirmed' not in st.session_state:
        st.session_state.mapping_confirmed = False
        st.session_state.manual_mappings = {}
    
    # Display current mappings with comprehensive verification
    mapping_verified = True
    mapping_df_data = []
    
    for field, mapping in header_mapping.items():
        master_header = mapping['master_col']
        olt_header = mapping['olt_col']
        
        # Find actual columns (using current renamed columns)
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
        
        # Check if user has manually overridden this mapping
        mapping_key = f"{field}_{master_header}_{olt_header}"
        if mapping_key in st.session_state.manual_mappings:
            manual_master = st.session_state.manual_mappings[mapping_key].get('master')
            manual_olt = st.session_state.manual_mappings[mapping_key].get('olt')
            if manual_master:
                master_col = manual_master
            if manual_olt:
                olt_col = manual_olt
        
        # Get comprehensive verification
        master_category = "N/A"
        master_confidence = 0
        olt_category = "N/A"
        olt_confidence = 0
        master_samples = []
        olt_samples = []
        pattern_match = False
        similarity_score = 0
        is_similar = False
        details = {}
        common_values = []
        category_match = "❌ Not Found"
        match_confidence = "None"
        match_description = ""
        
        if master_col:
            master_category, master_confidence, master_samples = detect_data_category(master_df[master_col])
        
        if olt_col:
            olt_category, olt_confidence, olt_samples = detect_data_category(olt_df[olt_col])
        
        if master_col and olt_col:
            # Category verification
            is_cat_match, cat_conf, cat_desc = verify_category_match(
                master_category, master_confidence,
                olt_category, olt_confidence
            )
            
            # Data similarity with pattern analysis
            similarity_score, is_similar, details, common_values, pattern_match = verify_data_similarity_with_patterns(
                master_df[master_col],
                olt_df[olt_col]
            )
            
            # Combined verification
            if is_cat_match and is_similar and pattern_match:
                category_match = "✅ All Good"
                match_confidence = "High"
                match_description = f"Category: {cat_desc}, Similarity: {similarity_score:.1%}"
            elif is_cat_match and (is_similar or pattern_match):
                category_match = "✅ Acceptable"
                match_confidence = "Medium"
                match_description = f"Category: {cat_desc}, Similarity: {similarity_score:.1%}"
            elif is_cat_match:
                category_match = "⚠️ Category OK, Data mismatch"
                match_confidence = "Low"
                match_description = f"Category matches but data patterns differ (Sim: {similarity_score:.1%})"
                mapping_verified = False
            else:
                category_match = "⚠️ Check Mapping"
                match_confidence = "Low"
                match_description = f"Category mismatch: {master_category} ↔ {olt_category}"
                mapping_verified = False
        elif master_col:
            category_match = "⚠️ Master Only"
            mapping_verified = False
        elif olt_col:
            category_match = "⚠️ OLT Only"
            mapping_verified = False
        else:
            category_match = "❌ Both Missing"
            mapping_verified = False
        
        mapping_df_data.append({
            'Field': field,
            'Description': mapping['description'],
            'Required': '✅' if mapping.get('required', False) else '',
            'Master Column': master_col or '❌ NOT FOUND',
            'OLT Column': olt_col or '❌ NOT FOUND',
            'Master Category': f"{master_category} ({master_confidence:.0%})" if master_col else 'N/A',
            'OLT Category': f"{olt_category} ({olt_confidence:.0%})" if olt_col else 'N/A',
            'Pattern Match': '✅' if pattern_match else '⚠️' if master_col and olt_col else 'N/A',
            'Data Similarity': f"{similarity_score:.1%}" if master_col and olt_col else 'N/A',
            'Common Values': ', '.join(common_values[:3]) if common_values else 'None',
            'Category Match': category_match,
            'Match Details': match_description if master_col and olt_col else ''
        })
    
    # Display mapping table
    mapping_df = pd.DataFrame(mapping_df_data)
    
    # Color coding function
    def color_mapping_status(val):
        if '✅ All Good' in str(val):
            return 'background-color: #90EE90'
        elif '✅ Acceptable' in str(val):
            return 'background-color: #98FB98'
        elif '⚠️' in str(val):
            return 'background-color: #FFD700'
        elif '❌' in str(val):
            return 'background-color: #FF6B6B'
        return ''
    
    styled_mapping = mapping_df.style.map(color_mapping_status, subset=['Category Match'])
    st.dataframe(styled_mapping, use_container_width=True)
    
    # Show data pattern analysis for critical columns
    with st.expander("📊 Detailed Data Pattern Analysis"):
        st.write("This shows the data patterns for each column, including uniqueness and repetition patterns.")
        
        # Analyze PLAID column
        if master_plaid_col_clean and olt_plaid_col_clean:
            st.write("**PLAID Column Pattern Analysis:**")
            col1, col2 = st.columns(2)
            
            with col1:
                st.write("**Master PLAID Patterns:**")
                master_pattern, master_repeated, master_unique = analyze_data_patterns(master_df[master_plaid_col_clean])
                st.write(f"- Pattern Type: {master_pattern['pattern_type']}")
                st.write(f"- Unique Values: {master_pattern['unique_values']}")
                st.write(f"- Uniqueness Score: {master_pattern['uniqueness_score']:.1%}")
                if master_repeated:
                    st.write("- Top Repeated Values:")
                    for val, count in list(master_repeated.items())[:3]:
                        st.write(f"  - {val}: {count} times")
            
            with col2:
                st.write("**OLT PLAID Patterns:**")
                olt_pattern, olt_repeated, olt_unique = analyze_data_patterns(olt_df[olt_plaid_col_clean])
                st.write(f"- Pattern Type: {olt_pattern['pattern_type']}")
                st.write(f"- Unique Values: {olt_pattern['unique_values']}")
                st.write(f"- Uniqueness Score: {olt_pattern['uniqueness_score']:.1%}")
                if olt_repeated:
                    st.write("- Top Repeated Values:")
                    for val, count in list(olt_repeated.items())[:3]:
                        st.write(f"  - {val}: {count} times")
        
        # Analyze status columns
        status_mappings = ['SITE STATUS', 'SCOPE STATUS']
        for field in status_mappings:
            if field in header_mapping:
                mapping = header_mapping[field]
                master_header = mapping['master_col']
                olt_header = mapping['olt_col']
                
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
                
                if master_col and olt_col:
                    st.write(f"\n**{field} Pattern Analysis:**")
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        master_pattern, master_repeated, _ = analyze_data_patterns(master_df[master_col])
                        st.write(f"Master: {master_pattern['pattern_type']}")
                        if master_repeated:
                            st.write("Top values:", ', '.join([f"{k}({v})" for k, v in list(master_repeated.items())[:3]]))
                    
                    with col2:
                        olt_pattern, olt_repeated, _ = analyze_data_patterns(olt_df[olt_col])
                        st.write(f"OLT: {olt_pattern['pattern_type']}")
                        if olt_repeated:
                            st.write("Top values:", ', '.join([f"{k}({v})" for k, v in list(olt_repeated.items())[:3]]))
    
    # Show warnings for missing or mismatched mappings
    issues = mapping_df[
        (mapping_df['Category Match'].str.contains('⚠️|❌', na=False)) |
        (mapping_df['Master Column'].str.contains('NOT FOUND', na=False)) |
        (mapping_df['OLT Column'].str.contains('NOT FOUND', na=False))
    ]
    
    if len(issues) > 0:
        st.warning(f"⚠️ Found {len(issues)} fields with mapping issues. Please review them above.")
    
    # Manual mapping override section
    with st.expander("🛠️ Manual Mapping Override (Advanced)"):
        st.write("If automatic mapping is incorrect, you can manually select the correct columns below:")
        
        # Get all available columns
        master_columns = list(master_df.columns)
        olt_columns = list(olt_df.columns)
        
        # For each field that needs verification, provide dropdowns
        for field, mapping in header_mapping.items():
            # Find current mapping
            current_master = None
            current_olt = None
            
            for col in master_df.columns:
                if clean_string_normalization(col) == clean_string_normalization(mapping['master_col']):
                    current_master = col
                    break
            
            for col in olt_df.columns:
                if clean_string_normalization(col) == clean_string_normalization(mapping['olt_col']):
                    current_olt = col
                    break
            
            col1, col2, col3 = st.columns([2, 2, 1])
            with col1:
                selected_master = st.selectbox(
                    f"Master column for '{field}'",
                    options=['None'] + master_columns,
                    index=0 if current_master is None else master_columns.index(current_master) + 1,
                    key=f"manual_master_{field}"
                )
            with col2:
                selected_olt = st.selectbox(
                    f"OLT column for '{field}'",
                    options=['None'] + olt_columns,
                    index=0 if current_olt is None else olt_columns.index(current_olt) + 1,
                    key=f"manual_olt_{field}"
                )
            with col3:
                if selected_master != 'None' or selected_olt != 'None':
                    mapping_key = f"{field}_{mapping['master_col']}_{mapping['olt_col']}"
                    st.session_state.manual_mappings[mapping_key] = {
                        'master': selected_master if selected_master != 'None' else None,
                        'olt': selected_olt if selected_olt != 'None' else None
                    }
                    st.success("✅ Set")
                else:
                    st.write("Default")
            st.write("---")
    
    # Verification confirmation
    st.write("### ✅ Confirm Mappings")
    
    if mapping_verified:
        st.success("✅ All column mappings are verified!")
    else:
        st.warning("⚠️ Some mappings have issues. Please review and fix them above.")
    
    # User confirmation checkbox
    confirm_mapping = st.checkbox(
        "I have reviewed and verified all column mappings above",
        value=st.session_state.mapping_confirmed,
        key="confirm_mapping_checkbox"
    )
    
    if confirm_mapping:
        st.session_state.mapping_confirmed = True
        st.success("✅ Mapping verified! Proceeding with data processing...")
    else:
        st.info("Please verify the column mappings before proceeding.")
        # Stop here if not confirmed
        st.stop()

    # -----------------------------
    # 📊 All Entries with Duplicate Highlighting
    # -----------------------------
    st.write(f"📂 **Active Master Sheet:** `{selected_master_sheet}` | **Active OLT Sheet:** `{selected_olt_sheet}`")

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
    # 📉 Missing Records Analysis (For Information Only)
    # -----------------------------
    master_clean_df = master_df[master_df[master_plaid_col_clean].str.lower() != "nan"].copy()
    missing_mask = ~master_clean_df[master_plaid_col_clean].isin(olt_df[olt_plaid_col_clean])
    missing_records = master_clean_df[missing_mask].copy()

    st.subheader("ℹ️ Missing Records Analysis (Information Only)")
    st.write(f"Total rows in Master: **{len(master_df)}**")
    st.write(f"Total rows already in OLT: **{len(olt_df)}**")
    st.write(f"Rows that would be NEW additions: **{len(missing_records)}**")
    st.write(f"Rows that would UPDATE existing entries: **{len(master_df) - len(missing_records)}**")
    
    if len(missing_records) > 0:
        with st.expander("📋 View Missing Records (New Additions)"):
            missing_with_dup, missing_styled = highlight_duplicates(missing_records, master_plaid_col_clean)
            st.dataframe(missing_styled, use_container_width=True)

    # -----------------------------
    # 🔄 Map ALL Master Records to OLT Format
    # -----------------------------
    st.subheader("🔄 Mapping ALL Master Records to OLT Format")
    st.info(f"Mapping all {len(master_df)} records from Master to OLT format...")
    
    # Create a DataFrame to hold ALL mapped records
    all_mapped_df = pd.DataFrame(columns=olt_df.columns)
    mapped_columns_log = []
    mapping_issues = []
    
    # Process each OLT column based on the mapping
    for olt_col in olt_df.columns:
        clean_olt_name = clean_string_normalization(olt_col)
        
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
                        
                        # Comprehensive verification
                        master_category, master_confidence, _ = detect_data_category(master_df[matched_master_col])
                        olt_category, olt_confidence, _ = detect_data_category(olt_df[olt_col])
                        is_cat_match, cat_conf, cat_desc = verify_category_match(
                            master_category, master_confidence,
                            olt_category, olt_confidence
                        )
                        
                        similarity_score, is_similar, details, common_values, pattern_match = verify_data_similarity_with_patterns(
                            master_df[matched_master_col],
                            olt_df[olt_col]
                        )
                        
                        if is_cat_match and is_similar and pattern_match:
                            match_confidence = "High"
                            verification_note = f"✅ Category: {master_category} ↔ {olt_category}, Similarity: {similarity_score:.1%}, Pattern: Match"
                        elif is_cat_match and (is_similar or pattern_match):
                            match_confidence = "Medium"
                            verification_note = f"✅ Category: {master_category} ↔ {olt_category}, Similarity: {similarity_score:.1%}"
                        elif is_cat_match:
                            match_confidence = "Low"
                            verification_note = f"⚠️ Category: {master_category} ↔ {olt_category}, Data patterns differ (Sim: {similarity_score:.1%})"
                            if mapping.get('required', False):
                                mapping_issues.append(f"⚠️ Required field '{field}' has low data similarity ({similarity_score:.1%})")
                        else:
                            match_confidence = "None"
                            verification_note = f"⚠️ Category mismatch: {master_category} ↔ {olt_category}"
                            if mapping.get('required', False):
                                mapping_issues.append(f"⚠️ Required field '{field}' has category mismatch")
                        
                        break
                if matched_master_col:
                    mapped_columns_log.append({
                        'olt_col': olt_col,
                        'master_col': matched_master_col,
                        'field': field,
                        'confidence': match_confidence,
                        'verification': verification_note,
                        'description': mapping['description']
                    })
                break
        
        # If not in mapping, try to match by column name with comprehensive verification
        if not matched_master_col:
            best_match = None
            best_score = 0
            best_details = {}
            
            for clean_m_col in master_df.columns:
                # Check name similarity
                if clean_string_normalization(clean_m_col) == clean_olt_name:
                    best_match = clean_m_col
                    best_score = 1.0
                    break
                
                # Check partial match
                if clean_olt_name in clean_string_normalization(clean_m_col) or clean_string_normalization(clean_m_col) in clean_olt_name:
                    # Verify data category and similarity
                    master_category, master_confidence, _ = detect_data_category(master_df[clean_m_col])
                    olt_category, olt_confidence, _ = detect_data_category(olt_df[olt_col])
                    is_cat_match, _, _ = verify_category_match(
                        master_category, master_confidence,
                        olt_category, olt_confidence
                    )
                    
                    if is_cat_match:
                        similarity_score, is_similar, details, _, _ = verify_data_similarity_with_patterns(
                            master_df[clean_m_col],
                            olt_df[olt_col]
                        )
                        
                        if similarity_score > best_score:
                            best_score = similarity_score
                            best_match = clean_m_col
                            best_details = details
            
            if best_match and best_score > 0.4:
                matched_master_col = best_match
                match_confidence = "High" if best_score > 0.8 else "Medium" if best_score > 0.6 else "Low"
                verification_note = f"🔀 Auto-matched by category & data patterns (Score: {best_score:.1%})"
                mapped_columns_log.append({
                    'olt_col': olt_col,
                    'master_col': matched_master_col,
                    'field': 'Auto-matched',
                    'confidence': match_confidence,
                    'verification': verification_note,
                    'description': 'Auto-detected by name, category & data patterns'
                })
        
        # Handle the PLAID column specially
        if clean_olt_name == clean_string_normalization('PLAID'):
            matched_master_col = master_plaid_col_clean
            match_confidence = "High"
            verification_note = "🔑 Key identifier verified"
            mapped_columns_log.append({
                'olt_col': olt_col,
                'master_col': matched_master_col,
                'field': 'PLAID',
                'confidence': match_confidence,
                'verification': verification_note,
                'description': 'Primary Key Identifier'
            })
        
        # Apply the mapping to ALL records
        if matched_master_col:
            # Apply any formatting function if defined
            if formatting_func:
                raw_values = master_df[matched_master_col].tolist()
                all_mapped_df[olt_col] = [formatting_func(val) for val in raw_values]
            else:
                all_mapped_df[olt_col] = master_df[matched_master_col].tolist()
        else:
            all_mapped_df[olt_col] = [""] * len(master_df)
            mapped_columns_log.append({
                'olt_col': olt_col,
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
        mapping_df_display = pd.DataFrame(mapped_columns_log)
        # Color code confidence levels
        def color_confidence(val):
            if val == 'High':
                return 'background-color: #90EE90'
            elif val == 'Medium':
                return 'background-color: #FFD700'
            elif val == 'Low':
                return 'background-color: #FFA500'
            return ''
        
        # Apply styling using map
        styled_mapping = mapping_df_display.style.map(color_confidence, subset=['confidence'])
        st.dataframe(styled_mapping, use_container_width=True)
    
    # Exclude formatting column anomalies
    all_mapped_df = all_mapped_df.loc[:, ~all_mapped_df.columns.astype(str).str.contains('track', case=False)]
    
    st.subheader("📋 All Mapped Records Preview")
    st.write(f"Total mapped records: **{len(all_mapped_df)}**")
    st.dataframe(all_mapped_df.head(100), use_container_width=True)
    
    # Show preview of filled data with verification
    if len(all_mapped_df) > 0:
        st.subheader("📊 Data Population Summary")
        
        # Show which columns have data
        filled_cols = []
        for col in all_mapped_df.columns:
            non_null_count = all_mapped_df[col].notna().sum()
            if non_null_count > 0:
                filled_cols.append({
                    'Column': col,
                    'Filled Rows': non_null_count,
                    'Fill Rate': f"{non_null_count/len(all_mapped_df)*100:.1f}%"
                })
        
        if filled_cols:
            filled_df = pd.DataFrame(filled_cols)
            st.dataframe(filled_df, use_container_width=True)
        
        # Preview data
        preview_cols = [col for col in all_mapped_df.columns if all_mapped_df[col].notna().any()][:10]
        if preview_cols:
            st.dataframe(all_mapped_df[preview_cols].head(10), use_container_width=True)

    # -----------------------------
    # 💾 Download Merged File - APPEND ALL RECORDS
    # -----------------------------
    if len(all_mapped_df) > 0:
        st.subheader("💾 Append All Records to OLT")
        st.info(f"⚠️ This will **APPEND** all {len(all_mapped_df)} records from Master to the OLT file.")
        
        if st.button("🚀 Append ALL Records to OLT Spreadsheet"):
            try:
                # Reload the original OLT file to preserve original column names
                base_olt_df = olt_xls.parse(selected_olt_sheet, header=olt_header_idx)
                
                # Get the actual column names from the original file
                original_olt_columns = list(base_olt_df.columns)
                
                # Create a mapping from cleaned to original column names
                clean_to_original = {}
                for orig_col in original_olt_columns:
                    clean_col = clean_string_normalization(orig_col)
                    # Find the matching cleaned column in all_mapped_df
                    for clean_name in all_mapped_df.columns:
                        if clean_string_normalization(clean_name) == clean_string_normalization(orig_col):
                            clean_to_original[clean_name] = orig_col
                            break
                
                # Rename all_mapped_df columns to match original OLT file
                all_mapped_df_renamed = all_mapped_df.copy()
                for clean_name, orig_name in clean_to_original.items():
                    if clean_name in all_mapped_df_renamed.columns:
                        all_mapped_df_renamed.rename(columns={clean_name: orig_name}, inplace=True)
                
                # Get the PLAID column in original names
                base_plaid_col = chosen_olt_plaid_raw
                
                # Ensure the PLAID column exists in all_mapped_df_renamed
                if base_plaid_col not in all_mapped_df_renamed.columns:
                    # Try to find it by checking cleaned names
                    for col in all_mapped_df_renamed.columns:
                        if clean_string_normalization(col) == clean_string_normalization('PLAID'):
                            base_plaid_col = col
                            break
                
                # Filter out records with empty PLAID
                valid_records = all_mapped_df_renamed[all_mapped_df_renamed[base_plaid_col].astype(str).str.strip() != '']
                valid_records = valid_records[valid_records[base_plaid_col].astype(str).str.strip() != 'nan']
                
                st.info(f"📊 **Data Summary:**\n"
                       f"- Total records to append: {len(all_mapped_df)}\n"
                       f"- Valid records (with PLAID): {len(valid_records)}\n"
                       f"- Skipped invalid records: {len(all_mapped_df) - len(valid_records)}")
                
                if len(valid_records) > 0:
                    # Append ALL valid records to the OLT file
                    final_combined_df = pd.concat([base_olt_df, valid_records], ignore_index=True)
                    
                    out_buffer = io.BytesIO()
                    with pd.ExcelWriter(out_buffer, engine='openpyxl') as writer:
                        final_combined_df.to_excel(writer, sheet_name=selected_olt_sheet, index=False)
                    
                    out_buffer.seek(0)
                    
                    st.success(f"🎉 Successfully appended {len(valid_records)} records to the OLT file!")
                    st.write(f"**Total rows in new OLT file:** {len(final_combined_df)}")
                    st.download_button(
                        label="⬇️ Download Updated Nokia OLT Tracker File",
                        data=out_buffer.getvalue(),
                        file_name=f"Updated_{olt_file.name}",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.warning("⚠️ No valid records to append. All records have empty or invalid PLAID values.")
                    
            except Exception as err:
                st.error(f"Failed to process file: {err}")