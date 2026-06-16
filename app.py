import streamlit as st
import pandas as pd
import string
import io
import openpyxl
from datetime import datetime
import re
from collections import Counter

st.set_page_config(page_title="OLT Data Copy Tool", layout="wide")

st.title("📋 OLT Data Copy Tool - Copy & Paste Ready")

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
# 📊 DATA QUALITY AND POPULATION ANALYSIS
# -----------------------------

def analyze_column_population(df):
    """
    Analyzes which columns are populated and their data quality
    Returns: DataFrame with population statistics
    """
    results = []
    for col in df.columns:
        total_rows = len(df)
        non_null = df[col].count()
        null_count = df[col].isna().sum()
        empty_strings = (df[col].astype(str).str.strip() == '').sum()
        empty_strings = empty_strings - null_count  # Don't double count nulls
        
        # Calculate effective population (non-null and non-empty)
        effective_populated = non_null - empty_strings if non_null > 0 else 0
        population_pct = (effective_populated / total_rows * 100) if total_rows > 0 else 0
        
        # Determine category
        if population_pct == 0:
            status = "🔴 Empty"
        elif population_pct < 30:
            status = "🟡 Sparse"
        elif population_pct < 70:
            status = "🟠 Partial"
        else:
            status = "🟢 Well Populated"
        
        # Get data category
        category, confidence, samples = detect_data_category(df[col])
        
        results.append({
            'Column': col,
            'Total Rows': total_rows,
            'Populated': effective_populated,
            'Null': null_count,
            'Empty Strings': empty_strings,
            'Population %': f"{population_pct:.1f}%",
            'Status': status,
            'Data Category': f"{category} ({confidence:.0%})" if confidence > 0 else category,
            'Sample Values': ', '.join(str(s) for s in samples[:3]) if samples else ''
        })
    
    return pd.DataFrame(results)

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
    
    # Show data quality summary first
    st.write("#### 📊 Data Population Summary")
    pop_df = analyze_column_population(df)
    
    # Color code the status column
    def color_status(val):
        if '🟢' in str(val):
            return 'background-color: #90EE90'
        elif '🟠' in str(val):
            return 'background-color: #FFD700'
        elif '🟡' in str(val):
            return 'background-color: #FFA500'
        elif '🔴' in str(val):
            return 'background-color: #FF6B6B'
        return ''
    
    styled_pop = pop_df.style.map(color_status, subset=['Status'])
    st.dataframe(styled_pop, use_container_width=True)
    
    st.info(f"Total columns: {len(df.columns)} | Well Populated: {len(pop_df[pop_df['Status'] == '🟢 Well Populated'])} | Empty: {len(pop_df[pop_df['Status'] == '🔴 Empty'])}")
    
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
        # Get population status
        pop_status = pop_df[pop_df['Column'] == orig_col]['Status'].values[0] if orig_col in pop_df['Column'].values else 'Unknown'
        mapping_data.append({
            'Original Name': orig_col,
            'Current Name': new_col,
            'Status': '✅ Renamed' if new_col != orig_col else 'Original',
            'Population': pop_status
        })
    
    mapping_df = pd.DataFrame(mapping_data)
    styled_mapping = mapping_df.style.map(color_status, subset=['Population'])
    st.dataframe(styled_mapping, use_container_width=True)
    
    # Button to reset all renames
    if st.button(f"🔄 Reset All Column Names for {df_name}", key=f"{key_prefix}_reset"):
        st.session_state[rename_key] = {col: col for col in df.columns}
        st.success("✅ All column names reset to original")
        st.rerun()
    
    # Return the rename mapping
    return st.session_state[rename_key]

# -----------------------------
# 📋 COPY DATA INTERFACE
# -----------------------------

def copy_data_interface(mapped_data, column_mapping):
    """
    Creates an interface for copying data column by column
    """
    st.subheader("📋 Copy Data - Ready for Manual Paste")
    st.info("Select a column below to view and copy its data. Each column is formatted with headers for easy pasting into your OLT file.")
    
    # Let user select which column to copy
    available_columns = list(mapped_data.columns)
    
    # Filter to only columns that have data
    populated_columns = []
    for col in available_columns:
        non_empty = mapped_data[col].notna().sum()
        if non_empty > 0:
            populated_columns.append(col)
    
    if not populated_columns:
        st.warning("No populated columns found to copy.")
        return
    
    # Group columns by type
    st.write("### Select Column to Copy")
    
    # Create a selectbox with all populated columns
    selected_col = st.selectbox(
        "Choose a column to copy:",
        options=populated_columns,
        key="copy_column_select"
    )
    
    if selected_col:
        st.write(f"### 📊 Data for: {selected_col}")
        
        # Get the data for this column
        data_series = mapped_data[selected_col]
        
        # Show statistics
        total_rows = len(data_series)
        non_null = data_series.count()
        null_count = data_series.isna().sum()
        empty_count = (data_series.astype(str).str.strip() == '').sum() - null_count
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Total Rows", total_rows)
        with col2:
            st.metric("Populated", non_null)
        with col3:
            st.metric("Null", null_count)
        with col4:
            st.metric("Empty", max(0, empty_count))
        
        # Create a copyable text format
        st.write("### 📋 Copy Ready Format")
        st.info("Copy the text below and paste it into your OLT file. Each value is on a new line with a row number.")
        
        # Format the data for copying - option 1: Simple list
        format_type = st.radio(
            "Select format:",
            ["Simple List (one per line)", "CSV Format (comma separated)", "Tab Separated"],
            horizontal=True
        )
        
        copy_text = ""
        header_text = ""
        
        # Get the PLAID for reference if available
        plaid_col = None
        for col in mapped_data.columns:
            if 'PLAID' in col.upper() or 'plaid' in col.lower():
                plaid_col = col
                break
        
        if format_type == "Simple List (one per line)":
            # Add header
            header_text = f"Column: {selected_col}\n"
            header_text += "=" * 50 + "\n"
            if plaid_col:
                header_text += f"PLAID\t{selected_col}\n"
                header_text += "-" * 50 + "\n"
            copy_text = header_text
            
            # Add each value
            for idx, value in enumerate(data_series):
                row_num = idx + 1
                if pd.isna(value) or str(value).strip() == '' or str(value).strip() == 'nan':
                    value_str = "[EMPTY]"
                else:
                    value_str = str(value).strip()
                
                if plaid_col:
                    plaid_val = mapped_data[plaid_col].iloc[idx]
                    if pd.isna(plaid_val) or str(plaid_val).strip() == '':
                        plaid_val = "N/A"
                    copy_text += f"{plaid_val}\t{value_str}\n"
                else:
                    copy_text += f"{row_num:4d}. {value_str}\n"
            
            st.code(copy_text, language="text")
            
        elif format_type == "CSV Format (comma separated)":
            # Create CSV format with header
            header = f"Row,{selected_col}"
            if plaid_col:
                header = f"Row,PLAID,{selected_col}"
            
            rows = []
            for idx, value in enumerate(data_series):
                row_num = idx + 1
                if pd.isna(value) or str(value).strip() == '' or str(value).strip() == 'nan':
                    value_str = ""
                else:
                    value_str = str(value).strip()
                    # Escape commas if needed
                    if ',' in value_str:
                        value_str = f'"{value_str}"'
                
                if plaid_col:
                    plaid_val = mapped_data[plaid_col].iloc[idx]
                    if pd.isna(plaid_val) or str(plaid_val).strip() == '':
                        plaid_val = ""
                    else:
                        plaid_val = str(plaid_val).strip()
                    rows.append(f"{row_num},{plaid_val},{value_str}")
                else:
                    rows.append(f"{row_num},{value_str}")
            
            copy_text = header + "\n" + "\n".join(rows)
            st.code(copy_text, language="csv")
            
        else:  # Tab Separated
            header = f"Row\t{selected_col}"
            if plaid_col:
                header = f"Row\tPLAID\t{selected_col}"
            
            rows = []
            for idx, value in enumerate(data_series):
                row_num = idx + 1
                if pd.isna(value) or str(value).strip() == '' or str(value).strip() == 'nan':
                    value_str = ""
                else:
                    value_str = str(value).strip()
                    # Remove tabs if present
                    value_str = value_str.replace('\t', ' ')
                
                if plaid_col:
                    plaid_val = mapped_data[plaid_col].iloc[idx]
                    if pd.isna(plaid_val) or str(plaid_val).strip() == '':
                        plaid_val = ""
                    else:
                        plaid_val = str(plaid_val).strip()
                    rows.append(f"{row_num}\t{plaid_val}\t{value_str}")
                else:
                    rows.append(f"{row_num}\t{value_str}")
            
            copy_text = header + "\n" + "\n".join(rows)
            st.code(copy_text, language="text")
        
        # Add copy button using JavaScript
        st.write("### 📌 Copy to Clipboard")
        st.info("Select all text above (Ctrl+A or Cmd+A), then copy (Ctrl+C or Cmd+C)")
        
        # Add download option for the column data
        st.write("### 💾 Download as Text File")
        if st.button(f"Download {selected_col} data as .txt"):
            st.download_button(
                label="Click to download",
                data=copy_text,
                file_name=f"{selected_col}_data.txt",
                mime="text/plain"
            )
        
        # Also show a preview of the data in a table
        with st.expander("📊 Preview Data in Table"):
            preview_df = pd.DataFrame({
                'Row': range(1, min(len(data_series), 100) + 1),
                selected_col: data_series.head(100).tolist()
            })
            if plaid_col:
                preview_df['PLAID'] = mapped_data[plaid_col].head(100).tolist()
            st.dataframe(preview_df, use_container_width=True)

# -----------------------------
# 📊 HEADER MAPPING FOR COPY
# -----------------------------

def get_enrichment_mapping():
    """
    Defines the mapping between Master Tracker and Nokia OLT Tracker headers for data copying
    """
    
    mapping = {
        # Core Identifiers
        'PLAID': {
            'master_col': 'PLAID',
            'olt_col': 'PLAID',
            'description': 'Primary Key Identifier',
            'is_key': True
        },
        
        # Site Information
        'SITE NAME': {
            'master_col': 'SITE NAME',
            'olt_col': 'SITE NAME',
            'description': 'Site Name'
        },
        
        # Project Information
        'PROJECT TAGGING': {
            'master_col': 'PROJECT',
            'olt_col': 'PROJECT TAGGING',
            'description': 'Project Tagging'
        },
        'PROJECT TYPE': {
            'master_col': 'OLT SCOPE',
            'olt_col': 'PROJECT TYPE',
            'description': 'Project Type from OLT Scope'
        },
        
        # Location Information
        'CLUSTER': {
            'master_col': 'CLUSTER',
            'olt_col': 'CLUSTERING',
            'description': 'Clustering/Cluster'
        },
        'PROVINCE': {
            'master_col': 'PROVINCE',
            'olt_col': 'PROVINCE',
            'description': 'Province'
        },
        'REGION': {
            'master_col': 'REGION',
            'olt_col': 'REGION',
            'description': 'Region'
        },
        
        # Year Information
        'BUILD YEAR': {
            'master_col': 'BUILD YEAR',
            'olt_col': 'BUILD YEAR',
            'description': 'Build Year (formatted with "build")',
            'format': lambda x: f"{str(x).split('.')[0].strip()} build" if pd.notna(x) and str(x).strip() != "" and str(x).lower() != "nan" else ""
        },
        
        # Equipment
        'EQUIPMENT TYPE': {
            'master_col': 'ELECTRONICS EQUIPMENT',
            'olt_col': 'EQUIPMENT TYPE',
            'description': 'Equipment Type'
        },
        
        # Status Fields
        'SITE STATUS': {
            'master_col': 'SITE STATUS',
            'olt_col': 'SITE STATUS',
            'description': 'Site Status'
        },
        'SCOPE STATUS': {
            'master_col': 'SCOPE STATUS',
            'olt_col': 'SCOPE STATUS',
            'description': 'Scope Status'
        },
        
        # Milestone Dates
        'SURVEY DATE': {
            'master_col': 'SURVEY DATE',
            'olt_col': 'SITE SURVEY ACTUAL DATE',
            'description': 'Survey Date',
            'format': lambda x: str(x).split(" ")[0] if pd.notna(x) and str(x).lower() != "nan" else ""
        },
        'INSTALLED DATE': {
            'master_col': 'INSTALLED DATE',
            'olt_col': 'INSTALLATION DONE ACTUAL DATE',
            'description': 'Installation Date',
            'format': lambda x: str(x).split(" ")[0] if pd.notna(x) and str(x).lower() != "nan" else ""
        },
        'POWER TAPPED DATE': {
            'master_col': 'POWER TAPPED DATE',
            'olt_col': 'POWER TAPPING DONE ACTUAL DATE',
            'description': 'Power Tapped Date',
            'format': lambda x: str(x).split(" ")[0] if pd.notna(x) and str(x).lower() != "nan" else ""
        },
        'INTEGRATED DATE': {
            'master_col': 'INTEGRATED DATE',
            'olt_col': 'INTEGRATION DONE ACTUAL DATE',
            'description': 'Integration Date',
            'format': lambda x: str(x).split(" ")[0] if pd.notna(x) and str(x).lower() != "nan" else ""
        },
        'PAT DATE': {
            'master_col': 'PAT DATE',
            'olt_col': 'PAT DONE ACTUAL DATE',
            'description': 'PAT Date',
            'format': lambda x: str(x).split(" ")[0] if pd.notna(x) and str(x).lower() != "nan" else ""
        },
        'PAC DATE': {
            'master_col': "PAC'ED",
            'olt_col': 'PAC APPROVAL DONE ACTUAL DATE',
            'description': 'PAC Date',
            'format': lambda x: str(x).split(" ")[0] if pd.notna(x) and str(x).lower() != "nan" else ""
        },
        'FAC DATE': {
            'master_col': "FAC'ED",
            'olt_col': 'FAC APPROVAL DONE ACTUAL DATE',
            'description': 'FAC Date',
            'format': lambda x: str(x).split(" ")[0] if pd.notna(x) and str(x).lower() != "nan" else ""
        },
        
        # Additional Fields
        'HANDOVER DATE': {
            'master_col': 'HANDOVER DATE',
            'olt_col': 'HANDOVER DATE',
            'description': 'Handover Date',
            'format': lambda x: str(x).split(" ")[0] if pd.notna(x) and str(x).lower() != "nan" else ""
        },
        'TARGET DATE': {
            'master_col': 'TARGET DATE',
            'olt_col': 'TARGET DATE',
            'description': 'Target Date',
            'format': lambda x: str(x).split(" ")[0] if pd.notna(x) and str(x).lower() != "nan" else ""
        },
        'REMARKS': {
            'master_col': 'REMARKS',
            'olt_col': 'REMARKS',
            'description': 'Remarks'
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
    # 📊 DATA QUALITY ANALYSIS
    # -----------------------------
    st.subheader("📊 Data Quality & Population Analysis")
    st.info("This analysis shows which columns have data and which are empty.")
    
    # Analyze both files
    master_pop_df = analyze_column_population(master_df)
    olt_pop_df = analyze_column_population(olt_df)
    
    # Show population summary
    col1, col2 = st.columns(2)
    with col1:
        st.write("**📋 Master Tracker - Column Population**")
        st.dataframe(master_pop_df, use_container_width=True)
    with col2:
        st.write("**📋 OLT Tracker - Column Population**")
        st.dataframe(olt_pop_df, use_container_width=True)

    # -----------------------------
    # 📊 HEADER INSPECTION AND RENAMING
    # -----------------------------
    st.subheader("🔍 Header Inspection and Renaming")
    st.info("Inspect the data in each column and rename headers if needed.")
    
    # Create tabs for Master and OLT header inspection
    tab1, tab2 = st.tabs(["📋 Master Tracker Headers", "📋 OLT Tracker Headers"])
    
    with tab1:
        master_rename_mapping = header_inspection_ui(master_df, "Master Tracker", "master")
        if master_rename_mapping:
            new_master_columns = []
            for col in master_df.columns:
                new_name = master_rename_mapping.get(col, col)
                new_master_columns.append(new_name)
            master_df.columns = new_master_columns
    
    with tab2:
        olt_rename_mapping = header_inspection_ui(olt_df, "OLT Tracker", "olt")
        if olt_rename_mapping:
            new_olt_columns = []
            for col in olt_df.columns:
                new_name = olt_rename_mapping.get(col, col)
                new_olt_columns.append(new_name)
            olt_df.columns = new_olt_columns
    
    # Confirm after inspection
    st.write("### ✅ Confirm After Header Inspection")
    confirm_inspection = st.checkbox(
        "I have inspected and renamed headers as needed. Proceed to copy data.",
        key="confirm_inspection"
    )
    
    if not confirm_inspection:
        st.info("Please inspect and verify the headers before proceeding.")
        st.stop()

    # -----------------------------
    # 📊 MAP DATA FOR COPYING
    # -----------------------------
    st.subheader("🔄 Map Data for Copying")
    st.info("Mapping Master data to OLT format for easy copying.")
    
    # Get enrichment mapping
    enrichment_mapping = get_enrichment_mapping()
    
    # Display the mapping being used
    st.write("### 📋 Field Mapping")
    mapping_display = []
    for field, mapping in enrichment_mapping.items():
        mapping_display.append({
            'Field': field,
            'Master Column': mapping['master_col'],
            'OLT Column': mapping['olt_col'],
            'Description': mapping['description'],
            'Is Key': '✅' if mapping.get('is_key', False) else ''
        })
    st.dataframe(pd.DataFrame(mapping_display), use_container_width=True)
    
    # Create mapped data for copying
    # Start with Master data, but format for OLT
    mapped_data = pd.DataFrame()
    
    # Get PLAID column for matching
    master_plaid_col = master_plaid_col_clean
    olt_plaid_col = olt_plaid_col_clean
    
    # Create a lookup dictionary for OLT data
    olt_lookup = {}
    for idx, row in olt_df.iterrows():
        plaid_val = str(row[olt_plaid_col]).strip()
        if plaid_val and plaid_val != 'nan' and plaid_val != '':
            olt_lookup[plaid_val] = row
    
    # Process each field in the mapping
    for field, mapping in enrichment_mapping.items():
        master_col = mapping['master_col']
        olt_col = mapping['olt_col']
        
        # Find actual columns
        actual_master_col = None
        actual_olt_col = None
        
        for col in master_df.columns:
            if clean_string_normalization(col) == clean_string_normalization(master_col):
                actual_master_col = col
                break
        
        for col in olt_df.columns:
            if clean_string_normalization(col) == clean_string_normalization(olt_col):
                actual_olt_col = col
                break
        
        # Create data for this column
        if actual_master_col and actual_olt_col:
            # Use the OLT column name as the output column name
            output_col_name = olt_col
            
            # Prepare data
            column_data = []
            for idx, row in master_df.iterrows():
                plaid_val = str(row[master_plaid_col]).strip()
                if plaid_val in olt_lookup:
                    olt_row = olt_lookup[plaid_val]
                    olt_value = olt_row[actual_olt_col]
                    # Apply formatting if defined
                    if mapping.get('format'):
                        olt_value = mapping['format'](olt_value)
                    column_data.append(olt_value)
                else:
                    # Use Master data if no OLT match
                    column_data.append(row[actual_master_col])
            
            mapped_data[output_col_name] = column_data
        elif actual_master_col:
            # If no OLT column, use Master data
            mapped_data[olt_col] = master_df[actual_master_col]
        else:
            # If neither exists, create empty column
            mapped_data[olt_col] = [""] * len(master_df)
    
    # Show the mapped data preview
    st.subheader("📊 Mapped Data Preview")
    st.write(f"**Total rows:** {len(mapped_data)}")
    st.dataframe(mapped_data.head(20), use_container_width=True)
    
    # Show which columns were mapped
    mapped_columns = list(mapped_data.columns)
    st.info(f"✅ Mapped {len(mapped_columns)} columns ready for copying")

    # -----------------------------
    # 📋 COPY DATA INTERFACE
    # -----------------------------
    st.markdown("---")
    copy_data_interface(mapped_data, enrichment_mapping)
    
    # -----------------------------
    # 💾 Download Option
    # -----------------------------
    st.markdown("---")
    st.subheader("💾 Download All Data")
    st.info("Download all mapped data as a CSV file that can be opened in Excel.")
    
    if st.button("📥 Download All Mapped Data as CSV"):
        csv_buffer = io.StringIO()
        mapped_data.to_csv(csv_buffer, index=False)
        csv_data = csv_buffer.getvalue()
        
        st.download_button(
            label="Click to download CSV",
            data=csv_data,
            file_name="mapped_olt_data.csv",
            mime="text/csv"
        )
    
else:
    st.info("Please upload both files to proceed.")