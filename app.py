import streamlit as st
import pandas as pd
import string
import io
import openpyxl
from datetime import datetime
import re
from collections import Counter

st.set_page_config(page_title="Data Merger Tool", layout="wide")

st.title("📊 Data File → Master File Merger Tool")

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

# -----------------------------
# 📋 HEADER MAPPING UI
# -----------------------------

def header_mapping_ui(data_df, master_df):
    """
    Creates an interactive UI for mapping columns from data file to master file
    Returns: dictionary mapping data columns to master columns
    """
    st.subheader("🔗 Map Columns from Data File to Master File")
    st.info("Select which columns from the Data File should be added to the Master File, and map them to the appropriate Master columns.")
    
    # Show column population for both files
    st.write("### 📊 Data File Columns")
    data_pop = analyze_column_population(data_df)
    st.dataframe(data_pop, use_container_width=True)
    
    st.write("### 📊 Master File Columns")
    master_pop = analyze_column_population(master_df)
    st.dataframe(master_pop, use_container_width=True)
    
    # Initialize session state for mapping
    if 'column_mapping' not in st.session_state:
        st.session_state.column_mapping = {}
    
    # Let user select which data columns to add
    st.write("### 📋 Select Columns to Add")
    
    # Get all data columns
    data_columns = data_df.columns.tolist()
    master_columns = master_df.columns.tolist()
    
    # Allow user to select columns to add
    selected_data_cols = st.multiselect(
        "Select columns from Data File to add to Master File:",
        options=data_columns,
        default=list(st.session_state.column_mapping.keys()) if st.session_state.column_mapping else []
    )
    
    if selected_data_cols:
        st.write("### 🔗 Map Selected Columns to Master Columns")
        
        # For each selected data column, let user choose which master column it should go to
        mapping_data = []
        for data_col in selected_data_cols:
            # Get current mapping
            current_master = st.session_state.column_mapping.get(data_col, "")
            
            # Show data preview for this column
            col1, col2, col3 = st.columns([2, 2, 1])
            with col1:
                st.write(f"**Data Column:** {data_col}")
                # Show sample values
                sample_vals = data_df[data_col].dropna().head(5).tolist()
                if sample_vals:
                    st.write(f"Sample: {', '.join(str(v) for v in sample_vals[:3])}")
                else:
                    st.write("(Empty column)")
            
            with col2:
                # Select master column
                selected_master = st.selectbox(
                    f"Map to Master column:",
                    options=['-- Skip --'] + master_columns,
                    index=0 if current_master == "" else master_columns.index(current_master) + 1,
                    key=f"map_{data_col}"
                )
                
                # Update mapping
                if selected_master != '-- Skip --':
                    st.session_state.column_mapping[data_col] = selected_master
                else:
                    if data_col in st.session_state.column_mapping:
                        del st.session_state.column_mapping[data_col]
            
            with col3:
                # Show what will happen
                if selected_master != '-- Skip --':
                    # Check if data types match
                    data_type = str(data_df[data_col].dtype)
                    master_type = str(master_df[selected_master].dtype)
                    if data_type == master_type:
                        st.success("✅ Type match")
                    else:
                        st.warning(f"⚠️ {data_type} → {master_type}")
            
            mapping_data.append({
                'Data Column': data_col,
                'Master Column': selected_master if selected_master != '-- Skip --' else 'Not mapped'
            })
        
        # Show mapping summary
        st.write("### 📋 Current Mapping Summary")
        mapping_df = pd.DataFrame(mapping_data)
        st.dataframe(mapping_df, use_container_width=True)
    
    # Button to clear all mappings
    if st.button("🔄 Clear All Mappings"):
        st.session_state.column_mapping = {}
        st.rerun()
    
    return st.session_state.column_mapping

# -----------------------------
# 📋 MERGE DATA FUNCTION
# -----------------------------

def merge_data(data_df, master_df, column_mapping, merge_key=None):
    """
    Merges data from data file into master file based on column mapping
    """
    if not column_mapping:
        st.warning("No columns mapped. Please map columns first.")
        return None
    
    # Create a copy of master file
    merged_df = master_df.copy()
    
    # Track what was added
    added_columns = []
    merged_results = []
    
    # Process each mapping
    for data_col, master_col in column_mapping.items():
        if data_col in data_df.columns and master_col in master_df.columns:
            # Check if we should merge based on a key or just add as new column
            if merge_key and merge_key in data_df.columns and merge_key in master_df.columns:
                # Merge based on key
                st.info(f"Merging '{data_col}' into '{master_col}' using '{merge_key}' as key")
                
                # Create a dictionary for fast lookup
                data_dict = {}
                for idx, row in data_df.iterrows():
                    key_val = str(row[merge_key]).strip()
                    if key_val and key_val != 'nan':
                        data_dict[key_val] = row[data_col]
                
                # Update master with data from data file
                matched_count = 0
                for idx, row in merged_df.iterrows():
                    key_val = str(row[merge_key]).strip()
                    if key_val in data_dict:
                        merged_df.loc[idx, master_col] = data_dict[key_val]
                        matched_count += 1
                
                merged_results.append({
                    'Data Column': data_col,
                    'Master Column': master_col,
                    'Rows Matched': matched_count,
                    'Total Rows': len(merged_df),
                    'Match Rate': f"{(matched_count/len(merged_df)*100):.1f}%"
                })
            else:
                # Just add as new column with suffix
                new_col_name = f"{data_col}_from_data"
                if new_col_name not in merged_df.columns:
                    merged_df[new_col_name] = data_df[data_col].values
                    added_columns.append(new_col_name)
                    merged_results.append({
                        'Data Column': data_col,
                        'Master Column': 'New Column Added',
                        'New Name': new_col_name,
                        'Rows Added': len(data_df)
                    })
    
    return merged_df, added_columns, merged_results

# -----------------------------
# ✅ File Upload Blocks
# -----------------------------

st.subheader("📂 Upload Files")

col1, col2 = st.columns(2)
with col1:
    master_file = st.file_uploader("Upload Master File (Target)", type=["xlsx"], key="master_upload")
with col2:
    data_file = st.file_uploader("Upload Data File (Source)", type=["xlsx"], key="data_upload")

if master_file and data_file:
    master_bytes = master_file.read()
    data_bytes = data_file.read()
    
    master_xls = pd.ExcelFile(io.BytesIO(master_bytes))
    data_xls = pd.ExcelFile(io.BytesIO(data_bytes))

    # -----------------------------
    # 🛠️ Sidebar Configuration & Controls
    # -----------------------------
    st.sidebar.header("🛠️ Configuration Controls")
    
    auto_master_sheet = detect_sheet(master_xls, ["master", "sheet", "list"])
    auto_data_sheet = detect_sheet(data_xls, ["data", "sheet", "list"])
    
    selected_master_sheet = st.sidebar.selectbox("Master Sheet Name", master_xls.sheet_names, index=master_xls.sheet_names.index(auto_master_sheet))
    selected_data_sheet = st.sidebar.selectbox("Data Sheet Name", data_xls.sheet_names, index=data_xls.sheet_names.index(auto_data_sheet))

    auto_master_idx = find_dynamic_header_row(master_xls, selected_master_sheet)
    auto_data_idx = find_dynamic_header_row(data_xls, selected_data_sheet)
    
    master_header_idx = st.sidebar.number_input("Master Header Row Index (1-based)", min_value=1, value=auto_master_idx + 1) - 1
    data_header_idx = st.sidebar.number_input("Data Header Row Index (1-based)", min_value=1, value=auto_data_idx + 1) - 1

    # Parse dataframes
    master_df = master_xls.parse(selected_master_sheet, header=master_header_idx)
    data_df = data_xls.parse(selected_data_sheet, header=data_header_idx)

    # Noise filtration
    master_df = master_df.loc[:, ~master_df.columns.astype(str).str.startswith('Unnamed:')]
    data_df = data_df.loc[:, ~data_df.columns.astype(str).str.startswith('Unnamed:')]
    
    # Clean columns
    master_df_cleaned = clean_columns(master_df.copy())
    data_df_cleaned = clean_columns(data_df.copy())
    
    master_df.columns = master_df_cleaned.columns
    data_df.columns = data_df_cleaned.columns

    # -----------------------------
    # 📊 FILE INFORMATION
    # -----------------------------
    st.subheader("📊 File Information")
    col1, col2 = st.columns(2)
    with col1:
        st.write(f"**Master File:** {master_file.name}")
        st.write(f"**Sheet:** {selected_master_sheet}")
        st.write(f"**Rows:** {len(master_df)}")
        st.write(f"**Columns:** {len(master_df.columns)}")
    with col2:
        st.write(f"**Data File:** {data_file.name}")
        st.write(f"**Sheet:** {selected_data_sheet}")
        st.write(f"**Rows:** {len(data_df)}")
        st.write(f"**Columns:** {len(data_df.columns)}")

    # -----------------------------
    # 📊 COLUMN MAPPING
    # -----------------------------
    st.markdown("---")
    column_mapping = header_mapping_ui(data_df, master_df)
    
    # Option to merge using a key column
    st.markdown("---")
    st.subheader("🔑 Merge Options")
    
    merge_by_key = st.checkbox("Merge by matching key column (e.g., PLAID)", value=False)
    
    merge_key = None
    if merge_by_key:
        # Find common columns
        common_cols = list(set(master_df.columns).intersection(set(data_df.columns)))
        if common_cols:
            merge_key = st.selectbox(
                "Select key column to match records:",
                options=common_cols,
                help="This column will be used to match records between files"
            )
            st.info(f"✅ Merging will match records using '{merge_key}'")
        else:
            st.warning("No common columns found. Please ensure both files have a matching identifier column (e.g., PLAID).")
    
    # Show preview of what will be added
    if column_mapping:
        st.markdown("---")
        st.subheader("📋 Preview of Data to be Added")
        
        # Show sample of data that will be added
        preview_data = {}
        for data_col, master_col in column_mapping.items():
            if data_col in data_df.columns:
                preview_data[f"{data_col} → {master_col}"] = data_df[data_col].head(5).tolist()
        
        if preview_data:
            preview_df = pd.DataFrame(preview_data)
            st.dataframe(preview_df, use_container_width=True)
        
        # Show merge preview
        if merge_by_key and merge_key:
            st.write(f"**Merge Preview (matching by '{merge_key}'):**")
            
            # Show sample of matching
            master_sample = master_df[[merge_key] + list(column_mapping.values())[:3]].head(5)
            data_sample = data_df[[merge_key] + list(column_mapping.keys())[:3]].head(5)
            
            col1, col2 = st.columns(2)
            with col1:
                st.write("**Master Sample:**")
                st.dataframe(master_sample, use_container_width=True)
            with col2:
                st.write("**Data Sample:**")
                st.dataframe(data_sample, use_container_width=True)
    
    # -----------------------------
    # 💾 MERGE AND DOWNLOAD
    # -----------------------------
    if column_mapping:
        st.markdown("---")
        st.subheader("💾 Merge and Download")
        
        if st.button("🚀 Merge Data into Master File", type="primary"):
            try:
                # Perform the merge
                result = merge_data(data_df, master_df, column_mapping, merge_key if merge_by_key else None)
                
                if result and result[0] is not None:
                    merged_df, added_columns, merge_results = result
                    
                    # Show results
                    st.success("✅ Merge completed successfully!")
                    
                    st.write("### 📊 Merge Results")
                    results_df = pd.DataFrame(merge_results)
                    st.dataframe(results_df, use_container_width=True)
                    
                    if added_columns:
                        st.info(f"Added {len(added_columns)} new columns: {', '.join(added_columns)}")
                    
                    # Show preview
                    st.write("### 📊 Merged Data Preview")
                    st.dataframe(merged_df.head(10), use_container_width=True)
                    
                    # Download button
                    out_buffer = io.BytesIO()
                    with pd.ExcelWriter(out_buffer, engine='openpyxl') as writer:
                        merged_df.to_excel(writer, sheet_name=selected_master_sheet, index=False)
                    
                    out_buffer.seek(0)
                    
                    st.download_button(
                        label="⬇️ Download Merged Master File",
                        data=out_buffer.getvalue(),
                        file_name=f"Merged_{master_file.name}",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.warning("No merge performed. Please check your mappings.")
                    
            except Exception as err:
                st.error(f"Failed to merge files: {err}")
    else:
        st.info("Please map at least one column from the Data File to the Master File.")

else:
    st.info("Please upload both files to proceed.")