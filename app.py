import streamlit as st
import pandas as pd
import string
import io
import openpyxl
from datetime import datetime
import re
from collections import Counter

st.set_page_config(page_title="Data Formatter & Manual Mapping Tool", layout="wide")

st.title("📊 Data File → Master File Formatter Tool")

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

def detect_column_type(series):
    """
    Detects the appropriate type for a column based on its data
    """
    # Get non-null values
    non_null = series.dropna()
    if len(non_null) == 0:
        return 'object'
    
    # Check if all values can be converted to numeric
    try:
        pd.to_numeric(non_null)
        # Check if they are integers
        try:
            if all(isinstance(x, (int, float)) and (isinstance(x, float) and x.is_integer()) or isinstance(x, int) for x in non_null):
                return 'int64'
        except:
            pass
        return 'float64'
    except:
        pass
    
    # Check if it's datetime
    try:
        pd.to_datetime(non_null)
        return 'datetime64'
    except:
        pass
    
    # Default to object (string)
    return 'object'

def get_value_for_copy(val):
    """
    Clean a value for copy-paste - handles all data types
    """
    if pd.isna(val) or val is None:
        return ""
    
    # Handle float with .0
    if isinstance(val, float):
        if val.is_integer():
            return str(int(val))
        else:
            return str(val)
    
    # Handle other types
    return str(val).strip()

def format_value_for_master(val, master_col_type='object'):
    """
    Format a value to match the master column type
    """
    if pd.isna(val) or val is None or val == "" or val == "nan":
        return ""
    
    # For numeric columns
    if 'float' in str(master_col_type) or 'int' in str(master_col_type):
        try:
            float_val = float(val)
            if 'int' in str(master_col_type) and float_val.is_integer():
                return str(int(float_val))
            else:
                return str(float_val)
        except:
            return ""
    
    # For string/text columns
    if isinstance(val, float):
        if val.is_integer():
            return str(int(val))
        else:
            return str(val)
    
    return str(val).strip()

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
        
        # Get column type
        col_type = detect_column_type(df[col])
        
        results.append({
            'Column': col,
            'Type': col_type,
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
# 📋 MANUAL HEADER MAPPING UI
# -----------------------------

def manual_header_mapping_ui(data_df, master_df):
    """
    Creates a manual header mapping interface where users can map columns one by one
    Returns: dictionary mapping data columns to master columns
    """
    st.subheader("🔗 Manual Header Mapping")
    st.info("Manually map each column from the Data File to the Master File. This gives you full control over the mapping.")
    
    # Show both file headers side by side
    col1, col2 = st.columns(2)
    with col1:
        st.write("### 📋 Data File Headers")
        st.write(", ".join(data_df.columns.tolist()))
    with col2:
        st.write("### 📋 Master File Headers")
        st.write(", ".join(master_df.columns.tolist()))
    
    st.markdown("---")
    
    # Initialize session state for mapping
    if 'manual_mapping' not in st.session_state:
        st.session_state.manual_mapping = {}
    
    # Get all data columns and master columns
    data_columns = data_df.columns.tolist()
    master_columns = master_df.columns.tolist()
    
    # Let user select which data columns to map
    st.write("### 📋 Select Data Columns to Map")
    
    # Show all data columns with checkboxes for selection
    selected_data_cols = st.multiselect(
        "Select columns from Data File to map:",
        options=data_columns,
        default=list(st.session_state.manual_mapping.keys()) if st.session_state.manual_mapping else [],
        key="manual_select_cols"
    )
    
    if selected_data_cols:
        st.write("### 🔗 Map Each Column")
        st.info("For each selected data column, choose which master column it should map to.")
        
        # For each selected data column, let user choose which master column it should go to
        mapping_data = []
        for data_col in selected_data_cols:
            # Get current mapping
            current_master = st.session_state.manual_mapping.get(data_col, "")
            
            # Show data preview for this column
            col1, col2, col3, col4 = st.columns([2, 2, 1, 1])
            with col1:
                st.write(f"**{data_col}**")
                # Show sample values
                sample_vals = data_df[data_col].dropna().head(3).tolist()
                if sample_vals:
                    st.write(f"Sample: {', '.join(str(v) for v in sample_vals)}")
                else:
                    st.write("(Empty column)")
            
            with col2:
                # Select master column
                selected_master = st.selectbox(
                    f"Map to:",
                    options=['-- Skip --', '-- New Column --'] + master_columns,
                    index=0 if current_master == "" else (1 if current_master == "new" else master_columns.index(current_master) + 2),
                    key=f"manual_map_{data_col}"
                )
                
                # Update mapping
                if selected_master == '-- Skip --':
                    if data_col in st.session_state.manual_mapping:
                        del st.session_state.manual_mapping[data_col]
                elif selected_master == '-- New Column --':
                    st.session_state.manual_mapping[data_col] = "new"
                else:
                    st.session_state.manual_mapping[data_col] = selected_master
            
            with col3:
                # Show data type
                data_type = detect_column_type(data_df[data_col])
                st.write(f"Type: {data_type}")
            
            with col4:
                # Show status
                if selected_master == '-- New Column --':
                    st.success("✅ New")
                elif selected_master != '-- Skip --':
                    # Check if types match
                    master_type = detect_column_type(master_df[selected_master])
                    if data_type == master_type:
                        st.success("✅ Match")
                    else:
                        st.warning(f"⚠️ {data_type}→{master_type}")
                else:
                    st.write("⏭️ Skipped")
            
            mapping_data.append({
                'Data Column': data_col,
                'Master Column': selected_master if selected_master != '-- Skip --' else 'Not mapped',
                'Data Type': detect_column_type(data_df[data_col])
            })
            
            st.write("---")
        
        # Show mapping summary
        st.write("### 📋 Current Mapping Summary")
        mapping_df = pd.DataFrame(mapping_data)
        st.dataframe(mapping_df, use_container_width=True)
        
        # Show mapping statistics
        mapped_count = len([m for m in mapping_data if m['Master Column'] != 'Not mapped'])
        new_count = len([m for m in mapping_data if m['Master Column'] == 'New Column'])
        skipped_count = len([m for m in mapping_data if m['Master Column'] == 'Not mapped'])
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Total Selected", len(mapping_data))
        with col2:
            st.metric("Mapped", mapped_count)
        with col3:
            st.metric("New Columns", new_count)
        with col4:
            st.metric("Skipped", skipped_count)
    
    # Button to clear all mappings
    if st.button("🔄 Clear All Manual Mappings"):
        st.session_state.manual_mapping = {}
        st.rerun()
    
    return st.session_state.manual_mapping

# -----------------------------
# 📋 AUTO MAPPING UI
# -----------------------------

def auto_mapping_ui(data_df, master_df):
    """
    Creates an automatic mapping interface based on column name similarity
    Returns: dictionary mapping data columns to master columns
    """
    st.subheader("🤖 Auto Mapping (Name-Based)")
    st.info("Automatically map columns based on name similarity. Review and confirm the mappings.")
    
    # Initialize session state for auto mapping
    if 'auto_mapping' not in st.session_state:
        st.session_state.auto_mapping = {}
    
    # Find potential matches based on name similarity
    st.write("### 📋 Suggested Mappings")
    
    data_columns = data_df.columns.tolist()
    master_columns = master_df.columns.tolist()
    
    # For each data column, find the best matching master column
    suggestions = []
    for data_col in data_columns:
        best_match = None
        best_score = 0
        
        # Clean the data column name for comparison
        clean_data = clean_string_normalization(data_col)
        
        for master_col in master_columns:
            clean_master = clean_string_normalization(master_col)
            
            # Check exact match
            if clean_data == clean_master:
                best_match = master_col
                best_score = 1.0
                break
            
            # Check if one is contained in the other
            if clean_data in clean_master or clean_master in clean_data:
                score = 0.8
                if score > best_score:
                    best_score = score
                    best_match = master_col
            
            # Check word overlap
            data_words = set(clean_data.split())
            master_words = set(clean_master.split())
            if data_words and master_words:
                overlap = len(data_words.intersection(master_words))
                max_words = max(len(data_words), len(master_words))
                if max_words > 0:
                    score = overlap / max_words
                    if score > best_score:
                        best_score = score
                        best_match = master_col
        
        # Only suggest if score is decent
        if best_match and best_score >= 0.5:
            suggestions.append({
                'Data Column': data_col,
                'Suggested Master': best_match,
                'Confidence': f"{best_score:.0%}",
                'Action': st.selectbox(
                    f"Map '{data_col}' to:",
                    options=['-- Skip --', '-- New Column --'] + master_columns,
                    index=master_columns.index(best_match) + 2 if best_match in master_columns else 0,
                    key=f"auto_map_{data_col}"
                )
            })
    
    # Show suggestions
    if suggestions:
        suggestion_df = pd.DataFrame(suggestions)
        st.dataframe(suggestion_df, use_container_width=True)
        
        # Apply auto mappings
        if st.button("✅ Apply Auto Mappings"):
            for suggestion in suggestions:
                action = suggestion['Action']
                data_col = suggestion['Data Column']
                if action == '-- Skip --':
                    if data_col in st.session_state.auto_mapping:
                        del st.session_state.auto_mapping[data_col]
                elif action == '-- New Column --':
                    st.session_state.auto_mapping[data_col] = "new"
                else:
                    st.session_state.auto_mapping[data_col] = action
            st.success("✅ Auto mappings applied! Review and confirm below.")
            st.rerun()
    else:
        st.warning("No good matches found. Try manual mapping.")
    
    return st.session_state.auto_mapping

# -----------------------------
# 📋 HEADER MAPPING UI - Combined
# -----------------------------

def header_mapping_ui(data_df, master_df):
    """
    Creates a combined UI for both manual and auto header mapping
    Returns: dictionary mapping data columns to master columns
    """
    st.subheader("🔗 Header Mapping")
    st.info("Choose a mapping method below to map columns from the Data File to the Master File.")
    
    # Let user choose mapping method
    mapping_method = st.radio(
        "Select mapping method:",
        ["Manual Mapping", "Auto Mapping (Name-Based)"],
        horizontal=True,
        key="mapping_method"
    )
    
    if mapping_method == "Manual Mapping":
        mapping = manual_header_mapping_ui(data_df, master_df)
    else:
        mapping = auto_mapping_ui(data_df, master_df)
    
    return mapping

# -----------------------------
# 📋 FORMAT DATA FOR MASTER FUNCTION
# -----------------------------

def format_data_for_master(data_df, master_df, column_mapping):
    """
    Formats data from data file to match master file structure
    Returns: formatted DataFrame ready for copy-paste
    """
    if not column_mapping:
        return None
    
    # Create a new DataFrame with master column headers
    formatted_df = pd.DataFrame()
    
    # First, get all master columns
    master_columns = master_df.columns.tolist()
    
    # For each master column, find matching data column or leave empty
    for master_col in master_columns:
        # Find if this master column is mapped from any data column
        mapped_data_col = None
        for data_col, map_to in column_mapping.items():
            if map_to == master_col:
                mapped_data_col = data_col
                break
        
        if mapped_data_col and mapped_data_col in data_df.columns:
            # Format the data to match master column type
            master_type = detect_column_type(master_df[master_col])
            formatted_values = []
            for val in data_df[mapped_data_col]:
                formatted_val = format_value_for_master(val, master_type)
                formatted_values.append(formatted_val)
            formatted_df[master_col] = formatted_values
        else:
            # If no mapping, leave column empty
            formatted_df[master_col] = [""] * len(data_df)
    
    # Add any new columns that were mapped to "New Column"
    for data_col, map_to in column_mapping.items():
        if map_to == "new":
            new_col_name = f"{data_col}_from_data"
            if new_col_name not in formatted_df.columns:
                formatted_values = []
                for val in data_df[data_col]:
                    formatted_val = get_value_for_copy(val)
                    formatted_values.append(formatted_val)
                formatted_df[new_col_name] = formatted_values
    
    return formatted_df

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

    # -----------------------------
    # 📋 FORMAT AND COPY DATA
    # -----------------------------
    if column_mapping:
        st.markdown("---")
        st.subheader("📋 Formatted Data Ready for Master File")
        st.info("This data is formatted to match the Master File structure. Copy and paste it directly into Excel.")
        
        # Format the data
        formatted_df = format_data_for_master(data_df, master_df, column_mapping)
        
        if formatted_df is not None:
            # Show preview
            st.write("### 📊 Formatted Data Preview")
            st.write(f"**Total rows:** {len(formatted_df)}")
            st.write(f"**Total columns:** {len(formatted_df.columns)}")
            st.dataframe(formatted_df.head(10), use_container_width=True)
            
            # ---------- COPY-PASTE SECTION ----------
            st.write("### 📋 Copy-Paste Ready Data")
            
            # Let user select which column to view
            available_cols = formatted_df.columns.tolist()
            selected_col = st.selectbox(
                "Select column to copy:",
                options=available_cols,
                key="copy_column_select"
            )
            
            if selected_col:
                # Get the data for this column
                col_data = formatted_df[selected_col]
                
                # Count non-empty values
                non_empty = 0
                for val in col_data:
                    if val is not None and str(val).strip() != '':
                        non_empty += 1
                
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Total Rows", len(col_data))
                with col2:
                    st.metric("Populated", non_empty)
                with col3:
                    st.metric("Empty", len(col_data) - non_empty)
                
                # Create clean copy text - ONLY THE VALUES
                st.write("### 📋 Copy This Data")
                st.info("Select all text below (Ctrl+A), then copy (Ctrl+C), and paste into Excel.")
                
                copy_text = ""
                for val in col_data:
                    if pd.isna(val) or val is None:
                        copy_text += "\n"
                    else:
                        copy_text += f"{str(val)}\n"
                
                # Show the text in a code block
                st.code(copy_text, language="text")
                
                # Copy instructions
                st.write("### 📌 How to Copy")
                st.info("""
                1. **Select all text** in the box above (Ctrl+A or Cmd+A)
                2. **Copy** the text (Ctrl+C or Cmd+C)
                3. **Go to Excel** and select the cell where you want to paste
                4. **Paste** (Ctrl+V or Cmd+V)
                """)
                
                # Download individual column
                st.write("### 💾 Download as Text File")
                if st.button(f"Download {selected_col} data"):
                    st.download_button(
                        label="Click to download .txt file",
                        data=copy_text,
                        file_name=f"{selected_col}_formatted.txt",
                        mime="text/plain"
                    )
            
            # ---------- DOWNLOAD ALL DATA ----------
            st.markdown("---")
            st.write("### 💾 Download All Formatted Data")
            
            col1, col2 = st.columns(2)
            
            with col1:
                # Download as CSV
                if st.button("📥 Download as CSV"):
                    csv_buffer = io.StringIO()
                    formatted_df.to_csv(csv_buffer, index=False)
                    csv_data = csv_buffer.getvalue()
                    
                    st.download_button(
                        label="Click to download CSV",
                        data=csv_data,
                        file_name="formatted_data_for_master.csv",
                        mime="text/csv"
                    )
            
            with col2:
                # Download as Excel
                if st.button("📥 Download as Excel"):
                    excel_buffer = io.BytesIO()
                    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                        formatted_df.to_excel(writer, sheet_name="Formatted_Data", index=False)
                    
                    excel_buffer.seek(0)
                    
                    st.download_button(
                        label="Click to download Excel",
                        data=excel_buffer.getvalue(),
                        file_name="formatted_data_for_master.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            
            # ---------- PREVIEW ALL DATA ----------
            with st.expander("📊 View All Formatted Data"):
                st.dataframe(formatted_df, use_container_width=True)
            
            # Show column mapping summary
            with st.expander("📋 Column Mapping Summary"):
                mapping_summary = []
                for data_col, map_to in column_mapping.items():
                    if map_to == "new":
                        mapping_summary.append({
                            'Data Column': data_col,
                            'Maps To': 'New Column',
                            'Data Type': detect_column_type(data_df[data_col]),
                            'Rows': len(data_df)
                        })
                    else:
                        mapping_summary.append({
                            'Data Column': data_col,
                            'Maps To': map_to,
                            'Data Type': detect_column_type(data_df[data_col]),
                            'Master Type': detect_column_type(master_df[map_to]) if map_to in master_df.columns else 'N/A',
                            'Rows': len(data_df)
                        })
                
                st.dataframe(pd.DataFrame(mapping_summary), use_container_width=True)

else:
    st.info("Please upload both files to proceed.")