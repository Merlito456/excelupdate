import streamlit as st
import pandas as pd
import string
import io

st.set_page_config(page_title="OLT Tracker Tool", layout="wide")

st.title("📊 Master Tracker → Nokia OLT Rollout Tool")

# -----------------------------
# ✅ Production Helper Functions
# -----------------------------

def clean_string_normalization(val) -> str:
    """Normalize a header or text to single-spaced lower-case without punctuation."""
    if pd.isna(val):
        return ""
    # Strip spaces, handle hard breaks/non-breaking spaces
    text = str(val).strip().replace("\n", " ").replace("\xa0", " ")
    # Strip punctuation
    text = text.translate(str.maketrans('', '', string.punctuation))
    # Collapse multiple internal spaces
    return " ".join(text.lower().split())

def find_dynamic_header_row(xls: pd.ExcelFile, sheet_name: str, lookahead_rows: int = 50) -> int:
    """
    Scans the beginning of an Excel sheet to find the actual header row
    containing target tracking markers like 'plaid'. Returns 0-based row index.
    """
    # Preview the upper boundary of the file without assuming a header row
    preview_df = xls.parse(sheet_name, nrows=lookahead_rows, header=None)
    
    for idx, row in preview_df.iterrows():
        row_values = [clean_string_normalization(val) for val in row.values]
        if "plaid" in row_values:
            return int(idx)
            
    return 0 # Fallback default if not discovered

def clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Normalizes DataFrame column names inline."""
    cleaned = []
    for col in df.columns:
        col_str = str(col).strip().replace("\n", " ").replace("\xa0", " ")
        col_str = col_str.translate(str.maketrans('', '', string.punctuation))
        cleaned.append(" ".join(col_str.split()))
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
# ✅ Core Business Automation Flow
# -----------------------------

master_file = st.file_uploader("Upload Master Tracker (Data File)", type=["xlsx"])
olt_file = st.file_uploader("Upload Nokia OLT Tracker (Rollout)", type=["xlsx"])

if master_file and olt_file:
    # 1. Load Excel file structures safely using context management
    master_bytes = master_file.read()
    olt_bytes = olt_file.read()
    
    master_xls = pd.ExcelFile(io.BytesIO(master_bytes))
    olt_xls = pd.ExcelFile(io.BytesIO(olt_bytes))

    master_sheet = detect_sheet(master_xls, ["master", "luzon"])
    olt_sheet = detect_sheet(olt_xls, ["rollout", "nokia", "olt"])

    st.write(f"📂 Found Master Sheet: `{master_sheet}` | Rollout Sheet: `{olt_sheet}`")

    # 2. Production Enhancement: Dynamic boundary scanning for hidden header blocks
    master_header_idx = find_dynamic_header_row(master_xls, master_sheet)
    olt_header_idx = find_dynamic_header_row(olt_xls, olt_sheet)
    
    st.info(f"⚙️ Headers dynamically located at Row {master_header_idx + 1} (Master) and Row {olt_header_idx + 1} (OLT). Skipping offset noise.")

    # 3. Parse DataFrames applying structural row offsets
    master_df = master_xls.parse(master_sheet, header=master_header_idx)
    olt_df = olt_xls.parse(olt_sheet, header=olt_header_idx)

    # 4. Standardize and clean schemas
    master_df = clean_columns(master_df)
    olt_df = clean_columns(olt_df)

    # 5. Extract Operational Columns
    master_plaid = find_column(master_df.columns, ["plaid"])
    master_site = find_column(master_df.columns, ["site name"])
    olt_plaid = find_column(olt_df.columns, ["plaid"])

    if not master_plaid or not olt_plaid:
        st.error("❌ Fatal Validation Failure: Structural matching failed.")
        with st.expander("Review Discovered Schemas"):
            st.write("Cleaned Master Elements:", list(master_df.columns))
            st.write("Cleaned OLT Elements:", list(olt_df.columns))
        st.stop()

    # 6. Vectorized Normalization (Enforces clean structural keys)
    master_df[master_plaid] = master_df[master_plaid].astype(str).str.strip()
    olt_df[olt_plaid] = olt_df[olt_plaid].astype(str).str.strip()

    # 7. Discrepancy Tracking (Identifies Missing Items)
    missing_mask = ~master_df[master_plaid].isin(olt_df[olt_plaid])
    missing_df = master_df[missing_mask].copy()

    st.subheader("❌ Missing Entries (Data → Rollout)")
    st.write(f"Total Missing Rows Isolated: **{len(missing_df)}**")
    st.dataframe(missing_df.head(100), use_container_width=True)

    # 8. Generation & Mapping Matrix
    st.subheader("🔄 Generated Data Structure Mapping")
    mapped = pd.DataFrame(index=master_df.index)

    mapped["Site Name"] = master_df[master_site] if master_site else ""
    mapped["PLAID"] = master_df[master_plaid]
    mapped["Region"] = master_df["Region"] if "Region" in master_df.columns else ""
    
    # Cascade lookup logic for build year aliases
    if "Build Year" in master_df.columns:
        mapped["Build Year"] = master_df["Build Year"]
    elif "YEAR" in master_df.columns:
        mapped["Build Year"] = master_df["YEAR"]
    else:
        mapped["Build Year"] = ""

    mapped["No. of Cards"] = master_df["Number of Cards"] if "Number of Cards" in master_df.columns else ""
    mapped["Equipment Type"] = master_df["Electronics Equipment"] if "Electronics Equipment" in master_df.columns else ""
    mapped["Site Status"] = master_df["Status"] if "Status" in master_df.columns else ""

    new_rows = mapped[missing_mask]
    st.dataframe(new_rows.head(100), use_container_width=True)

    # 9. In-Memory Workbook Packaging & Serialization
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        missing_df.to_excel(writer, sheet_name="Missing_Records", index=False)
        new_rows.to_excel(writer, sheet_name="Formatted_Upload_Rows", index=False)
    
    st.download_button(
        label="⬇️ Download Discrepancy Reports (.xlsx)",
        data=buffer.getvalue(),
        file_name="OLT_Missing_Entries.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )