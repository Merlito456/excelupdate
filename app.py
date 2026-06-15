import streamlit as st
import pandas as pd

st.title("📊 OLT Tracker Auto-Updater")

# ----------------------------
# Upload files
# ----------------------------
master_file = st.file_uploader("Upload Master Tracker File", type=["xlsx"])
olt_file = st.file_uploader("Upload Nokia OLT Tracker", type=["xlsx"])

# ----------------------------
# CLEAN COLUMN NAMES FUNCTION
# ----------------------------
def clean_columns(df):
    df.columns = (
        df.columns
        .astype(str)
        .str.replace("\n", " ", regex=True)
        .str.replace("  ", " ")
        .str.strip()
    )
    return df

# ----------------------------
# AUTO LOAD SHEET
# ----------------------------
def load_sheet(file, keyword):
    xl = pd.ExcelFile(file)
    for sheet in xl.sheet_names:
        if keyword.lower() in sheet.lower():
            return pd.read_excel(file, sheet_name=sheet)
    return None

# ----------------------------
# MAIN PROCESS
# ----------------------------
if master_file and olt_file:

    # Load sheets dynamically
    master_df = load_sheet(master_file, "master")
    olt_df = load_sheet(olt_file, "rollout")

    if master_df is None:
        st.error("❌ MASTER LIST sheet not found")
        st.stop()

    if olt_df is None:
        st.error("❌ Rollout sheet not found")
        st.stop()

    # Clean columns
    master_df = clean_columns(master_df)
    olt_df = clean_columns(olt_df)

    st.success("✅ Files loaded successfully")

    # ----------------------------
    # REQUIRED COLUMN CHECK
    # ----------------------------
    required_master_cols = ["PLAID", "Site Name", "Region", "YEAR"]
    required_olt_cols = ["PLAID", "Site Name"]

    if not all(col in master_df.columns for col in required_master_cols):
        st.error("❌ Master file missing required columns")
        st.stop()

    if not all(col in olt_df.columns for col in required_olt_cols):
        st.error("❌ OLT file missing required columns")
        st.stop()

    # ----------------------------
    # FIND MISSING ENTRIES
    # ----------------------------
    existing_plaids = olt_df["PLAID"].astype(str).str.strip()
    master_df["PLAID"] = master_df["PLAID"].astype(str).str.strip()

    missing_df = master_df[~master_df["PLAID"].isin(existing_plaids)]

    st.subheader(f"🔍 Missing Entries Found: {len(missing_df)}")

    if len(missing_df) == 0:
        st.success("✅ No missing entries. Files are aligned.")
    else:
        st.dataframe(missing_df[["PLAID", "Site Name", "Region"]])

    # ----------------------------
    # TRANSFORMATION LOGIC
    # ----------------------------
    def transform_row(row):
        return {
            "Project Tagging": "AUTO-ADDED",
            "Build Year": row.get("YEAR", ""),
            "Region": row.get("Region", ""),
            "Clustering": "",
            "Project Type": row.get("PROJECT or PROGRAM", ""),
            "Site Name": row.get("Site Name", ""),
            "PLAID": row.get("PLAID", ""),
            "ORIGINAL NO. OF LINES": "",
            "NO. OF LINES TO BUILD": "",
            "Cabinet Location": "",
            "Equipment Type": row.get("Electronics Equipment", ""),
            "No. of Chassis": "",
            "No. of Cards": row.get("Number of Cards", ""),
            "Site Status": row.get("Status", ""),
            "Target Month": "",
            "GO/STOP": "",
            "Milestone": row.get("Latest Milestone", "")
        }

    transformed_rows = pd.DataFrame([transform_row(r) for _, r in missing_df.iterrows()])

    # ----------------------------
    # APPEND TO OLT
    # ----------------------------
    updated_df = pd.concat([olt_df, transformed_rows], ignore_index=True)

    # Mark new rows
    updated_df["NEW_ENTRY"] = updated_df["PLAID"].isin(missing_df["PLAID"])

    # ----------------------------
    # HIGHLIGHT FUNCTION
    # ----------------------------
    def highlight_row(row):
        if row["NEW_ENTRY"]:
            return ["background-color: yellow"] * len(row)
        return [""] * len(row)

    st.subheader("📌 Updated OLT Tracker Preview")
    st.dataframe(updated_df.style.apply(highlight_row, axis=1))

    # ----------------------------
    # DOWNLOAD FILE
    # ----------------------------
    output_file = "updated_OLT_tracker.xlsx"
    updated_df.to_excel(output_file, index=False)

    with open(output_file, "rb") as f:
        st.download_button(
            label="📥 Download Updated File",
            data=f,
            file_name=output_file
        )