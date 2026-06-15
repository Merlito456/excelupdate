import streamlit as st
import pandas as pd
import numpy as np

st.title("📊 Master Tracker → Nokia OLT Tracker Tool")

# =========================
# FILE UPLOAD
# =========================
master_file = st.file_uploader("Upload Master Tracker (Luzon)", type=["xlsx"])
olt_file = st.file_uploader("Upload Nokia OLT Tracker", type=["xlsx"])

if master_file and olt_file:

    # =========================
    # LOAD FILES
    # =========================
    master_df = pd.read_excel(master_file, sheet_name="MASTERS LIST")
    olt_df = pd.read_excel(olt_file, sheet_name="rollout")

    st.success("✅ Files loaded successfully")

    # =========================
    # COLUMN NORMALIZATION
    # =========================
    master_df.columns = master_df.columns.str.strip()
    olt_df.columns = olt_df.columns.str.strip()

    # =========================
    # REQUIRED COLUMNS
    # =========================
    master_key_cols = {
        "Site Name": "Site Name",
        "PLAID": "PLAID",
        "Region": "Region",
        "YEAR": "Build Year",
        "Electronics Equipment": "Equipment Type",
        "Number of Cards": "No. of Cards",
        "Scope Status": "Site Status",
        "Survey Date": "Site Survey Actual Date",
        "TSSR Approved Date": "TSSR Approval Actual Date",
        "Installed date": "Installation done Actual Date",
        "POWERTAPPED DATE": "Powertapping done Actual Date",
        "Integrated Date": "Integration done Actual Date",
        "PAT'ed": "PAT Done Actual Date",
        "PAC'ed": "PAC submission Actual Date",
        "FAC'ed": "FAC Submission Actual Date"
    }

    # =========================
    # SAFE COLUMN GETTER
    # =========================
    def get_col(df, col_name):
        if col_name in df.columns:
            return df[col_name]
        return np.nan

    # =========================
    # BUILD STRUCTURED DATA
    # =========================
    structured_rows = pd.DataFrame()

    for m_col, o_col in master_key_cols.items():
        structured_rows[o_col] = get_col(master_df, m_col)

    # Add missing required Nokia columns with defaults
    structured_rows["Project Tagging"] = "Auto-Generated"
    structured_rows["Clustering"] = ""
    structured_rows["Project Type"] = "OLT New Build"
    structured_rows["ORIGINAL NO. OF LINES"] = ""
    structured_rows["NO. OF LINES TO BUILD"] = ""
    structured_rows["Cabinet Location"] = ""
    structured_rows["No. of Chassis"] = ""
    structured_rows["Equipment Type (Expansion)"] = ""
    structured_rows["No. of Cards (Expansion)"] = ""
    structured_rows["Cab No."] = ""
    structured_rows["Site Status"] = get_col(master_df, "Scope Status")
    structured_rows["GO/STOP"] = ""
    structured_rows["Milestone"] = ""

    st.write("### ✅ Structured Data Preview")
    st.dataframe(structured_rows)

    # =========================
    # FIND MISSING ENTRIES
    # =========================
    master_plaid = set(master_df["PLAID"].dropna().astype(str))
    olt_plaid = set(olt_df["PLAID"].dropna().astype(str))

    missing_plaid = master_plaid - olt_plaid

    missing_df = master_df[master_df["PLAID"].astype(str).isin(missing_plaid)]

    st.write("### ❌ Missing Entries (Not in OLT Tracker)")
    st.dataframe(missing_df[["Site Name", "PLAID", "Region"]])

    # =========================
    # SAVE TO EXCEL WITH HIGHLIGHT
    # =========================
    output_file = "output_with_missing.xlsx"

    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        structured_rows.to_excel(writer, sheet_name="Structured_Data", index=False)
        missing_df.to_excel(writer, sheet_name="Missing_Entries", index=False)

        workbook = writer.book
        worksheet = writer.sheets["Missing_Entries"]

        highlight_format = workbook.add_format({'bg_color': '#FF9999'})

        for row_num in range(len(missing_df)):
            worksheet.set_row(row_num + 1, cell_format=highlight_format)

    # =========================
    # DOWNLOAD BUTTON
    # =========================
    with open(output_file, "rb") as f:
        st.download_button(
            label="📥 Download Result File (with highlight)",
            data=f,
            file_name=output_file
        )