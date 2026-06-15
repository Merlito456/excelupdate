import streamlit as st
import pandas as pd
from io import BytesIO

st.title("📊 Master Tracker → Nokia OLT Sync Tool")

# =========================
# FILE UPLOAD
# =========================
master_file = st.file_uploader("Upload Master Tracker (Data File)", type=["xlsx"])
rollout_file = st.file_uploader("Upload Nokia OLT Tracker", type=["xlsx"])

if master_file and rollout_file:

    master_xls = pd.ExcelFile(master_file)
    rollout_xls = pd.ExcelFile(rollout_file)

    st.subheader("Select Sheets")

    master_sheet = st.selectbox("Select MASTER LIST sheet", master_xls.sheet_names)
    rollout_sheet = st.selectbox("Select Rollout sheet", rollout_xls.sheet_names)

    # =========================
    # LOAD DATA
    # =========================
    master_df = pd.read_excel(master_file, sheet_name=master_sheet)
    rollout_df = pd.read_excel(rollout_file, sheet_name=rollout_sheet)

    # Clean column names
    master_df.columns = master_df.columns.str.strip()
    rollout_df.columns = rollout_df.columns.str.strip()

    # =========================
    # REQUIRED COLUMNS CHECK
    # =========================
    required_master = ["Site Name", "PLAID"]
    required_rollout = ["Site Name", "PLAID"]

    if not all(col in master_df.columns for col in required_master):
        st.error("❌ Master file missing required columns")
        st.stop()

    if not all(col in rollout_df.columns for col in required_rollout):
        st.error("❌ Rollout file missing required columns")
        st.stop()

    # =========================
    # MAPPING (FULL STRUCTURE)
    # =========================
    def map_row(master_row):
        """Translate Master → Rollout structure"""

        return {
            "Project Tagging": master_row.get("PROJECT or PROGRAM"),
            "Build Year": master_row.get("YEAR"),
            "Region": master_row.get("Region"),
            "Project Type": master_row.get("OLT Scope"),
            "Site Name": master_row.get("Site Name"),
            "PLAID": master_row.get("PLAID"),
            "Equipment Type": master_row.get("Electronics Equipment"),
            "No. of Cards": master_row.get("Number of Cards"),
            "Site Status": master_row.get("Status"),
            "Site Survey Actual Date": master_row.get("Survey Date"),
            "TSSR Approval Actual Date": master_row.get("TSSR Approved Date"),
            "Installation done Actual Date": master_row.get("Installed date"),
            "Powertapping done Actual Date": master_row.get("POWERTAPPED DATE"),
            "Integration done Actual Date": master_row.get("Integrated Date"),
            "PAT Done Actual Date": master_row.get("Pre PAT Date"),
            "PAC Approval Actual Date": master_row.get("PAC'ed"),
            "FAC Approval Actual Date": master_row.get("FAC'ed"),
        }

    # =========================
    # FIND MISSING
    # =========================
    master_df["KEY"] = master_df["PLAID"].astype(str).str.strip()
    rollout_df["KEY"] = rollout_df["PLAID"].astype(str).str.strip()

    missing_rows = master_df[~master_df["KEY"].isin(rollout_df["KEY"])]

    st.success(f"✅ Missing Entries Found: {len(missing_rows)}")

    # =========================
    # ADD NEW ROWS
    # =========================
    new_entries = []

    for _, row in missing_rows.iterrows():
        new_entries.append(map_row(row))

    new_df = pd.DataFrame(new_entries)

    updated_rollout = pd.concat([rollout_df, new_df], ignore_index=True)

    # =========================
    # HIGHLIGHT FUNCTION
    # =========================
    def highlight_rows(row):
        if row["KEY"] in missing_rows["KEY"].values:
            return ["background-color: yellow"] * len(row)
        return [""] * len(row)

    styled_df = updated_rollout.style.apply(highlight_rows, axis=1)

    st.subheader("📋 Preview Updated Rollout")
    st.dataframe(styled_df)

    # =========================
    # DOWNLOAD
    # =========================
    output = BytesIO()
    updated_rollout.to_excel(output, index=False)
    output.seek(0)

    st.download_button(
        label="⬇️ Download Updated Rollout File",
        data=output,
        file_name="updated_rollout.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
