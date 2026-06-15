import streamlit as st
import pandas as pd
from io import BytesIO

st.title("OLT Tracker Sync Tool")

# ======================
# FILE UPLOAD
# ======================
master_file = st.file_uploader("Upload Master Tracker (Luzon)", type=["xlsx"])
olt_file = st.file_uploader("Upload Nokia OLT Tracker", type=["xlsx"])

# ======================
# AUTO SHEET DETECTION
# ======================
def load_sheet(file, possible_names):
    xl = pd.ExcelFile(file)
    for sheet in xl.sheet_names:
        for name in possible_names:
            if name.lower() in sheet.lower():
                return xl.parse(sheet)
    st.error(f"Sheet not found. Available sheets: {xl.sheet_names}")
    st.stop()

# ======================
# PROCESS
# ======================
if master_file and olt_file:

    # ✅ Flexible sheet names
    master_df = load_sheet(master_file, ["MASTER LIST"])
    olt_df = load_sheet(olt_file, ["rollout"])

    st.success("Files loaded successfully ✅")

    # ======================
    # CLEAN COLUMN NAMES
    # ======================
    master_df.columns = master_df.columns.str.strip()
    olt_df.columns = olt_df.columns.str.strip()

    # ======================
    # KEY COLUMN CHECK
    # ======================
    if "PLAID" not in master_df.columns or "PLAID" not in olt_df.columns:
        st.error("PLAID column missing in one of the files")
        st.stop()

    # ======================
    # FIND MISSING ENTRIES
    # ======================
    master_plaids = set(master_df["PLAID"].astype(str))
    olt_plaids = set(olt_df["PLAID"].astype(str))

    missing_plaids = master_plaids - olt_plaids

    missing_df = master_df[master_df["PLAID"].astype(str).isin(missing_plaids)].copy()

    st.write(f"### Missing Entries Count: {len(missing_df)}")

    # ======================
    # MAPPING ENGINE
    # ======================
    def map_row(row):
        return {
            "Project Tagging": "AUTO-ADD",
            "Build Year": row.get("YEAR"),
            "Region": row.get("Region"),
            "Project Type": row.get("PROJECT or PROGRAM"),
            "Site Name": row.get("Site Name"),
            "PLAID": row.get("PLAID"),
            "Equipment Type": row.get("Electronics Equipment"),
            "No. of Cards": row.get("Number of Cards"),
            "Site Status": row.get("Status"),
            "Milestone": row.get("Latest Milestone"),
            "Installation done Actual Date": row.get("Installed date"),
            "Integration done Actual Date": row.get("Integrated Date"),
            "PAC Approval Actual Date": row.get("PAC'ed"),
            "FAC Approval Actual Date": row.get("FAC'ed"),
        }

    mapped_rows = missing_df.apply(map_row, axis=1)
    mapped_df = pd.DataFrame(mapped_rows.tolist())

    # ======================
    # ADD TO OLT DATA
    # ======================
    updated_df = pd.concat([olt_df, mapped_df], ignore_index=True)

    # ======================
    # HIGHLIGHT FLAG
    # ======================
    updated_df["NEW ENTRY"] = updated_df["PLAID"].isin(missing_plaids)

    def highlight(row):
        if row["NEW ENTRY"]:
            return ["background-color: yellow"] * len(row)
        return [""] * len(row)

    st.write("### Updated Rollout (Preview)")
    st.dataframe(updated_df.style.apply(highlight, axis=1))

    # ======================
    # DOWNLOAD
    # ======================
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        updated_df.to_excel(writer, index=False, sheet_name="Updated Rollout")

    st.download_button(
        "Download Updated File",
        data=output.getvalue(),
        file_name="updated_rollout.xlsx"
    )
``