import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="OLT Tracker Sync Tool", layout="wide")

st.title("📊 Nokia OLT Tracker Sync Tool")
st.write("Upload Master Tracker (Data File) and Nokia OLT Tracker (Rollout File)")

# Upload files
master_file = st.file_uploader("Upload Master Tracker Luzon file", type=["xlsx"])
rollout_file = st.file_uploader("Upload Nokia OLT Tracker file", type=["xlsx"])

def find_sheet(xls, possible_names):
    for name in xls.sheet_names:
        if name.strip().lower() in [p.lower() for p in possible_names]:
            return name
    return None

if master_file and rollout_file:
    try:
        # Load Excel files
        master_xls = pd.ExcelFile(master_file)
        rollout_xls = pd.ExcelFile(rollout_file)

        # Auto-detect sheets ✅
        master_sheet = find_sheet(master_xls, ["MASTER LIST"])
        rollout_sheet = find_sheet(rollout_xls, ["Rollout", "ROLL OUT", "ROLLOUT"])

        if not master_sheet:
            st.error("❌ MASTER LIST sheet not found")
            st.stop()

        if not rollout_sheet:
            st.error("❌ Rollout sheet not found")
            st.stop()

        st.success(f"✅ Detected sheets: {master_sheet} | {rollout_sheet}")

        master_df = master_xls.parse(master_sheet)
        rollout_df = rollout_xls.parse(rollout_sheet)

        # Clean columns
        master_df.columns = master_df.columns.str.strip()
        rollout_df.columns = rollout_df.columns.str.strip()

        # ✅ Use REAL columns from your files
        master_cols = {
            "site_name": "Site Name",
            "plaid": "PLAID",
            "region": "Region",
            "year": "YEAR",
            "equipment": "Electronics Equipment",
            "cards": "Number of Cards",
            "status": "Status",
            "survey": "Survey Date",
            "tssr": "TSSR Approved Date",
            "installed": "Installed date",
            "power": "POWERTAPPED DATE",
            "integration": "Integrated Date",
            "pat": "PAT'ed",
            "pac": "PAC'ed",
            "fac": "FAC'ed"
        }

        rollout_cols = {
            "site_name": "Site Name",
            "plaid": "PLAID",
            "region": "Region",
            "year": "Build Year",
            "equipment": "Equipment Type",
            "cards": "No. of Cards",
            "status": "Site Status",
            "survey": "Site Survey Actual Date",
            "tssr": "TSSR Approval Actual Date",
            "installed": "Installation done Actual Date",
            "power": "Powertapping done Actual Date",
            "integration": "Integration done Actual Date",
            "pat": "PAT Done Actual Date",
            "pac": "PAC Approval Actual Date",
            "fac": "FAC Approval Actual Date"
        }

        # ✅ Normalize PLAID for matching
        master_df["PLAID"] = master_df["PLAID"].astype(str).str.strip()
        rollout_df["PLAID"] = rollout_df["PLAID"].astype(str).str.strip()

        rollout_plaids = set(rollout_df["PLAID"])

        # ✅ Find missing entries
        missing_df = master_df[~master_df["PLAID"].isin(rollout_plaids)]

        st.subheader(f"🔍 Missing Entries Found: {len(missing_df)}")

        # ✅ Build NEW rows mapped to rollout structure
        new_rows = []

        for _, row in missing_df.iterrows():
            new_row = {}

            for key in rollout_cols:
                try:
                    new_row[rollout_cols[key]] = row[master_cols[key]]
                except:
                    new_row[rollout_cols[key]] = ""

            # Defaults
            new_row["Project Type"] = "OLT New Build"
            new_row["GO/STOP"] = "GO"

            new_rows.append(new_row)

        new_rows_df = pd.DataFrame(new_rows)

        # ✅ Combine
        updated_df = pd.concat([rollout_df, new_rows_df], ignore_index=True)

        # ✅ Highlight new rows
        def highlight_rows(row):
            if row["PLAID"] in set(missing_df["PLAID"]):
                return ["background-color: yellow"] * len(row)
            return [""] * len(row)

        styled_df = updated_df.style.apply(highlight_rows, axis=1)

        st.subheader("📋 Updated Rollout Preview")
        st.dataframe(styled_df, height=500)

        # ✅ Download file
        output = BytesIO()
        updated_df.to_excel(output, index=False)

        st.download_button(
            "⬇ Download Updated Rollout File",
            data=output.getvalue(),
            file_name="Updated_OLT_Tracker.xlsx"
        )

    except Exception as e:
        st.error(f"❌ Error: {str(e)}")