import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="OLT Tracker Sync Tool", layout="wide")

st.title("📊 OLT Tracker Sync Tool")
st.write("Compare Master Tracker vs Nokia OLT Tracker, detect missing entries, and generate rollout-ready file.")

# -----------------------------
# FILE UPLOAD
# -----------------------------
master_file = st.file_uploader("Upload MASTER Tracker File", type=["xlsx"])
olt_file = st.file_uploader("Upload Nokia OLT Tracker File", type=["xlsx"])

# -----------------------------
# FUNCTIONS
# -----------------------------
def load_excel_with_auto_sheet(file):
    xls = pd.ExcelFile(file)
    return xls.sheet_names, xls

def get_sheet(xls, selected_sheet):
    return pd.read_excel(xls, sheet_name=selected_sheet)

def normalize_columns(df):
    df.columns = df.columns.str.strip().str.replace("\n", " ").str.replace("  ", " ")
    return df

def detect_column(df, keywords):
    for col in df.columns:
        for key in keywords:
            if key.lower() in col.lower():
                return col
    return None

# -----------------------------
# MAIN PROCESS
# -----------------------------
if master_file and olt_file:

    # Load available sheets
    master_sheets, master_xls = load_excel_with_auto_sheet(master_file)
    olt_sheets, olt_xls = load_excel_with_auto_sheet(olt_file)

    st.subheader("📑 Select Sheets")
    master_sheet = st.selectbox("Master Sheet", master_sheets)
    olt_sheet = st.selectbox("OLT Sheet", olt_sheets)

    master_df = get_sheet(master_xls, master_sheet)
    olt_df = get_sheet(olt_xls, olt_sheet)

    master_df = normalize_columns(master_df)
    olt_df = normalize_columns(olt_df)

    st.success("✅ Files Loaded Successfully")

    # -----------------------------
    # AUTO DETECT KEY COLUMNS
    # -----------------------------
    master_plaid = detect_column(master_df, ["PLAID"])
    master_site = detect_column(master_df, ["Site Name"])

    olt_plaid = detect_column(olt_df, ["PLAID"])
    olt_site = detect_column(olt_df, ["Site Name"])

    st.write("### 🔍 Detected Columns")
    st.write(f"Master PLAID: {master_plaid}")
    st.write(f"Master Site: {master_site}")
    st.write(f"OLT PLAID: {olt_plaid}")
    st.write(f"OLT Site: {olt_site}")

    if not all([master_plaid, olt_plaid]):
        st.error("❌ Could not detect PLAID columns automatically.")
        st.stop()

    # -----------------------------
    # FIND MISSING ENTRIES
    # -----------------------------
    missing = master_df[~master_df[master_plaid].isin(olt_df[olt_plaid])]

    st.subheader("🚨 Missing Entries")
    st.write(f"Total Missing: {len(missing)}")
    st.dataframe(missing[[master_site, master_plaid]])

    # -----------------------------
    # MAPPING ENGINE (STRUCTURE TRANSLATION)
    # -----------------------------
    st.subheader("🔄 Mapping to OLT Structure")

    def map_to_olt(master_df):

        mapped = pd.DataFrame()

        # Core mapping based on your headers
        mapped["Project Tagging"] = master_df.get("PROJECT or PROGRAM")
        mapped["Build Year"] = master_df.get("Build Year", master_df.get("YEAR"))
        mapped["Region"] = master_df.get("Region")
        mapped["Project Type"] = master_df.get("OLT Scope")

        mapped["Site Name"] = master_df.get("Site Name")
        mapped["PLAID"] = master_df.get("PLAID")

        mapped["Equipment Type"] = master_df.get("Electronics Equipment")
        mapped["No. of Cards"] = master_df.get("Number of Cards")

        mapped["Site Status"] = master_df.get("Status")

        mapped["Site Survey Actual Date"] = master_df.get("Survey Date")
        mapped["TSSR Approval Actual Date"] = master_df.get("TSSR Approved Date")

        mapped["Installation done Actual Date"] = master_df.get("Installed date")
        mapped["Powertapping done Actual Date"] = master_df.get("POWERTAPPED DATE")
        mapped["Integration done Actual Date"] = master_df.get("Integrated Date")

        mapped["PAT Done Actual Date"] = master_df.get("PAT")
        mapped["PAC submission Actual Date"] = master_df.get("PAC")
        mapped["FAC Submission Actual Date"] = master_df.get("FAC")

        return mapped

    mapped_missing = map_to_olt(missing)

    st.write("Preview of mapped entries:")
    st.dataframe(mapped_missing)

    # -----------------------------
    # HIGHLIGHTING OUTPUT
    # -----------------------------
    def highlight_missing(df):
        return df.style.apply(lambda x: ["background-color: yellow"] * len(x), axis=1)

    st.subheader("🟡 Highlighted Missing Entries")
    st.dataframe(highlight_missing(mapped_missing))

    # -----------------------------
    # DOWNLOAD OUTPUT FILE
    # -----------------------------
    def convert_to_excel(df):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Missing Entries')
        return output.getvalue()

    excel_file = convert_to_excel(mapped_missing)

    st.download_button(
        label="📥 Download Missing Entries",
        data=excel_file,
        file_name="missing_olt_entries.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )