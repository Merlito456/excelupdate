import streamlit as st
import pandas as pd

st.set_page_config(page_title="OLT Tracker Sync Tool", layout="wide")

st.title("📊 OLT Tracker Auto Sync + Missing Checker")

# -----------------------------
# Helper: find column by keywords
# -----------------------------
def find_column(columns, keywords):
    for col in columns:
        for key in keywords:
            if key.lower() in col.lower():
                return col
    return None

# -----------------------------
# Helper: auto sheet loader
# -----------------------------
def load_excel_smart(file, possible_names):
    xls = pd.ExcelFile(file)
    for name in xls.sheet_names:
        for expected in possible_names:
            if expected.lower() in name.lower():
                return pd.read_excel(file, sheet_name=name)
    return None


# -----------------------------
# Upload UI
# -----------------------------
master_file = st.file_uploader("Upload Master Tracker (Data File)", type=["xlsx"])
olt_file = st.file_uploader("Upload Nokia OLT Tracker", type=["xlsx"])


if master_file and olt_file:

    # -----------------------------
    # Load sheets (AUTO DETECT)
    # -----------------------------
    master_df = load_excel_smart(master_file, ["master list"])
    olt_df = load_excel_smart(olt_file, ["rollout"])

    if master_df is None:
        st.error("❌ MASTER LIST sheet not found")
        st.stop()

    if olt_df is None:
        st.error("❌ Rollout sheet not found")
        st.stop()

    # -----------------------------
    # Normalize column names
    # -----------------------------
    master_df.columns = master_df.columns.str.strip()
    olt_df.columns = olt_df.columns.str.strip()

    # -----------------------------
    # Detect columns dynamically
    # -----------------------------
    master_plaid = find_column(master_df.columns, ["plaid"])
    master_site = find_column(master_df.columns, ["site name"])
    master_region = find_column(master_df.columns, ["region"])
    master_year = find_column(master_df.columns, ["year"])
    master_cards = find_column(master_df.columns, ["number of cards"])

    olt_plaid = find_column(olt_df.columns, ["plaid"])
    olt_site = find_column(olt_df.columns, ["site name"])
    olt_region = find_column(olt_df.columns, ["region"])
    olt_year = find_column(olt_df.columns, ["build year"])
    olt_cards = find_column(olt_df.columns, ["no. of cards"])

    if not master_plaid or not olt_plaid:
        st.error("❌ PLAID column not detected")
        st.stop()

    st.success("✅ Columns detected successfully")

    # -----------------------------
    # Extract lists
    # -----------------------------
    master_list = master_df[master_plaid].astype(str).str.strip()
    olt_list = olt_df[olt_plaid].astype(str).str.strip()

    # -----------------------------
    # Find missing
    # -----------------------------
    missing_mask = ~master_list.isin(olt_list)
    missing_df = master_df[missing_mask]

    st.subheader("🔍 Missing Entries")
    st.write(f"Total Missing: {missing_df.shape[0]}")
    st.dataframe(missing_df)

    # -----------------------------
    # ROW BUILDER (STRUCTURE TRANSLATION)
    # -----------------------------
    st.subheader("🔄 Converted Rows (Ready for Rollout)")

    new_rows = pd.DataFrame({
        "Project Tagging": "AUTO-ADDED",
        "Build Year": master_df[master_year] if master_year else "",
        "Region": master_df[master_region] if master_region else "",
        "Site Name": master_df[master_site] if master_site else "",
        "PLAID": master_df[master_plaid],
        "No. of Cards": master_df[master_cards] if master_cards else "",
        "Site Status": "FOR VALIDATION"
    })

    new_rows = new_rows.loc[missing_mask]

    st.dataframe(new_rows)

    # -----------------------------
    # Highlight Excel Output
    # -----------------------------
    def highlight_missing(row):
        return ['background-color: yellow']*len(row)

    styled_missing = missing_df.style.apply(highlight_missing, axis=1)

    output_file = "OLT_Missing_Output.xlsx"

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        missing_df.to_excel(writer, sheet_name="Missing", index=False)
        new_rows.to_excel(writer, sheet_name="Ready_to_Add", index=False)

    with open(output_file, "rb") as f:
        st.download_button("⬇ Download Result File", f, file_name=output_file)