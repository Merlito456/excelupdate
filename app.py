import streamlit as st
import pandas as pd

st.set_page_config(page_title="OLT Tracker Tool", layout="wide")

st.title("📊 Master Tracker → Nokia OLT Rollout Tool")

# -----------------------------
# ✅ Helper Functions
# -----------------------------

def clean_columns(df):
    df.columns = [
        str(col).strip().replace("\n", " ").replace("  ", " ")
        for col in df.columns
    ]
    return df

def find_column(columns, keywords):
    for col in columns:
        col_str = str(col).lower()
        for key in keywords:
            if key.lower() in col_str:
                return col
    return None

def detect_sheet(xls, keywords):
    for sheet in xls.sheet_names:
        name = sheet.lower()
        if any(k in name for k in keywords):
            return sheet
    return xls.sheet_names[0]  # fallback

# -----------------------------
# ✅ Upload Files
# -----------------------------

master_file = st.file_uploader(
    "Upload Master Tracker (Data File)",
    type=["xlsx"]
)

olt_file = st.file_uploader(
    "Upload Nokia OLT Tracker (Rollout)",
    type=["xlsx"]
)

if master_file and olt_file:

    # -----------------------------
    # ✅ Load Excel with auto sheet detection
    # -----------------------------
    master_xls = pd.ExcelFile(master_file)
    olt_xls = pd.ExcelFile(olt_file)

    master_sheet = detect_sheet(master_xls, ["master"])
    olt_sheet = detect_sheet(olt_xls, ["rollout"])

    st.write(f"✅ Detected Master Sheet: `{master_sheet}`")
    st.write(f"✅ Detected Rollout Sheet: `{olt_sheet}`")

    # read
    master_df = master_xls.parse(master_sheet)
    olt_df = olt_xls.parse(olt_sheet)

    # clean headers
    master_df = clean_columns(master_df)
    olt_df = clean_columns(olt_df)

    # -----------------------------
    # ✅ Auto Column Detection
    # -----------------------------
    master_plaid = find_column(master_df.columns, ["plaid"])
    master_site = find_column(master_df.columns, ["site name"])

    olt_plaid = find_column(olt_df.columns, ["plaid"])
    olt_site = find_column(olt_df.columns, ["site name"])

    olt_region = find_column(olt_df.columns, ["region"])
    olt_cards = find_column(olt_df.columns, ["cards"])

    if not master_plaid or not olt_plaid:
        st.error("❌ Cannot detect PLAID column automatically.")
        st.stop()

    st.success("✅ Columns detected successfully")

    # -----------------------------
    # ✅ Normalize keys
    # -----------------------------
    master_df[master_plaid] = master_df[master_plaid].astype(str).str.strip()
    olt_df[olt_plaid] = olt_df[olt_plaid].astype(str).str.strip()

    # -----------------------------
    # ✅ Missing Entry Detection
    # -----------------------------
    missing_mask = ~master_df[master_plaid].isin(olt_df[olt_plaid])
    missing_df = master_df[missing_mask].copy()

    st.subheader("❌ Missing Entries (Data → Rollout)")
    st.write(f"Total Missing: {len(missing_df)}")

    st.dataframe(missing_df.head(50))

    # -----------------------------
    # ✅ Highlight Output
    # -----------------------------
    def highlight_missing(row):
        color = "background-color: red"
        return [color if x else "" for x in row]

    highlight_df = master_df.copy()
    highlight_df["MISSING"] = missing_mask

    styled = highlight_df.style.apply(
        lambda row: ["background-color: #ffcccc" if row["MISSING"] else "" for _ in row],
        axis=1
    )

    st.subheader("📌 Highlighted Master List")
    st.dataframe(styled)

    # -----------------------------
    # ✅ Mapping (Data → Rollout Structure)
    # -----------------------------
    st.subheader("🔄 Generate New Rows for Rollout")

    mapped = pd.DataFrame()

    mapped["Site Name"] = master_df[master_site] if master_site else ""
    mapped["PLAID"] = master_df[master_plaid]
    mapped["Region"] = master_df.get("Region", "")
    mapped["Build Year"] = master_df.get("YEAR", "")

    # Optional fields
    mapped["No. of Cards"] = master_df.get("Number of Cards", "")
    mapped["Equipment Type"] = master_df.get("Electronics Equipment", "")
    mapped["Site Status"] = master_df.get("Status", "")

    # only missing
    new_rows = mapped[missing_mask]

    st.write(f"✅ New Rows Ready: {len(new_rows)}")
    st.dataframe(new_rows.head(50))

    # -----------------------------
    # ✅ Download Button
    # -----------------------------
    output_file = "OLT_Missing_Entries.xlsx"

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        missing_df.to_excel(writer, sheet_name="Missing", index=False)
        new_rows.to_excel(writer, sheet_name="Formatted", index=False)

    with open(output_file, "rb") as f:
        st.download_button(
            "⬇ Download Result",
            f,
            output_file
        )