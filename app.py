import streamlit as st
import pandas as pd
import string

st.set_page_config(page_title="OLT Tracker Tool", layout="wide")

st.title("📊 Master Tracker → Nokia OLT Rollout Tool")

# -----------------------------
# ✅ Helper Functions
# -----------------------------

def clean_columns(df):
    cleaned = []
    for col in df.columns:
        # Convert to string and handle non-breaking spaces (\xa0) and newlines
        col_str = str(col).strip().replace("\n", " ").replace("\xa0", " ")
        
        # Remove literal punctuation (like double quotes around "Number of Cards" or "PLAID")
        col_str = col_str.translate(str.maketrans('', '', string.punctuation))
        
        # Collapse multiple spaces into a single space
        col_str = " ".join(col_str.split())
        cleaned.append(col_str)
    
    df.columns = cleaned
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

    master_sheet = detect_sheet(master_xls, ["master", "luzon"])
    olt_sheet = detect_sheet(olt_xls, ["rollout", "nokia", "olt"])

    st.write(f"✅ Detected Master Sheet: `{master_sheet}`")
    st.write(f"✅ Detected Rollout Sheet: `{olt_sheet}`")

    # Read data
    master_df = master_xls.parse(master_sheet)
    olt_df = olt_xls.parse(olt_sheet)

    # Clean headers 
    master_df = clean_columns(master_df)
    olt_df = clean_columns(olt_df)

    # -----------------------------
    # 🔍 DEBUG WINDOW
    # -----------------------------
    with st.expander("👀 Click to inspect extracted columns (Debug)"):
        st.write("**Master Tracker Cleaned Columns:**", list(master_df.columns))
        st.write("**Olt Tracker Cleaned Columns:**", list(olt_df.columns))

    # -----------------------------
    # ✅ Auto Column Detection
    # -----------------------------
    master_plaid = find_column(master_df.columns, ["plaid"])
    master_site = find_column(master_df.columns, ["site name"])

    olt_plaid = find_column(olt_df.columns, ["plaid"])

    if not master_plaid or not olt_plaid:
        st.error(f"❌ Cannot detect PLAID column automatically. Found Master Plaid: '{master_plaid}', OLT Plaid: '{olt_plaid}'")
        st.info("💡 Please look at the Debug Expander above. If you see names like 'Unnamed: 0', your Excel sheet has blank rows at the very top. You can fix this by adjusting the header row index.")
        st.stop()

    st.success("✅ Main linking columns detected successfully")

    # -----------------------------
    # ✅ Normalize Keys
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
    highlight_df = master_df.copy()
    highlight_df["MISSING"] = missing_mask

    styled = highlight_df.style.apply(
        lambda row: ["background-color: #ffcccc" if row["MISSING"] else "" for _ in row],
        axis=1
    )

    st.subheader("📌 Highlighted Master List")
    st.dataframe(styled)

    # -----------------------------
    # ✅ Dynamic Mapping (Data → Rollout Structure)
    # -----------------------------
    st.subheader("🔄 Generate New Rows for Rollout")

    mapped = pd.DataFrame(index=master_df.index)

    mapped["Site Name"] = master_df[master_site] if master_site else ""
    mapped["PLAID"] = master_df[master_plaid]
    
    mapped["Region"] = master_df["Region"] if "Region" in master_df.columns else ""
    mapped["Build Year"] = master_df["Build Year"] if "Build Year" in master_df.columns else (master_df["YEAR"] if "YEAR" in master_df.columns else "")
    mapped["No. of Cards"] = master_df["Number of Cards"] if "Number of Cards" in master_df.columns else ""
    mapped["Equipment Type"] = master_df["Electronics Equipment"] if "Electronics Equipment" in master_df.columns else ""
    mapped["Site Status"] = master_df["Status"] if "Status" in master_df.columns else ""

    # Filter out only the missing rows
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