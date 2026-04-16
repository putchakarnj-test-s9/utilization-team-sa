import streamlit as st
import pandas as pd
import re

st.set_page_config(page_title="Team Utilization", layout="wide")

st.title("📊 Team Utilization Dashboard")

# ---------- CONFIG ----------
CAPACITY_HOURS = 160

# ---------- TIME PARSER ----------
def parse_time_to_hours(time_str):
    if pd.isna(time_str):
        return 0

    pattern = r'(\d+)w|(\d+)d|(\d+)h|(\d+)m'
    matches = re.findall(pattern, str(time_str))

    total_hours = 0
    for w, d, h, m in matches:
        if w:
            total_hours += int(w) * 40
        if d:
            total_hours += int(d) * 8
        if h:
            total_hours += int(h)
        if m:
            total_hours += int(m) / 60

    return total_hours


# ---------- FILE READER ----------
def load_file(file):
    filename = file.name.lower()

    if filename.endswith(".xlsx"):
        df = pd.read_excel(file)

    elif filename.endswith(".csv"):
        df = pd.read_csv(file)

    else:
        # try tab-separated first, fallback to comma
        try:
            df = pd.read_csv(file, sep="\t")
        except:
            df = pd.read_csv(file)

    return df


# ---------- TRANSFORM ----------
def transform(df):
    df.columns = [c.strip() for c in df.columns]

    required_cols = {"User", "Project", "Total"}
    if not required_cols.issubset(set(df.columns)):
        st.error("❌ File must contain columns: User, Project, Total")
        return None

    # remove TOTAL rows
    df = df[df["Project"].str.strip().str.lower() != "total"]

    # convert time
    df["Hours"] = df["Total"].apply(parse_time_to_hours)

    # aggregate
    grouped = df.groupby("User").agg({
        "Hours": "sum",
        "Project": lambda x: " · ".join(x)
    }).reset_index()

    grouped.rename(columns={"Hours": "Logged"}, inplace=True)

    # calculations
    grouped["Available"] = CAPACITY_HOURS - grouped["Logged"]
    grouped["Utilization (%)"] = (grouped["Logged"] / CAPACITY_HOURS * 100).round(0)

    def get_status(util):
        if util >= 70:
            return "On track"
        elif util >= 40:
            return "Under-used"
        return "Critical low"

    grouped["Status"] = grouped["Utilization (%)"].apply(get_status)

    return grouped


# ---------- UI ----------
uploaded_file = st.file_uploader(
    "Upload your file (Excel, CSV, TXT)",
    type=["xlsx", "csv", "txt"]
)

if uploaded_file:
    df_raw = load_file(uploaded_file)

    st.subheader("📄 Raw Data")
    st.dataframe(df_raw)

    result = transform(df_raw)

    if result is not None:
        st.subheader("📊 Utilization Summary")
        st.dataframe(result)

        # charts
        col1, col2 = st.columns(2)

        with col1:
            st.subheader("Hours per Member")
            st.bar_chart(result.set_index("User")["Logged"])

        with col2:
            st.subheader("Utilization %")
            st.bar_chart(result.set_index("User")["Utilization (%)"])
