import streamlit as st
import pandas as pd
import re
import matplotlib.pyplot as plt
from matplotlib.patches import Patch

# ---------- PAGE CONFIG ----------
st.set_page_config(page_title="Team Utilization Dashboard", layout="wide")

st.title("📊 Team Utilization Dashboard")

# ---------- CONFIG ----------
CAPACITY_HOURS = 160  # 1 month capacity

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
    name = file.name.lower()

    if name.endswith(".xlsx"):
        return pd.read_excel(file)
    elif name.endswith(".csv"):
        return pd.read_csv(file)
    else:
        return pd.read_csv(file, sep="\t")


# ---------- TRANSFORM ----------
def transform(df):
    df.columns = [c.strip() for c in df.columns]

    required = {"User", "Project", "Total"}
    if not required.issubset(set(df.columns)):
        st.error("❌ File must contain columns: User, Project, Total")
        return None

    # remove TOTAL rows
    df = df[df["Project"].str.strip().str.lower() != "total"]

    # convert to hours
    df["Hours"] = df["Total"].apply(parse_time_to_hours)

    return df


# ---------- UPLOAD ----------
uploaded_file = st.file_uploader(
    "Upload file (Excel / CSV / TXT)",
    type=["xlsx", "csv", "txt"]
)

if uploaded_file:
    df_raw = load_file(uploaded_file)

    st.subheader("📄 Raw Data")
    st.dataframe(df_raw)

    df = transform(df_raw)

    if df is not None:

        st.subheader("📄 Clean Data (Converted to Hours)")
        st.dataframe(df)

        # ---------- USER SELECT ----------
        users = sorted(df["User"].unique())
        selected_user = st.selectbox("👤 Select User", users)

        user_df = df[df["User"] == selected_user]

        # ---------- PROJECT SUMMARY ----------
        project_summary = (
            user_df.groupby("Project")["Hours"]
            .sum()
            .reset_index()
        )

        # ---------- S9 FLAG ----------
        project_summary["is_s9"] = project_summary["Project"].str.lower().str.contains("s9")

        # sort: S9 first, then hours desc
        project_summary = project_summary.sort_values(
            by=["is_s9", "Hours"], ascending=[False, False]
        )

        st.subheader(f"📊 {selected_user} - Hours per Project")
        st.dataframe(project_summary.drop(columns=["is_s9"]))

        # ---------- COLOR LOGIC ----------
        colors = [
            "#ef4444" if is_s9 else "#94a3b8"
            for is_s9 in project_summary["is_s9"]
        ]

        # ---------- BAR CHART ----------
        fig, ax = plt.subplots()

        ax.barh(
            project_summary["Project"],
            project_summary["Hours"],
            color=colors
        )

        ax.set_xlabel("Hours")
        ax.set_ylabel("Project")
        ax.set_title("Hours per Project (S9 Highlighted)")

        # legend
        legend_elements = [
            Patch(facecolor="#ef4444", label="S9 Project"),
            Patch(facecolor="#94a3b8", label="Other Projects"),
        ]
        ax.legend(handles=legend_elements)

        st.pyplot(fig)

        # ---------- PIE CHART ----------
        st.subheader("📌 Project Distribution")

        fig2, ax2 = plt.subplots()

        ax2.pie(
            project_summary["Hours"],
            labels=project_summary["Project"],
            autopct="%1.1f%%"
        )

        ax2.set_title("Project Distribution")

        st.pyplot(fig2)

        # ---------- UTILIZATION SUMMARY ----------
        total_logged = user_df["Hours"].sum()
        utilization = round((total_logged / CAPACITY_HOURS) * 100, 0)

        if utilization >= 70:
            status = "On track"
        elif utilization >= 40:
            status = "Under-used"
        else:
            status = "Critical low"

        st.subheader("📈 Utilization Summary")

        col1, col2, col3 = st.columns(3)

        col1.metric("Total Hours", f"{total_logged:.1f} h")
        col2.metric("Utilization", f"{utilization}%")
        col3.metric("Status", status)
