import streamlit as st
import pandas as pd
import re
import matplotlib.pyplot as plt
from matplotlib.patches import Patch

# ---------- PAGE CONFIG ----------
st.set_page_config(page_title="Team Utilization Dashboard", layout="wide")

st.title("📊 Team Utilization Dashboard")

# ---------- CONFIG ----------
CAPACITY_HOURS = 160  # total hours in month

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

    df = df[df["Project"].str.strip().str.lower() != "total"]
    df["Hours"] = df["Total"].apply(parse_time_to_hours)

    return df

# ---------- UPLOAD ----------
uploaded_file = st.file_uploader(
    "Upload file (Excel / CSV / TXT)",
    type=["xlsx", "csv", "txt"]
)

if uploaded_file:
    df_raw = load_file(uploaded_file)
    df = transform(df_raw)

    if df is not None:

        # ---------- USER SELECT ----------
        users = sorted(df["User"].unique())
        selected_user = st.selectbox("👤 Select User", users)

        user_df = df[df["User"] == selected_user]

        # ---------- PROJECT SUMMARY ----------
        project_summary = (
            user_df.groupby("Project")["Hours"]
            .sum()
            .reset_index()
            .sort_values(by="Project")
        )

        st.subheader(f"📊 {selected_user} - Hours per Project")
        st.dataframe(project_summary, hide_index=True)

        # ---------- COLOR LOGIC ----------
        def get_color(project):
            p = project.lower()
            if "s9 - work order" in p:
                return "#ef4444"  # red
            elif "s9 - tech support" in p:
                return "#3b82f6"  # blue
            return "#22c55e"      # green

        colors = [get_color(p) for p in project_summary["Project"]]

        # ---------- TOTAL HOURS ----------
        total_logged = user_df["Hours"].sum()

        # ---------- BAR CHART ----------
        fig, ax = plt.subplots()

        ax.barh(
            project_summary["Project"],
            project_summary["Hours"],
            color=colors
        )

        # set max range to total hours of month
        ax.set_xlim(0, max(total_logged, CAPACITY_HOURS))

        ax.set_xlabel("Hours")
        ax.set_ylabel("Project")
        ax.set_title("Hours per Project (Max = Total Monthly Hours)")

        legend_elements = [
            Patch(facecolor="#ef4444", label="S9 - Work Order"),
            Patch(facecolor="#3b82f6", label="S9 - Tech Support"),
            Patch(facecolor="#22c55e", label="Other Projects"),
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

        # ---------- UTILIZATION ----------
        non_s9_work_order_hours = user_df[
            ~user_df["Project"].str.lower().str.contains("s9 - work order")
        ]["Hours"].sum()

        utilization = round((non_s9_work_order_hours / total_logged) * 100, 0) if total_logged > 0 else 0

        st.subheader("📈 Utilization Summary")

        col1, col2, col3 = st.columns(3)

        col1.metric("Total Hours (Month)", f"{total_logged:.1f} h")
        col2.metric("Capacity", f"{CAPACITY_HOURS} h")
        col3.metric("Utilization (Excl. S9 Work Order)", f"{utilization}%")
