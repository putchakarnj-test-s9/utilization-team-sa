import streamlit as st
import pandas as pd
import re
import matplotlib.pyplot as plt
from matplotlib.patches import Patch
from datetime import datetime

# ---------- PAGE CONFIG ----------
st.set_page_config(page_title="Team Utilization Dashboard", layout="wide")

# ---------- MONTH DISPLAY (FROM FILE NAME) ----------
def extract_month_from_filename(filename):
    match = re.search(r"(\d{2})\.(\d{2})\.(\d{4})_(\d{2})\.(\d{2})\.(\d{4})", filename)
    
    if match:
        day, month, year = match.group(1), match.group(2), match.group(3)
        try:
            date_obj = datetime.strptime(f"{day}.{month}.{year}", "%d.%m.%Y")
            return date_obj.strftime("%B %Y")
        except:
            return None
    return None

# Default title
st.title("📊 Team Utilization Dashboard")

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

    # Exclude unwanted rows
    df = df[df["Project"].str.strip().str.lower() != "total"]
    df = df[df["User"].str.strip().str.lower() != "summary"]

    df["Hours"] = df["Total"].apply(parse_time_to_hours)

    return df

# ---------- UPLOAD ----------
uploaded_file = st.file_uploader(
    "Upload file (Excel / CSV / TXT)",
    type=["xlsx", "csv", "txt"]
)

if uploaded_file:

    # ---------- SET TITLE FROM FILE NAME ----------
    month_label = extract_month_from_filename(uploaded_file.name)
    if month_label:
        st.title(f"📊 Team Utilization Dashboard - {month_label}")
    else:
        st.title("📊 Team Utilization Dashboard")

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

        total_logged = project_summary["Hours"].sum()

        # Add TOTAL row
        total_row = pd.DataFrame({
            "Project": ["TOTAL"],
            "Hours": [total_logged]
        })

        display_table = pd.concat([project_summary, total_row], ignore_index=True)

        st.subheader(f"📊 {selected_user} - Hours per Project")
        st.dataframe(display_table, hide_index=True, use_container_width=True)

        # ---------- COLOR LOGIC ----------
        def get_color(project):
            p = project.lower()
            if "s9 - work order" in p:
                return "#ef4444"  # red
            elif "s9 - tech support" in p:
                return "#3b82f6"  # blue
            return "#22c55e"      # green

        colors = [get_color(p) for p in project_summary["Project"]]

        # ---------- BAR CHART (CENTERED) ----------
        col1, col2, col3 = st.columns([1, 3, 1])

        with col2:
            fig, ax = plt.subplots(figsize=(10, 5))

            ax.barh(
                project_summary["Project"],
                project_summary["Hours"],
                color=colors
            )

            ax.set_xlim(0, total_logged)

            ax.set_xlabel("Hours")
            ax.set_ylabel("Project")
            ax.set_title("Hours per Project (Max = Total Logged Hours)")

            legend_elements = [
                Patch(facecolor="#ef4444", label="S9 - Work Order"),
                Patch(facecolor="#3b82f6", label="S9 - Tech Support"),
                Patch(facecolor="#22c55e", label="Other Projects"),
            ]
            ax.legend(handles=legend_elements)

            st.pyplot(fig)

        # ---------- PIE CHART ----------
        st.subheader("📌 Project Distribution")

        col4, col5, col6 = st.columns([1, 2, 1])

        with col5:
            fig2, ax2 = plt.subplots(figsize=(5, 5))

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

        utilization = round(
            (non_s9_work_order_hours / total_logged) * 100, 0
        ) if total_logged > 0 else 0

        # ---------- METRICS ----------
        st.subheader("📈 Utilization Summary")

        col1, col2 = st.columns(2)

        col1.metric("Total Hours (Month)", f"{total_logged:.1f} h")
        col2.metric("Utilization (Excl. S9 Work Order)", f"{utilization}%")
