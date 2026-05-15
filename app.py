import streamlit as st
import pandas as pd
import re
import matplotlib.pyplot as plt
from matplotlib.patches import Patch
from datetime import datetime

# =========================================================
# PAGE CONFIG
# =========================================================
st.set_page_config(
    page_title="Team Utilization Dashboard",
    layout="wide"
)

# =========================================================
# DEFAULT TITLE
# =========================================================
st.title("🖥 Team Utilization Dashboard")

# =========================================================
# EXTRACT MONTH FROM FILE NAME
# Example:
# worklogs_01.03.2026_31.03.2026.xlsx
# =========================================================
def extract_month_from_filename(filename):

    pattern = r"(\d{2})\.(\d{2})\.(\d{4})_(\d{2})\.(\d{2})\.(\d{4})"

    match = re.search(pattern, filename)

    if match:

        start_day = match.group(1)
        start_month = match.group(2)
        start_year = match.group(3)

        try:

            date_obj = datetime.strptime(
                f"{start_day}.{start_month}.{start_year}",
                "%d.%m.%Y"
            )

            return date_obj.strftime("%B %Y")

        except:
            return None

    return None


# =========================================================
# TIME PARSER
# 1w = 40h
# 1d = 8h
# =========================================================
def parse_time_to_hours(value):

    if pd.isna(value):
        return 0

    text = str(value).lower()

    total_hours = 0

    # weeks
    week_match = re.search(r"(\d+)w", text)
    if week_match:
        total_hours += int(week_match.group(1)) * 40

    # days
    day_match = re.search(r"(\d+)d", text)
    if day_match:
        total_hours += int(day_match.group(1)) * 8

    # hours
    hour_match = re.search(r"(\d+)h", text)
    if hour_match:
        total_hours += int(hour_match.group(1))

    # minutes
    minute_match = re.search(r"(\d+)m", text)
    if minute_match:
        total_hours += int(minute_match.group(1)) / 60

    return round(total_hours, 2)


# =========================================================
# LOAD FILE
# =========================================================
def load_file(uploaded_file):

    filename = uploaded_file.name.lower()

    try:

        if filename.endswith(".xlsx"):
            return pd.read_excel(uploaded_file)

        elif filename.endswith(".csv"):
            return pd.read_csv(uploaded_file)

        else:
            return pd.read_csv(uploaded_file, sep="\t")

    except Exception as e:

        st.error("❌ Unable to read uploaded file")
        st.exception(e)
        return None


# =========================================================
# TRANSFORM DATA
# =========================================================
def transform_data(df):

    if df is None:
        return None

    # clean column names
    df.columns = [str(c).strip() for c in df.columns]

    required_columns = ["User", "Project", "Total"]

    for col in required_columns:

        if col not in df.columns:
            st.error(
                f"❌ Missing required column: {col}"
            )
            return None

    # remove TOTAL rows
    df = df[
        df["Project"]
        .astype(str)
        .str.strip()
        .str.lower() != "total"
    ]

    # remove SUMMARY user
    df = df[
        df["User"]
        .astype(str)
        .str.strip()
        .str.lower() != "summary"
    ]

    # create hours column
    df["Hours"] = df["Total"].apply(
        parse_time_to_hours
    )

    return df


# =========================================================
# COLOR LOGIC
# =========================================================
def get_project_color(project_name):

    project = str(project_name).lower()

    if "s9 - work order" in project:
        return "#ef4444"  # red

    elif "s9 - tech support" in project:
        return "#3b82f6"  # blue

    return "#22c55e"      # green


# =========================================================
# FILE UPLOAD
# =========================================================
uploaded_file = st.file_uploader(
    "Upload Excel / CSV File",
    type=["xlsx", "csv", "txt"]
)

# =========================================================
# MAIN
# =========================================================
if uploaded_file is not None:

    # -----------------------------------------------------
    # UPDATE TITLE FROM FILE NAME
    # -----------------------------------------------------
    month_label = extract_month_from_filename(
        uploaded_file.name
    )

    if month_label:
        st.title(
            f"📊 Team Utilization Dashboard - {month_label}"
        )

    # -----------------------------------------------------
    # LOAD DATA
    # -----------------------------------------------------
    raw_df = load_file(uploaded_file)

    # -----------------------------------------------------
    # TRANSFORM
    # -----------------------------------------------------
    df = transform_data(raw_df)

    if df is not None and not df.empty:

        # =================================================
        # USER SELECT
        # =================================================
        users = sorted(
            df["User"]
            .dropna()
            .astype(str)
            .unique()
        )

        if len(users) == 0:
            st.warning("No users found")
            st.stop()

        selected_user = st.selectbox(
            "🙎 Select User",
            users
        )

        # =================================================
        # FILTER USER DATA
        # =================================================
        user_df = df[
            df["User"] == selected_user
        ].copy()

        if user_df.empty:
            st.warning("No data found")
            st.stop()

        # =================================================
        # PROJECT SUMMARY
        # =================================================
        project_summary = (
            user_df
            .groupby("Project")["Hours"]
            .sum()
            .reset_index()
            .sort_values(by="Project")
        )

        total_logged = round(
            project_summary["Hours"].sum(),
            1
        )

        # =================================================
        # DISPLAY TABLE
        # =================================================
        total_row = pd.DataFrame({
            "Project": ["TOTAL"],
            "Hours": [total_logged]
        })

        display_table = pd.concat(
            [project_summary, total_row],
            ignore_index=True
        )

        st.subheader(
            f"⏱ {selected_user} - Hours per Project"
        )

        st.dataframe(
            display_table,
            hide_index=True,
            use_container_width=True
        )

        # =================================================
        # BAR CHART
        # =================================================
        colors = [
            get_project_color(p)
            for p in project_summary["Project"]
        ]

        left, center, right = st.columns([1, 3, 1])

        with center:

            fig, ax = plt.subplots(
                figsize=(10, 5)
            )

            ax.barh(
                project_summary["Project"],
                project_summary["Hours"],
                color=colors
            )

            max_hours = total_logged

            if max_hours <= 0:
                max_hours = 1

            ax.set_xlim(0, max_hours)

            ax.set_xlabel("Hours")
            ax.set_ylabel("Project")

            ax.set_title(
                "Hours per Project"
            )

            legend_items = [
                Patch(
                    facecolor="#ef4444",
                    label="S9 - Work Order"
                ),
                Patch(
                    facecolor="#3b82f6",
                    label="S9 - Tech Support"
                ),
                Patch(
                    facecolor="#22c55e",
                    label="Other Projects"
                ),
            ]

            ax.legend(
                handles=legend_items
            )

            st.pyplot(fig)

        # =================================================
        # PIE CHART
        # =================================================
        st.subheader(
            "📌 Project Distribution"
        )

        left2, center2, right2 = st.columns(
            [1, 2, 1]
        )

        with center2:

            fig2, ax2 = plt.subplots(
                figsize=(5, 5)
            )

            ax2.pie(
                project_summary["Hours"],
                labels=project_summary["Project"],
                autopct="%1.1f%%"
            )

            ax2.set_title(
                "Project Distribution"
            )

            st.pyplot(fig2)

        # =================================================
        # UTILIZATION
        # Exclude ONLY S9 Work Order
        # =================================================
        non_s9_hours = user_df[
            ~user_df["Project"]
            .astype(str)
            .str.lower()
            .str.contains(
                "s9 - work order",
                na=False
            )
        ]["Hours"].sum()

        if total_logged > 0:

            utilization = round(
                (
                    non_s9_hours
                    / total_logged
                ) * 100,
                0
            )

        else:
            utilization = 0

        # =================================================
        # METRICS
        # =================================================
        st.subheader(
            "📈 Utilization Summary"
        )

        metric1, metric2 = st.columns(2)

        metric1.metric(
            "Total Hours (Month)",
            f"{total_logged:.1f} h"
        )

        metric2.metric(
            "Utilization (Exclude S9 Work Order)",
            f"{utilization:.0f}%"
        )

        # =================================================
        # OPTIONAL S9 BREAKDOWN
        # only if Issues column exists
        # =================================================
        if "Issues" in user_df.columns:

            s9_df = user_df[
                user_df["Project"]
                .astype(str)
                .str.lower()
                .str.contains(
                    "s9 - work order",
                    na=False
                )
            ].copy()

            if not s9_df.empty:

                # extract SWO type
                s9_df["SWO Type"] = (
                    s9_df["Issues"]
                    .astype(str)
                    .str.extract(
                        r"(SWO-\d+)",
                        expand=False
                    )
                )

                # remove null
                s9_df = s9_df[
                    s9_df["SWO Type"].notna()
                ]

                if not s9_df.empty:

                    swo_summary = (
                        s9_df
                        .groupby("SWO Type")["Hours"]
                        .sum()
                        .reset_index()
                    )

                    total_swo = swo_summary[
                        "Hours"
                    ].sum()

                    if total_swo > 0:

                        swo_summary["Percentage"] = (
                            swo_summary["Hours"]
                            / total_swo
                            * 100
                        ).round(0)

                        st.subheader(
                            "📌 S9 Work Order Breakdown"
                        )

                        swo_cols = st.columns(
                            max(
                                1,
                                len(swo_summary)
                            )
                        )

                        for idx, row in swo_summary.iterrows():

                            swo_cols[idx].metric(
                                row["SWO Type"],
                                f"{row['Percentage']:.0f}%"
                            )

    else:
        st.warning("No valid data found")
