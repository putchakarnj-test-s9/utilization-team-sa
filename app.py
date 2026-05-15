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

    week_match = re.search(r"(\d+)w", text)
    if week_match:
        total_hours += int(week_match.group(1)) * 40

    day_match = re.search(r"(\d+)d", text)
    if day_match:
        total_hours += int(day_match.group(1)) * 8

    hour_match = re.search(r"(\d+)h", text)
    if hour_match:
        total_hours += int(hour_match.group(1))

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

    # clean columns
    df.columns = [str(c).strip() for c in df.columns]

    required_columns = [
        "User",
        "Project",
        "Total"
    ]

    for col in required_columns:

        if col not in df.columns:

            st.error(
                f"❌ Missing required column: {col}"
            )

            return None

    # remove SUMMARY rows
    df = df[
        df["User"]
        .astype(str)
        .str.strip()
        .str.lower() != "summary"
    ]

    # remove TOTAL rows
    df = df[
        df["Project"]
        .astype(str)
        .str.strip()
        .str.lower() != "total"
    ]

    # remove issue = total
    if "Issues" in df.columns:

        df = df[
            ~df["Issues"]
            .astype(str)
            .str.strip()
            .str.lower()
            .eq("total")
        ]

    # convert hours
    df["Hours"] = df["Total"].apply(
        parse_time_to_hours
    )

    return df


# =========================================================
# PROJECT COLOR
# =========================================================
def get_project_color(project_name):

    project = str(project_name).lower()

    if "s9 - work order" in project:
        return "#ef4444"

    elif "s9 - tech support" in project:
        return "#3b82f6"

    return "#22c55e"


# =========================================================
# EXTRACT SWO DETAIL
# =========================================================
def extract_swo_detail(issue_text):

    text = str(issue_text).strip()

    # normalize spaces only
    text = re.sub(r"\s+", " ", text)

    # ---------------------------------------------
    # detect SWO type
    # ---------------------------------------------
    swo_match = re.search(
        r"(SWO\s*-\s*\d+)",
        text,
        re.IGNORECASE
    )

    if swo_match:

        swo_type = (
            swo_match.group(1)
            .upper()
            .replace(" ", "")
        )

    else:

        swo_type = "Other"

    # ---------------------------------------------
    # description = original issue label
    # remove only SWO code prefix
    # ---------------------------------------------
    description = re.sub(
        r"^SWO\s*-\s*\d+\s*:?\s*",
        "",
        text,
        flags=re.IGNORECASE
    ).strip()

    # fallback
    if description == "":
        description = text

    return pd.Series([
        swo_type,
        description
    ])

# =========================================================
# FILE UPLOADER
# =========================================================
uploaded_file = st.file_uploader(
    "Upload Excel / CSV File",
    type=["xlsx", "csv", "txt"]
)

# =========================================================
# MAIN
# =========================================================
if uploaded_file is not None:

    # =====================================================
    # MONTH LABEL
    # =====================================================
    month_label = extract_month_from_filename(
        uploaded_file.name
    )

    if month_label:

        st.title(
            f"📊 Team Utilization Dashboard - {month_label}"
        )

    # =====================================================
    # LOAD DATA
    # =====================================================
    raw_df = load_file(uploaded_file)

    # =====================================================
    # TRANSFORM DATA
    # =====================================================
    df = transform_data(raw_df)

    if df is not None and not df.empty:

        # =================================================
        # USER DROPDOWN
        # =================================================
        users = sorted(
            df["User"]
            .dropna()
            .astype(str)
            .unique()
        )

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
        # SUMMARY TABLE
        # =================================================
        st.subheader(
            f"⏱ {selected_user} - Hours per Project"
        )

        total_row = pd.DataFrame({
            "Project": ["TOTAL"],
            "Hours": [total_logged]
        })

        display_table = pd.concat(
            [project_summary, total_row],
            ignore_index=True
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

        col1, col2, col3 = st.columns([1, 3, 1])

        with col2:

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
                )
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

        col4, col5, col6 = st.columns([1, 2, 1])

        with col5:

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
        # S9 WORK ORDER DETAILS
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

                st.subheader(
                    "📝 S9 Work Order Breakdown"
                )

                # extract details
                s9_df[
                    ["SWO Type", "Description"]
                ] = s9_df["Issues"].apply(
                    extract_swo_detail
                )

                # summary
                detail_summary = (
                    s9_df
                    .groupby(
                        [
                            "SWO Type",
                            "Description"
                        ]
                    )["Hours"]
                    .sum()
                    .reset_index()
                )

                # total by type
                detail_summary["Type Total"] = (
                    detail_summary
                    .groupby("SWO Type")["Hours"]
                    .transform("sum")
                )

                # percentage
                detail_summary["Percentage"] = (
                    detail_summary["Hours"]
                    / detail_summary["Type Total"]
                    * 100
                ).round(0)

                # sort
                detail_summary = detail_summary.sort_values(
                    by=[
                        "SWO Type",
                        "Hours"
                    ],
                    ascending=[True, False]
                )

                # display
                st.dataframe(
                    detail_summary[
                        [
                            "SWO Type",
                            "Description",
                            "Hours",
                            "Percentage"
                        ]
                    ],
                    hide_index=True,
                    use_container_width=True
                )

    else:

        st.warning(
            "No valid data found"
        )
