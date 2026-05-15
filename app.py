import streamlit as st
import pandas as pd
import re
import matplotlib.pyplot as plt
from matplotlib.patches import Patch
from datetime import datetime

# ---------- PAGE CONFIG ----------
st.set_page_config(
    page_title="🖥 Team Utilization Dashboard",
    layout="wide"
)

# ---------- MONTH DISPLAY (FROM FILE NAME) ----------
def extract_month_from_filename(filename):

    match = re.search(
        r"(\d{2})\.(\d{2})\.(\d{4})_(\d{2})\.(\d{2})\.(\d{4})",
        filename
    )

    if match:

        day = match.group(1)
        month = match.group(2)
        year = match.group(3)

        try:
            date_obj = datetime.strptime(
                f"{day}.{month}.{year}",
                "%d.%m.%Y"
            )

            return date_obj.strftime("%B %Y")

        except:
            return None

    return None


# ---------- DEFAULT TITLE ----------
st.title("🖥 Utilization Dashboard")


# ---------- TIME PARSER ----------
def parse_time_to_hours(time_str):

    if pd.isna(time_str):
        return 0

    text = str(time_str).lower()

    total_hours = 0

    week_match = re.search(r'(\d+)w', text)
    day_match = re.search(r'(\d+)d', text)
    hour_match = re.search(r'(\d+)h', text)
    minute_match = re.search(r'(\d+)m', text)

    if week_match:
        total_hours += int(week_match.group(1)) * 40

    if day_match:
        total_hours += int(day_match.group(1)) * 8

    if hour_match:
        total_hours += int(hour_match.group(1))

    if minute_match:
        total_hours += int(minute_match.group(1)) / 60

    return round(total_hours, 2)


# ---------- FILE READER ----------
def load_file(file):

    filename = file.name.lower()

    if filename.endswith(".xlsx"):
        return pd.read_excel(file)

    elif filename.endswith(".csv"):
        return pd.read_csv(file)

    else:
        return pd.read_csv(file, sep="\t")


# ---------- TRANSFORM ----------
def transform(df):

    df.columns = [str(c).strip() for c in df.columns]

    required_columns = {"User", "Project", "Total"}

    if not required_columns.issubset(set(df.columns)):

        st.error(
            "❌ File must contain columns: User, Project, Total"
        )

        return None

    # Remove unwanted rows
    df = df[
        df["Project"]
        .astype(str)
        .str.strip()
        .str.lower() != "total"
    ]

    df = df[
        df["User"]
        .astype(str)
        .str.strip()
        .str.lower() != "summary"
    ]

    # Convert to hours
    df["Hours"] = df["Total"].apply(
        parse_time_to_hours
    )

    return df


# ---------- FILE UPLOAD ----------
uploaded_file = st.file_uploader(
    "Upload file (Excel / CSV / TXT)",
    type=["xlsx", "csv", "txt"]
)


# ---------- MAIN ----------
if uploaded_file is not None:

    # ---------- TITLE ----------
    month_label = extract_month_from_filename(
        uploaded_file.name
    )

    if month_label:
        st.title(
            f"📊 Team Utilization Dashboard - {month_label}"
        )

    else:
        st.title("📊 Team Utilization Dashboard")

    # ---------- LOAD DATA ----------
    try:

        df_raw = load_file(uploaded_file)

        df = transform(df_raw)

        if df is not None and not df.empty:

            # ---------- USER LIST ----------
            users = sorted(
                df["User"]
                .dropna()
                .astype(str)
                .unique()
            )

            if len(users) == 0:
                st.warning("No users found.")
                st.stop()

            selected_user = st.selectbox(
                "🙎 Select User",
                users
            )

            # ---------- FILTER USER ----------
            user_df = df[
                df["User"] == selected_user
            ].copy()

            if user_df.empty:
                st.warning("No data found for selected user.")
                st.stop()

            # ---------- PROJECT SUMMARY ----------
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

            # ---------- TOTAL ROW ----------
            total_row = pd.DataFrame({
                "Project": ["TOTAL"],
                "Hours": [total_logged]
            })

            display_table = pd.concat(
                [project_summary, total_row],
                ignore_index=True
            )

            # ---------- TABLE ----------
            st.subheader(
                f"⏱ {selected_user} - Hours per Project"
            )

            st.dataframe(
                display_table,
                hide_index=True,
                use_container_width=True
            )

            # ---------- COLOR LOGIC ----------
            def get_color(project):

                p = str(project).lower()

                if "s9 - work order" in p:
                    return "#ef4444"

                elif "s9 - tech support" in p:
                    return "#3b82f6"

                return "#22c55e"

            colors = [
                get_color(p)
                for p in project_summary["Project"]
            ]

            # ---------- BAR CHART ----------
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

                graph_max = (
                    total_logged
                    if total_logged > 0
                    else 1
                )

                ax.set_xlim(0, graph_max)

                ax.set_xlabel("Hours")
                ax.set_ylabel("Project")

                ax.set_title(
                    "Hours per Project"
                )

                legend_elements = [
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
                    handles=legend_elements
                )

                st.pyplot(fig)

            # ---------- PIE CHART ----------
            st.subheader(
                "📌 Project Distribution"
            )

            col4, col5, col6 = st.columns(
                [1, 2, 1]
            )

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

            # ---------- UTILIZATION ----------
            non_s9_work_order_hours = user_df[
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
                        non_s9_work_order_hours
                        / total_logged
                    ) * 100,
                    0
                )

            else:
                utilization = 0

            # ---------- METRICS ----------
            st.subheader(
                "📈 Utilization Summary"
            )

            metric_col1, metric_col2 = st.columns(2)

            metric_col1.metric(
                "Total Hours (Month)",
                f"{total_logged:.1f} h"
            )

            metric_col2.metric(
                "Utilization (Excl. S9 Work Order)",
                f"{utilization:.0f}%"
            )

            # ---------- OPTIONAL S9 BREAKDOWN ----------
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

                    # Extract SWO Type
                    s9_df["SWO Type"] = (
                        s9_df["Issues"]
                        .astype(str)
                        .str.extract(
                            r"(SWO-\d+)",
                            expand=False
                        )
                    )

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

                        total_swo = round(
                            swo_summary["Hours"].sum(),
                            1
                        )

                        if total_swo > 0:

                            swo_summary["Percentage"] = (
                                swo_summary["Hours"]
                                / total_swo
                                * 100
                            ).round(0)

                            st.subheader(
                                "📌 S9 Work Order Breakdown"
                            )

                            if len(swo_summary) > 0:

                                cols = st.columns(
                                    len(swo_summary)
                                )

                                for idx, row in swo_summary.iterrows():

                                    cols[idx].metric(
                                        row["SWO Type"],
                                        f'{row["Percentage"]:.0f}%'
                                    )

        else:
            st.warning("No valid data found.")

    except Exception as e:

        st.error("❌ Error processing file")

        st.exception(e)
