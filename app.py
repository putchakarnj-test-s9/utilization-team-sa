import streamlit as st
import pandas as pd
import re
import matplotlib.pyplot as plt
from matplotlib.patches import Patch
from datetime import datetime

# ---------- PAGE CONFIG ----------
st.set_page_config(page_title="🖥 Team Utilization Dashboard", layout="wide")

# ---------- CATEGORY CONFIG ----------
# Each issue is classified by checking its title against these keyword rules,
# IN ORDER (first match wins). Edit / reorder these to tune the classification.
CATEGORY_RULES = [
    ("external", "External Meeting"),       # e.g. "External Client Activity"
    ("internal meeting", "Internal Meeting"),  # e.g. "Internal Meeting"
    ("ลา", "Leave"),                        # Thai: leave
    ("leave", "Leave"),                     # e.g. "Leave (ลาพักร้อน/ลาทุกชนิด)"
    ("training", "Internal Meeting"),       # e.g. "Training / Meeting / กิจกรรมบริษัท"
    ("meeting", "Internal Meeting"),        # any other meeting -> internal
    ("กิจกรรม", "Internal Meeting"),         # Thai: company activity
]
DEFAULT_CATEGORY = "Task"

CATEGORY_COLORS = {
    "Task": "#22c55e",             # green
    "Internal Meeting": "#3b82f6",  # blue
    "External Meeting": "#f59e0b",  # amber
    "Leave": "#94a3b8",            # gray
}


def categorize_issue(issue):
    """Classify a single issue title into a category."""
    s = str(issue).lower()
    for keyword, category in CATEGORY_RULES:
        if keyword in s:
            return category
    return DEFAULT_CATEGORY


# ---------- MONTH DISPLAY (FROM FILE NAME) ----------
def extract_month_from_filename(filename):
    match = re.search(r"(\d{2})\.(\d{2})\.(\d{4})_(\d{2})\.(\d{2})\.(\d{4})", filename)

    if match:
        day, month, year = match.group(1), match.group(2), match.group(3)
        try:
            date_obj = datetime.strptime(f"{day}.{month}.{year}", "%d.%m.%Y")
            return date_obj.strftime("%B %Y")
        except Exception:
            return None
    return None


# Default title
st.title("Utilization Dashboard")


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

    # Jira/Tempo exports leave merged cells blank for repeated User / Project /
    # Issues values -> forward-fill so every task row carries its context.
    df["User"] = df["User"].ffill()
    df["Project"] = df["Project"].ffill()

    has_issues = "Issues" in df.columns
    if has_issues:
        df["Issues"] = df["Issues"].astype(str).str.strip()

    # Exclude subtotal / grand-total / summary rows
    df = df[df["Project"].astype(str).str.strip().str.lower() != "total"]
    df = df[df["User"].astype(str).str.strip().str.lower() != "summary"]
    if has_issues:
        df = df[df["Issues"].str.lower() != "total"]
        df = df[df["Issues"].str.lower() != "nan"]

    # Keep only rows that actually have logged time
    df = df[df["Total"].notna()]
    df["Hours"] = df["Total"].apply(parse_time_to_hours)
    df = df[df["Hours"] > 0]

    if has_issues:
        df["Category"] = df["Issues"].apply(categorize_issue)

    return df


# ---------- COLOR LOGIC (PROJECTS) ----------
def get_color(project):
    p = str(project).lower()
    if "s9 - work order" in p:
        return "#ef4444"  # red
    elif "s9 - tech support" in p:
        return "#3b82f6"  # blue
    return "#22c55e"      # green


# ---------- 100% STACKED BAR HELPER ----------
def stacked_100_bar(segments, colors, figsize=(12, 1.9), min_label_pct=5):
    """segments: list of (label, hours). Draws one horizontal bar = 100%."""
    total = sum(h for _, h in segments)
    fig, ax = plt.subplots(figsize=figsize)

    left = 0
    for (label, hours), color in zip(segments, colors):
        pct = (hours / total * 100) if total > 0 else 0
        ax.barh(0, pct, left=left, color=color, edgecolor="white")
        if pct >= min_label_pct:
            ax.text(
                left + pct / 2, 0,
                f"{label}\n{pct:.1f}%",
                ha="center", va="center",
                fontsize=8, color="white", fontweight="bold",
            )
        left += pct

    ax.set_xlim(0, 100)
    ax.set_ylim(-0.5, 0.5)
    ax.set_yticks([])
    ax.set_xlabel("% of total")
    for spine in ["top", "right", "left"]:
        ax.spines[spine].set_visible(False)
    fig.tight_layout()
    return fig


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

        has_issues = "Issues" in df.columns and "Category" in df.columns

        # ---------- USER SELECT ----------
        users = sorted(df["User"].unique())
        selected_user = st.selectbox("🙎 Select User", users)

        user_df = df[df["User"] == selected_user]

        # ---------- PROJECT SUMMARY ----------
        project_summary = (
            user_df.groupby("Project")["Hours"]
            .sum()
            .reset_index()
            .sort_values(by="Hours", ascending=False)
        )

        total_logged = project_summary["Hours"].sum()

        # ---------- TABLE: HOURS PER PROJECT ----------
        display_table = project_summary.copy()
        display_table["% of Total"] = (
            display_table["Hours"] / total_logged * 100
        ).round(1) if total_logged > 0 else 0
        total_row = pd.DataFrame({
            "Project": ["TOTAL"],
            "Hours": [total_logged],
            "% of Total": [100.0 if total_logged > 0 else 0],
        })
        display_table = pd.concat([display_table, total_row], ignore_index=True)

        st.subheader(f"⏱ {selected_user} — Hours per Project")
        st.dataframe(display_table, hide_index=True, use_container_width=True)

        # ---------- 100% STACKED BAR: ALL PROJECTS = 100% ----------
        st.subheader("📌 Project Distribution — All Projects = 100%")

        segments = list(zip(project_summary["Project"], project_summary["Hours"]))
        colors = [get_color(p) for p, _ in segments]
        fig = stacked_100_bar(segments, colors, figsize=(12, 2.0))
        st.pyplot(fig)

        # ---------- S9 - WORK ORDER FOCUSED DRILL-DOWN ----------
        # Other projects intentionally show only the project-level totals above.
        # The per-issue drill-down is reserved for S9 - Work Order.
        if has_issues:
            s9_mask = user_df["Project"].str.lower().str.contains("s9 - work order")
            s9_df = user_df[s9_mask]

            if not s9_df.empty:
                s9_project_name = s9_df["Project"].iloc[0]
                s9_total = s9_df["Hours"].sum()

                st.subheader(f"🔍 {s9_project_name} — Detailed Breakdown")
                st.caption(
                    f"Total {s9_total:.2f} h "
                    f"({s9_total / total_logged * 100:.1f}% of "
                    f"{selected_user}'s logged time)"
                )

                cat_legend = [
                    Patch(facecolor=CATEGORY_COLORS[c], label=c)
                    for c in ["Task", "Internal Meeting", "External Meeting", "Leave"]
                ]

                # --- 100% stacked bar by category within S9 - Work Order ---
                cat_summary = (
                    s9_df.groupby("Category")["Hours"]
                    .sum()
                    .reset_index()
                    .sort_values(by="Hours", ascending=False)
                )
                cat_segments = list(
                    zip(cat_summary["Category"], cat_summary["Hours"])
                )
                cat_colors = [
                    CATEGORY_COLORS.get(c, "#999999") for c, _ in cat_segments
                ]
                fig_cat = stacked_100_bar(
                    cat_segments, cat_colors, figsize=(12, 1.6)
                )
                fig_cat.axes[0].legend(
                    handles=cat_legend, loc="upper center",
                    bbox_to_anchor=(0.5, -0.55), ncol=4, frameon=False,
                    fontsize=8,
                )
                st.pyplot(fig_cat)

                # --- horizontal bar chart: each issue by total hours ---
                issue_summary = (
                    s9_df.groupby(["Issues", "Category"])["Hours"]
                    .sum()
                    .reset_index()
                    .sort_values(by="Hours", ascending=True)  # smallest at bottom
                )
                bar_colors = [
                    CATEGORY_COLORS.get(c, "#999999")
                    for c in issue_summary["Category"]
                ]
                fig_iss, ax_iss = plt.subplots(
                    figsize=(12, max(2.5, 0.5 * len(issue_summary)))
                )
                bars = ax_iss.barh(
                    issue_summary["Issues"],
                    issue_summary["Hours"],
                    color=bar_colors,
                )
                max_hours = issue_summary["Hours"].max() if len(issue_summary) else 1
                for bar, h in zip(bars, issue_summary["Hours"]):
                    pct = h / s9_total * 100 if s9_total > 0 else 0
                    ax_iss.text(
                        bar.get_width() + max_hours * 0.01,
                        bar.get_y() + bar.get_height() / 2,
                        f"{h:.1f} h  ({pct:.1f}%)",
                        va="center", fontsize=9,
                    )
                ax_iss.set_xlim(0, max_hours * 1.18)
                ax_iss.set_xlabel("Hours")
                ax_iss.set_title("Issues in S9 - Work Order, by total hours")
                for spine in ["top", "right"]:
                    ax_iss.spines[spine].set_visible(False)
                ax_iss.legend(
                    handles=cat_legend, loc="lower right",
                    fontsize=8, frameon=False,
                )
                fig_iss.tight_layout()
                st.pyplot(fig_iss)

                # --- detail table: every issue in S9 - Work Order ---
                detail = s9_df[["Issues", "Category", "Total", "Hours"]].copy()
                detail["% of S9"] = (
                    detail["Hours"] / s9_total * 100
                ).round(1) if s9_total > 0 else 0
                detail = detail.sort_values(by="Hours", ascending=False)
                detail = detail.rename(columns={"Total": "Logged"})

                total_row = pd.DataFrame({
                    "Issues": ["TOTAL"],
                    "Category": [""],
                    "Logged": [""],
                    "Hours": [s9_total],
                    "% of S9": [100.0],
                })
                detail = pd.concat([detail, total_row], ignore_index=True)

                st.dataframe(
                    detail, hide_index=True, use_container_width=True
                )
            else:
                st.info(
                    "ℹ️ No **S9 - Work Order** entries found for this user."
                )
        else:
            st.info(
                "ℹ️ Add an **Issues** column to your file to see the "
                "S9 - Work Order detailed breakdown."
            )

        # ---------- UTILIZATION ----------
        non_s9_work_order_hours = user_df[
            ~user_df["Project"].str.lower().str.contains("s9 - work order")
        ]["Hours"].sum()

        utilization = round(
            (non_s9_work_order_hours / total_logged) * 100, 0
        ) if total_logged > 0 else 0

        # ---------- METRICS ----------
        st.subheader("📈 Utilization Summary")

        col1, col2, col3 = st.columns(3)
        col1.metric("Total Hours (Month)", f"{total_logged:.1f} h")
        col2.metric("Utilization (Excl. S9 Work Order)", f"{utilization}%")

        if has_issues:
            meeting_hours = user_df[
                user_df["Category"].isin(["Internal Meeting", "External Meeting"])
            ]["Hours"].sum()
            meeting_pct = round(
                meeting_hours / total_logged * 100, 1
            ) if total_logged > 0 else 0
            col3.metric("Time in Meetings", f"{meeting_pct}%")
