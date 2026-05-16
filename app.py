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


def hours_to_jira_format(hours):
    """Inverse of parse_time_to_hours: 50 -> '1w 1d 2h', 29 -> '3d 5h'."""
    if hours is None or pd.isna(hours) or hours <= 0:
        return "0h"
    minutes_total = int(round(float(hours) * 60))
    weeks, rem = divmod(minutes_total, 40 * 60)
    days, rem = divmod(rem, 8 * 60)
    hrs, mins = divmod(rem, 60)
    parts = []
    if weeks:
        parts.append(f"{weeks}w")
    if days:
        parts.append(f"{days}d")
    if hrs:
        parts.append(f"{hrs}h")
    if mins:
        parts.append(f"{mins}m")
    return " ".join(parts) if parts else "0h"


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

    # ----- Flexible "Issues" column detection -----
    # Different Jira / Tempo exports name this column differently:
    #   "Issues", "Issue", "Issue Summary", "Summary", or split into
    #   "Issue Key" + "Issue Summary". Match case-insensitively, combine
    #   key + summary into "KEY: Summary" when both are present.
    cols_lower = {c.lower(): c for c in df.columns}

    def _find(names):
        for n in names:
            if n.lower() in cols_lower:
                return cols_lower[n.lower()]
        return None

    issue_name_col = _find(
        ["Issues", "Issue", "Issue Summary", "Summary", "Worklog Description"]
    )
    issue_key_col = _find(["Issue Key", "Key", "Issue ID"])
    issues_source = None

    if issue_name_col and issue_key_col and issue_name_col != issue_key_col:
        keys = df[issue_key_col].fillna("").astype(str).str.strip()
        names = df[issue_name_col].fillna("").astype(str).str.strip()
        combined = []
        for k, n in zip(keys, names):
            if k and n:
                # Avoid double "KEY: KEY: ..." if name already starts with key
                if n.startswith(f"{k}:") or n.startswith(f"{k} "):
                    combined.append(n)
                else:
                    combined.append(f"{k}: {n}")
            elif k:
                combined.append(k)
            else:
                combined.append(n)
        df["Issues"] = combined
        issues_source = f"{issue_key_col} + {issue_name_col}"
    elif issue_name_col:
        df["Issues"] = df[issue_name_col].astype(str).str.strip()
        issues_source = issue_name_col
    elif issue_key_col:
        df["Issues"] = df[issue_key_col].astype(str).str.strip()
        issues_source = issue_key_col

    has_issues = issues_source is not None
    if has_issues:
        # Empty / placeholder values -> friendly label, so project-level
        # entries (no issue attached) still appear in the breakdown.
        df["Issues"] = df["Issues"].replace(
            {"": "(no issue)", "nan": "(no issue)",
             "None": "(no issue)", "<NA>": "(no issue)"}
        )
        # Stash for the UI to display in a caption
        df.attrs["issues_source_column"] = issues_source
    else:
        df.attrs["all_columns"] = list(df.columns)

    # Exclude subtotal / grand-total / summary rows
    df = df[df["Project"].astype(str).str.strip().str.lower() != "total"]
    df = df[df["User"].astype(str).str.strip().str.lower() != "summary"]
    if has_issues:
        df = df[df["Issues"].str.lower() != "total"]

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

        # ---------- COLUMN-DETECTION FEEDBACK ----------
        # Tell the user which column was read as the issue name (or warn
        # if none was found, so they can rename their export).
        if has_issues:
            src = df.attrs.get("issues_source_column", "Issues")
            if src.lower() != "issues":
                st.caption(
                    f"📎 Reading issue names from column **{src}**."
                )
        else:
            all_cols = df.attrs.get("all_columns", list(df.columns))
            st.warning(
                "⚠️ No issue-name column detected. The S9 - Work Order "
                "drill-down needs one of: **Issues**, **Issue**, "
                "**Issue Summary**, **Issue Key**, or **Summary**. "
                f"Columns in your file: {', '.join(all_cols)}"
            )

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
        # Other projects intentionally show only the project-level totals
        # above. The per-issue drill-down is reserved for S9 - Work Order.
        if has_issues:
            s9_mask = user_df["Project"].str.lower().str.contains("s9 - work order")
            s9_df = user_df[s9_mask]

            if not s9_df.empty:
                s9_project_name = s9_df["Project"].iloc[0]
                s9_total = s9_df["Hours"].sum()

                st.subheader(f"🔍 {s9_project_name} — Detailed Breakdown")
                st.caption(
                    f"Total {hours_to_jira_format(s9_total)} "
                    f"({s9_total:.2f} h, "
                    f"{s9_total / total_logged * 100:.1f}% of "
                    f"{selected_user}'s logged time)"
                )

                # --- Aggregate by issue: one row per unique issue ---
                issue_summary = (
                    s9_df.groupby("Issues", as_index=False)["Hours"]
                    .sum()
                    .sort_values("Hours", ascending=False)
                )
                issue_summary["% of S9"] = (
                    issue_summary["Hours"] / s9_total * 100
                ).round(1) if s9_total > 0 else 0
                issue_summary["Logged"] = issue_summary["Hours"].apply(
                    hours_to_jira_format
                )

                # --- horizontal bar chart: each issue, share of S9 ---
                plot_df = issue_summary.sort_values("Hours", ascending=True)
                # Color gradient: largest share darkest red, smallest lightest
                n = len(plot_df)
                if n > 0:
                    # Map % -> shade: 100% -> deep red, smaller -> lighter
                    pcts = plot_df["% of S9"].to_numpy()
                    max_pct = pcts.max() if pcts.max() > 0 else 1
                    shades = []
                    for p in pcts:
                        # Light pink -> deep red based on share
                        ratio = p / max_pct
                        r = int(239 - (239 - 185) * (1 - ratio))   # 185..239
                        g = int(68 + (200 - 68) * (1 - ratio))     # 68..200
                        b = int(68 + (200 - 68) * (1 - ratio))     # 68..200
                        shades.append(f"#{r:02x}{g:02x}{b:02x}")
                else:
                    shades = []

                fig_iss, ax_iss = plt.subplots(
                    figsize=(12, max(2.5, 0.6 * max(n, 1)))
                )
                bars = ax_iss.barh(
                    plot_df["Issues"], plot_df["Hours"], color=shades,
                )
                max_hours = plot_df["Hours"].max() if n else 1
                for bar, h, pct in zip(
                    bars, plot_df["Hours"], plot_df["% of S9"]
                ):
                    ax_iss.text(
                        bar.get_width() + max_hours * 0.01,
                        bar.get_y() + bar.get_height() / 2,
                        f"{hours_to_jira_format(h)}  ·  {pct:.1f}%",
                        va="center", fontsize=10,
                    )
                ax_iss.set_xlim(0, max_hours * 1.30)
                ax_iss.set_xlabel("Hours")
                ax_iss.set_title(
                    "Issues in S9 - Work Order — share of S9 total"
                )
                for spine in ["top", "right"]:
                    ax_iss.spines[spine].set_visible(False)
                fig_iss.tight_layout()
                st.pyplot(fig_iss)

                # --- detail table: Issue / Logged / Hours / % of S9 ---
                detail = issue_summary[
                    ["Issues", "Logged", "Hours", "% of S9"]
                ].copy()
                detail["Hours"] = detail["Hours"].round(2)

                total_row = pd.DataFrame({
                    "Issues": ["TOTAL"],
                    "Logged": [hours_to_jira_format(s9_total)],
                    "Hours": [round(s9_total, 2)],
                    "% of S9": [100.0],
                })
                detail = pd.concat([detail, total_row], ignore_index=True)

                st.dataframe(
                    detail, hide_index=True, use_container_width=True
                )

                # Friendly hint if the export didn't include any issue names
                if (issue_summary["Issues"] == "(no issue)").any():
                    st.caption(
                        "⚠️ Some S9 - Work Order entries have no issue name "
                        "in the export and are grouped as **(no issue)**. "
                        "If you expected issue codes here, check that the "
                        "Tempo export includes the **Issues** column."
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
