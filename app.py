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
    """Convert a time value to decimal hours.
    Accepts:
      - NaN -> 0
      - numeric (e.g. 7.5) -> treated as already-decimal hours
        (this is what the Details sheet 'Time spent (hours)' uses)
      - pure-number strings (e.g. '7.5') -> same as numeric
      - Jira-style strings ('1w 1d 2h 30m') -> parsed via regex
    """
    if pd.isna(time_str):
        return 0

    # Already numeric? Decimal hours.
    if isinstance(time_str, (int, float)):
        return float(time_str)

    s = str(time_str).strip()
    if s == "":
        return 0

    # Pure number string -> decimal hours
    try:
        return float(s)
    except ValueError:
        pass

    # Jira / Tempo format: "1w 1d 2h 30m"
    pattern = r'(\d+)w|(\d+)d|(\d+)h|(\d+)m'
    matches = re.findall(pattern, s)

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
    """Read the uploaded file. For .xlsx with both 'Data' and 'Details'
    sheets (Tempo timesheet format), prefer 'Details' because it has the
    per-issue rows. Map its columns to the standard names the rest of the
    pipeline expects."""
    name = file.name.lower()
    
    if name.endswith(".xlsx"):
        try:
            xls = pd.ExcelFile(file)
            sheet_map = {s.lower(): s for s in xls.sheet_names}
 
            if "details" in sheet_map:
                df = pd.read_excel(xls, sheet_name=sheet_map["details"])
 
                # Normalize key columns the downstream code expects
                cols_ci = {c.lower(): c for c in df.columns}
                rename = {}
                if "time spent (hours)" in cols_ci:
                    rename[cols_ci["time spent (hours)"]] = "Total"
                df = df.rename(columns=rename)
 
                # The Details sheet can have 200+ columns; trim to what we
                # actually use so the column picker stays manageable.
                wanted = [
                    "User", "Project", "Total",
                    "Issue key", "Issue Key",
                    "Issue summary", "Issue Summary",
                    "Issues", "Issue", "Summary",
                    "Worklog Description", "Description",
                ]
                keep = [c for c in wanted if c in df.columns]
                if {"User", "Project", "Total"}.issubset(set(keep)):
                    df = df[keep]
                return df
 
            return pd.read_excel(xls, sheet_name=xls.sheet_names[0])
        except Exception:
            return pd.read_excel(file)
    elif name.endswith(".csv"):
        return pd.read_csv(file)
    else:
        return pd.read_csv(file, sep="\t")


# ---------- TRANSFORM ----------
def transform(df, issues_override=None):
    df.columns = [c.strip() for c in df.columns]

    required = {"User", "Project", "Total"}
    if not required.issubset(set(df.columns)):
        st.error("❌ File must contain columns: User, Project, Total")
        return None

    # Jira/Tempo exports leave merged cells blank for repeated User / Project /
    # Issues values -> forward-fill so every task row carries its context.
    df["User"] = df["User"].ffill()
    df["Project"] = df["Project"].ffill()

    # ----- Issue column: manual override wins, then auto-detect -----
    # Auto-detect handles common Jira / Tempo names ("Issues", "Issue",
    # "Issue Summary", "Summary", "Worklog Description", and split
    # "Issue Key" + "Issue Summary"). The UI also lets the user pick the
    # column explicitly, which takes priority.
    cols_lower = {c.lower(): c for c in df.columns}

    def _find(names):
        for n in names:
            if n.lower() in cols_lower:
                return cols_lower[n.lower()]
        return None

    issues_source = None

    if issues_override and issues_override in df.columns:
        df["Issues"] = df[issues_override].astype(str).str.strip()
        issues_source = issues_override
    else:
        issue_name_col = _find(
            ["Issues", "Issue", "Issue Summary", "Summary",
             "Worklog Description", "Description", "Task", "Topic"]
        )
        issue_key_col = _find(["Issue Key", "Key", "Issue ID"])

        if issue_name_col and issue_key_col and issue_name_col != issue_key_col:
            keys = df[issue_key_col].fillna("").astype(str).str.strip()
            names = df[issue_name_col].fillna("").astype(str).str.strip()
            combined = []
            for k, n in zip(keys, names):
                if k and n:
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

    # ---------- ISSUE-NAME COLUMN PICKER ----------
    # Auto-detect first, then let the user override with the actual column
    # that holds issue names (in case the export uses a non-standard header).
    raw_cols = [str(c).strip() for c in df_raw.columns]
    cols_ci = {c.lower(): c for c in raw_cols}
    auto_default = None
    for cand in ["Issues", "Issue", "Issue Summary", "Summary",
                 "Worklog Description", "Description", "Task", "Topic"]:
        if cand.lower() in cols_ci:
            auto_default = cols_ci[cand.lower()]
            break

    picker_options = ["(auto-detect)"] + raw_cols
    default_idx = 0  # "(auto-detect)"
    selected_issue_col = st.selectbox(
        "🏷️ Which column has the issue name?  "
        "(SWO-6: R & D, SWO-4: …, etc.)",
        options=picker_options,
        index=default_idx,
        help=(
            "If your S9 - Work Order table shows '(no issue)', the "
            "auto-detect picked the wrong column. Pick the column that "
            "contains the issue keys / summaries here."
            + (f"  Auto-detect found: **{auto_default}**." if auto_default
               else "  Auto-detect found nothing usable.")
        ),
    )
    issues_override = (
        None if selected_issue_col == "(auto-detect)" else selected_issue_col
    )

    df = transform(df_raw, issues_override=issues_override)

    if df is not None:

        has_issues = "Issues" in df.columns and "Category" in df.columns

        # ---------- COLUMN-DETECTION FEEDBACK ----------
        if has_issues:
            src = df.attrs.get("issues_source_column", "Issues")
            tag = "override" if issues_override else "auto-detect"
            st.caption(
                f"📎 Reading issue names from column **{src}** ({tag})."
            )
        else:
            all_cols = df.attrs.get("all_columns", list(df.columns))
            st.warning(
                "⚠️ No issue-name column was selected or detected. "
                "Pick the right column in the dropdown above. "
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
        # Goal: take every row where Project == "S9 - Work Order", read each
        # issue's name from the Issues column AS-IS (no hardcoded list),
        # sum the hours per issue, and show how each issue's share builds
        # up the S9 - Work Order share of the user's total time.
        #
        # Worked example: if S9 - Work Order = 90% of the user's total time,
        # and SWO-6: R & D = 30% of total, SWO-4: ... = 60% of total, then
        # 30% + 60% = 90% — i.e. issue shares of total sum to the S9 share.
        if has_issues:
            s9_mask = user_df["Project"].str.lower().str.contains(
                "s9 - work order"
            )
            s9_df = user_df[s9_mask]

            if not s9_df.empty:
                s9_project_name = s9_df["Project"].iloc[0]
                s9_total = s9_df["Hours"].sum()
                s9_pct_of_total = (
                    (s9_total / total_logged * 100) if total_logged > 0 else 0
                )
                other_total = total_logged - s9_total

                # Column label = whatever column the issue name was read from
                issues_source = df.attrs.get("issues_source_column", "Issues")
                issues_label = (
                    "Issue" if "+" in issues_source else issues_source
                )

                st.subheader(f"🔍 {s9_project_name} — Issue Breakdown")
                st.caption(
                    f"{s9_project_name} = **{s9_pct_of_total:.1f}%** of "
                    f"{selected_user}'s total time "
                    f"({hours_to_jira_format(s9_total)} of "
                    f"{hours_to_jira_format(total_logged)})"
                )

                # --- Aggregate by issue (names read from the Issues column) ---
                issue_summary = (
                    s9_df.groupby("Issues", as_index=False)["Hours"]
                    .sum()
                    .sort_values("Hours", ascending=False)
                    .reset_index(drop=True)
                )
                issue_summary["Logged"] = issue_summary["Hours"].apply(
                    hours_to_jira_format
                )
                issue_summary["% of S9"] = (
                    (issue_summary["Hours"] / s9_total * 100).round(1)
                    if s9_total > 0 else 0
                )
                issue_summary["% of Total"] = (
                    (issue_summary["Hours"] / total_logged * 100).round(1)
                    if total_logged > 0 else 0
                )

                # --- Red gradient shades for S9 issues (largest = deepest) ---
                n_issues = len(issue_summary)

                def _issue_shade(i, n):
                    if n <= 1:
                        return "#1e3a8a"
                    ratio = i / (n - 1)            # 0 (largest) .. 1 (smallest)
                    r = int(30 + (191 - 30) * ratio)
                    g = int(58 + (219 - 58) * ratio)
                    b = int(138 + (254 - 138) * ratio)
                    return f"#{r:02x}{g:02x}{b:02x}"

                issue_colors = [
                    _issue_shade(i, n_issues) for i in range(n_issues)
                ]

                # --- Per-issue horizontal bar chart (% of total) ---
                plot_df = issue_summary.sort_values("Hours", ascending=True)
                plot_colors = [
                    _issue_shade(
                        n_issues - 1 - list(plot_df.index).index(idx),
                        n_issues,
                    )
                    for idx in plot_df.index
                ]
                fig_iss, ax_iss = plt.subplots(
                    figsize=(13, max(2.6, 0.7 * n_issues))
                )
                bars = ax_iss.barh(
                    plot_df["Issues"] = plot_df["Issues"].apply(clean_text_safe),
                    plot_df["% of Total"],
                    color=plot_colors,
                )
                max_pct = plot_df["% of Total"].max() if n_issues else 1
                for bar, h, pct_total, pct_s9 in zip(
                    bars, plot_df["Hours"],
                    plot_df["% of Total"], plot_df["% of S9"],
                ):
                    ax_iss.text(
                        bar.get_width() + max_pct * 0.01,
                        bar.get_y() + bar.get_height() / 2,
                        f"{pct_total:.1f}% of all project logs  ·  "
                        f"{pct_s9:.1f}% of S9 Work Order  ·  "
                        f"{hours_to_jira_format(h)}",
                        va="center", fontsize=9,
                    )
                ax_iss.set_xlim(0, max_pct * 1.55)
                ax_iss.set_xlabel("% of total logged time")
                ax_iss.set_title(
                    f"Each issue in {s9_project_name} — "
                    f"share of {selected_user}'s total time"
                )
                for spine in ["top", "right"]:
                    ax_iss.spines[spine].set_visible(False)
                fig_iss.tight_layout()
                st.pyplot(fig_iss)

                # --- Detail table ---
                detail = issue_summary[
                    ["S9 Topics", "Logged", "Hours", "% of S9", "% of Total"]
                ].copy()
                detail["Hours"] = detail["Hours"].round(2)
                detail = detail.rename(columns={"Issues": issues_label})

                total_row = pd.DataFrame({
                    issues_label: [f"TOTAL — {s9_project_name}"],
                    "Logged": [hours_to_jira_format(s9_total)],
                    "Hours": [round(s9_total, 2)],
                    "% of S9": [100.0],
                    "% of Total": [round(s9_pct_of_total, 1)],
                })
                detail = pd.concat([detail, total_row], ignore_index=True)

                st.dataframe(
                    detail, hide_index=True, use_container_width=True
                )

                # Sanity check the math for the user
                st.caption(
                    f"✅ Sum of issue *% of Total* = "
                    f"**{issue_summary['% of Total'].sum():.1f}%** "
                    f"= {s9_project_name}'s share of total time."
                )

                # Warn if any issue rows have no name in the export
                if (issue_summary["Issues"] == "(no issue)").any():
                    st.caption(
                        "⚠️ Some S9 - Work Order rows have no issue name in "
                        "the export and were grouped as **(no issue)**."
                    )
            else:
                st.info(
                    "ℹ️ No **S9 - Work Order** entries found for this user."
                )
        else:
            st.info(
                "ℹ️ Add an **Issues** column to your file to see the "
                "S9 - Work Order issue breakdown."
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
