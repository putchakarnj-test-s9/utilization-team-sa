import streamlit as st
import pandas as pd
import re
import os
import urllib.request
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
from matplotlib.patches import Patch
from datetime import datetime


# ---------- THAI FONT SETUP ----------
# Matplotlib only renders glyphs the active font actually contains. Naming a
# font in rcParams isn't enough — it has to exist on the host. This function
# guarantees a Thai-capable font is registered (bundled, system, or
# auto-downloaded), so charts never show boxes for Thai characters.

_THAI_CODEPOINTS = list(range(0x0E01, 0x0E5C))   # Thai Unicode block
_MIN_THAI_GLYPHS = 30                            # threshold for "real" Thai font


def _font_has_thai(font_path):
    """Return True if the .ttf/.otf at font_path covers Thai."""
    try:
        from fontTools.ttLib import TTFont
        tt = TTFont(font_path, fontNumber=0, lazy=True)
        cmap = tt.getBestCmap()
        tt.close()
        if not cmap:
            return False
        return sum(1 for cp in _THAI_CODEPOINTS if cp in cmap) >= _MIN_THAI_GLYPHS
    except Exception:
        return False


def setup_thai_font():
    """Find or install a Thai-capable font. Returns the font name set, or None."""
    here = os.path.dirname(os.path.abspath(__file__))

    # 1) Look in ./fonts/ next to app.py (drop a .ttf there to bundle a font)
    bundled_dir = os.path.join(here, "fonts")
    if os.path.isdir(bundled_dir):
        for fname in sorted(os.listdir(bundled_dir)):
            if fname.lower().endswith((".ttf", ".otf")):
                path = os.path.join(bundled_dir, fname)
                if _font_has_thai(path):
                    fm.fontManager.addfont(path)
                    name = fm.FontProperties(fname=path).get_name()
                    plt.rcParams["font.family"] = [name]
                    plt.rcParams["axes.unicode_minus"] = False
                    return name

    # 2) Scan installed system fonts for actual Thai glyph coverage
    for fpath in fm.findSystemFonts():
        if _font_has_thai(fpath):
            fm.fontManager.addfont(fpath)
            name = fm.FontProperties(fname=fpath).get_name()
            plt.rcParams["font.family"] = [name]
            plt.rcParams["axes.unicode_minus"] = False
            return name

    # 3) Last resort: download Sarabun (popular Thai font) from Google Fonts
    cache_dir = os.path.join(here, ".thai_font_cache")
    os.makedirs(cache_dir, exist_ok=True)
    font_path = os.path.join(cache_dir, "Sarabun-Regular.ttf")

    if not (os.path.exists(font_path) and _font_has_thai(font_path)):
        urls = [
            "https://github.com/google/fonts/raw/main/ofl/sarabun/Sarabun-Regular.ttf",
            "https://raw.githubusercontent.com/google/fonts/main/ofl/sarabun/Sarabun-Regular.ttf",
            "https://github.com/google/fonts/raw/main/ofl/notosansthai/NotoSansThai-Regular.ttf",
        ]
        for url in urls:
            try:
                urllib.request.urlretrieve(url, font_path)
                if _font_has_thai(font_path):
                    break
            except Exception:
                continue

    if os.path.exists(font_path) and _font_has_thai(font_path):
        fm.fontManager.addfont(font_path)
        name = fm.FontProperties(fname=font_path).get_name()
        plt.rcParams["font.family"] = [name]
        plt.rcParams["axes.unicode_minus"] = False
        return name

    return None


_THAI_FONT_NAME = setup_thai_font()


# ---------- PAGE CONFIG ----------
st.set_page_config(page_title="🖥 Team Utilization Dashboard", layout="wide")

# ---------- CATEGORY CONFIG ----------
# Each issue is classified by checking its title against these keyword rules,
# IN ORDER (first match wins). Edit / reorder these to tune the classification.
CATEGORY_RULES = [
    ("external", "External Meeting"),       # e.g. "External Client Activity"
    ("internal meeting", "Internal Meeting"),
    ("ลาพักร้อน", "Leave"),                   # Thai: leave
    ("ลาทุกชนิด", "Holiday"),                  # Thai: any leave
    ("leave", "Leave"),                     # e.g. "Leave (ลาพักร้อน/ลาทุกชนิด)"
    ("training", "Training"),               # e.g. "Training / Meeting / กิจกรรมบริษัท"
    ("meeting", "Meeting"),                 # any other meeting
    ("กิจกรรมบริษัท", "Company Activities"),    # Thai: company activity
    ("R & D", "R & D"),
]
DEFAULT_CATEGORY = "Task"

CATEGORY_COLORS = {
    "Task": "#22c55e",              # green
    "Internal Meeting": "#3b82f6",   # blue
    "External Meeting": "#f59e0b",   # amber
    "Leave": "#94a3b8",             # gray
}


# ---------- PROJECT REASSIGNMENT (FIX MISFILED ISSUES) ----------
# Some issues get logged under "S9 - Work Order" in Tempo even though they
# really belong to a separate project (data-entry mistake). Each rule moves
# those rows to a target project so the project chart and S9 breakdown both
# reflect reality. Match is case-insensitive substring on the issue name.
S9_REASSIGN = [
    # (keyword in Issue name, target Project name)
    ("GGP(PH2)", "GGP - Phase 2"),
    # add more as you find them, e.g.:
    # ("AWC", "AWC : Pikul"),
    # ("Win Performance", "Win Performance"),
]


# ---------- PROJECT COLORS ----------
# Reserved colours for the two S9 projects; everything else cycles through
# the palette below so each non-S9 project shows in a distinct colour.
PROJECT_PALETTE = [
    "#22c55e",  # green
    "#a855f7",  # purple
    "#f59e0b",  # amber
    "#0ea5e9",  # sky
    "#ec4899",  # pink
    "#14b8a6",  # teal
    "#84cc16",  # lime
    "#f97316",  # orange
    "#06b6d4",  # cyan
    "#eab308",  # yellow
    "#8b5cf6",  # violet
    "#10b981",  # emerald
]


def get_project_colors(projects):
    """Return one colour per project, in the same order as the input list."""
    colors = []
    palette_idx = 0
    for p in projects:
        pl = str(p).lower()
        if "s9 - work order" in pl:
            colors.append("#ef4444")  # red, reserved
        elif "s9 - tech support" in pl:
            colors.append("#3b82f6")  # blue, reserved
        else:
            colors.append(PROJECT_PALETTE[palette_idx % len(PROJECT_PALETTE)])
            palette_idx += 1
    return colors


def categorize_issue(issue):
    """Classify a single issue title into a category."""
    s = str(issue).lower()
    for keyword, category in CATEGORY_RULES:
        if keyword.lower() in s:
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

    if isinstance(time_str, (int, float)):
        return float(time_str)

    s = str(time_str).strip()
    if s == "":
        return 0

    try:
        return float(s)
    except ValueError:
        pass

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

                cols_ci = {c.lower(): c for c in df.columns}
                rename = {}
                if "time spent (hours)" in cols_ci:
                    rename[cols_ci["time spent (hours)"]] = "Total"
                df = df.rename(columns=rename)

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

    df["User"] = df["User"].ffill()
    df["Project"] = df["Project"].ffill()

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
        df["Issues"] = df["Issues"].replace(
            {"": "(no issue)", "nan": "(no issue)",
             "None": "(no issue)", "<NA>": "(no issue)"}
        )
        df.attrs["issues_source_column"] = issues_source

        # Reassign misfiled S9 issues to their real project (S9_REASSIGN)
        if S9_REASSIGN:
            proj_lower = df["Project"].astype(str).str.lower()
            issue_lower = df["Issues"].astype(str).str.lower()
            for keyword, target_project in S9_REASSIGN:
                mask = (
                    proj_lower.str.contains("s9 - work order",
                                            na=False, regex=False)
                    & issue_lower.str.contains(keyword.lower(),
                                               na=False, regex=False)
                )
                if mask.any():
                    df.loc[mask, "Project"] = target_project
                    proj_lower = df["Project"].astype(str).str.lower()
    else:
        df.attrs["all_columns"] = list(df.columns)

    df = df[df["Project"].astype(str).str.strip().str.lower() != "total"]
    df = df[df["User"].astype(str).str.strip().str.lower() != "summary"]
    if has_issues:
        df = df[df["Issues"].str.lower() != "total"]

    df = df[df["Total"].notna()]
    df["Hours"] = df["Total"].apply(parse_time_to_hours)
    df = df[df["Hours"] > 0]

    if has_issues:
        df["Category"] = df["Issues"].apply(categorize_issue)

    return df


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

        if not has_issues:
            all_cols = df.attrs.get("all_columns", list(df.columns))
            st.warning(
                "⚠️ No issue-name column was detected. "
                f"Columns in your file: {', '.join(all_cols)}"
            )

        # Heads-up if no Thai font could be loaded -> Thai will render as boxes
        if _THAI_FONT_NAME is None:
            st.warning(
                "⚠️ No Thai font available — Thai characters will show as "
                "boxes in charts. Install a Thai font (Sarabun, TH Sarabun "
                "New, Noto Sans Thai, Tahoma) or drop a `.ttf` file into a "
                "`fonts/` folder next to `app.py`."
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
        colors = get_project_colors([p for p, _ in segments])
        fig = stacked_100_bar(segments, colors, figsize=(12, 2.0))
        st.pyplot(fig)

        # ---------- S9 - WORK ORDER FOCUSED DRILL-DOWN ----------
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

                n_issues = len(issue_summary)

                def _issue_shade(i, n):
                    if n <= 1:
                        return "#1e3a8a"
                    ratio = i / (n - 1)
                    r = int(30 + (191 - 30) * ratio)
                    g = int(58 + (219 - 58) * ratio)
                    b = int(138 + (254 - 138) * ratio)
                    return f"#{r:02x}{g:02x}{b:02x}"

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
                    plot_df["Issues"],
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
                    ["Issues", "Logged", "Hours", "% of S9", "% of Total"]
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

                st.caption(
                    f"✅ Sum of issue *% of Total* = "
                    f"**{issue_summary['% of Total'].sum():.1f}%** "
                    f"= {s9_project_name}'s share of total time."
                )

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
