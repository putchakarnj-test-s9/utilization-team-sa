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
# worklogs_01.05.2026_31.05.2026.xlsx
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

        )
