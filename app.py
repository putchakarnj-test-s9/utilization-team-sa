import streamlit as st
import pandas as pd
import re

st.title("Team Utilization Dashboard")

uploaded_file = st.file_uploader("Upload file", type=["txt", "csv"])

def parse_time_to_hours(time_str):
    pattern = r'(\d+)w|(\d+)d|(\d+)h|(\d+)m'
    matches = re.findall(pattern, time_str)

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

if uploaded_file:
    df = pd.read_csv(uploaded_file, sep="\t")
    df = df[df["Project"] != "Total"]
    df["Hours"] = df["Total"].apply(parse_time_to_hours)

    grouped = df.groupby("User").agg({
        "Hours": "sum"
    }).reset_index()

    grouped["Available"] = 160 - grouped["Hours"]
    grouped["Utilization"] = (grouped["Hours"] / 160 * 100).round(0)

    st.dataframe(grouped)

    st.bar_chart(grouped.set_index("User")["Hours"])
