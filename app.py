import streamlit as st
import pandas as pd

st.title("SA Utilization Dashboard")

file = st.file_uploader("Upload Worklog Excel", type=["xlsx"])

if file:
    df = pd.read_excel(file)

    # identify date columns (ตัวเลข 1-31)
    date_cols = [col for col in df.columns if str(col).isdigit()]

    # รวมชั่วโมงต่อ row
    df["total_hours"] = df[date_cols].sum(axis=1)

    # รวมต่อ user
    user_summary = df.groupby("User")["total_hours"].sum().reset_index()

    # 🔥 คำนวณ working days จริง
    working_days = len(date_cols)
    capacity = working_days * 8

    user_summary["utilization"] = user_summary["total_hours"] / capacity

    # KPI
    st.metric("Working Days", working_days)
    st.metric("Capacity / User (hrs)", capacity)
    st.metric("Avg Utilization", f"{user_summary['utilization'].mean():.2%}")

    # Chart
    st.subheader("Utilization by User")
    st.bar_chart(user_summary.set_index("User")["utilization"])

    st.subheader("Detail")
    st.dataframe(user_summary)
