# ---------- S9 WORK ORDER BREAKDOWN ----------
s9_df = user_df[
    user_df["Project"].str.lower().str.contains("s9 - work order")
]

if not s9_df.empty:

    st.subheader("📌 S9 Work Order Breakdown")

    # extract SWO type
    s9_df["SWO Type"] = s9_df["Issues"].str.extract(r"(SWO-\d+)")

    swo_summary = (
        s9_df.groupby("SWO Type")["Hours"]
        .sum()
        .reset_index()
        .sort_values(by="Hours", ascending=False)
    )

    total_swo = swo_summary["Hours"].sum()

    swo_summary["Percentage"] = (
        swo_summary["Hours"] / total_swo * 100
    ).round(1)

    st.dataframe(swo_summary, hide_index=True)

    # ---------- PIE CHART ----------
    fig3, ax3 = plt.subplots(figsize=(5, 5))

    ax3.pie(
        swo_summary["Hours"],
        labels=swo_summary["SWO Type"],
        autopct="%1.1f%%"
    )

    ax3.set_title("S9 Work Order Distribution")

    st.pyplot(fig3)
