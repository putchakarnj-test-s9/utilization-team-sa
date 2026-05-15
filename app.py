# =================================================
# S9 WORK ORDER DETAILS
# =================================================
if "Issues" in user_df.columns:

    try:

        # -----------------------------------------
        # FILTER S9 WORK ORDER
        # -----------------------------------------
        s9_df = user_df[
            user_df["Project"]
            .astype(str)
            .str.lower()
            .str.contains(
                "s9 - work order",
                na=False
            )
        ].copy()

        if len(s9_df) > 0:

            # -----------------------------------------
            # EXTRACT SWO TYPE
            # Example:
            # SWO-2: Leave
            # SWO-4: Training
            # -----------------------------------------
            s9_df["SWO Type"] = (
                s9_df["Issues"]
                .astype(str)
                .str.extract(
                    r"(SWO-\d+)",
                    expand=False
                )
            )

            # remove empty
            s9_df = s9_df[
                s9_df["SWO Type"].notna()
            ]

            if len(s9_df) > 0:

                # -----------------------------------------
                # SUMMARY
                # -----------------------------------------
                swo_summary = (
                    s9_df
                    .groupby("SWO Type")["Hours"]
                    .sum()
                    .reset_index()
                )

                total_swo_hours = (
                    swo_summary["Hours"].sum()
                )

                if total_swo_hours > 0:

                    swo_summary["Percentage"] = (
                        swo_summary["Hours"]
                        / total_swo_hours
                        * 100
                    ).round(0)

                    # -----------------------------------------
                    # TITLE
                    # -----------------------------------------
                    st.subheader(
                        "📌 S9 Work Order Breakdown"
                    )

                    # -----------------------------------------
                    # KPI
                    # -----------------------------------------
                    column_count = len(swo_summary)

                    if column_count <= 0:
                        column_count = 1

                    metric_columns = st.columns(
                        column_count
                    )

                    for i in range(column_count):

                        row = swo_summary.iloc[i]

                        metric_columns[i].metric(
                            row["SWO Type"],
                            f"{row['Percentage']:.0f}%"
                        )

                    # -----------------------------------------
                    # DETAIL TABLE
                    # -----------------------------------------
                    st.subheader(
                        "📝 S9 Work Order Details"
                    )

                    detail_df = s9_df[
                        [
                            "SWO Type",
                            "Issues",
                            "Hours"
                        ]
                    ].copy()

                    detail_df = detail_df.sort_values(
                        by=[
                            "SWO Type",
                            "Hours"
                        ],
                        ascending=[
                            True,
                            False
                        ]
                    )

                    st.dataframe(
                        detail_df,
                        hide_index=True,
                        use_container_width=True
                    )

    except Exception as e:

        st.error(
            "Error loading S9 Work Order details"
        )

        st.exception(e)
