        st.info("No yearly records to compare.")

    section("YEAR-END OUTLOOK")
    o1, o2, o3 = st.columns(3)
    with o1:
        st.metric("Actual net profit to date", fmt_kes(summary["net"]))
    with o2:
        st.metric("Projected year-end net profit", fmt_kes(projected_year_profit))
    with o3:
        st.metric(
            "Projected vs recorded-month target",
            f"{projected_variance_pct:+.1f}%" if summary["target"] else "No target",
        )
    st.caption(
        "The projection uses actual net profit plus the average net profit of recorded months "
        "multiplied by the remaining calendar months. It is an estimate only."
    )

    section("DOWNLOAD YEARLY WORD REPORT")
    report_download_button("yearly_word_report")


# ---------------------------------------------------------------------
# MONTHLY REPORT
# ---------------------------------------------------------------------

elif page == "Monthly Report":
    page_header(
        "MONTHLY REPORT & PRINTABLE SUMMARY",
        f"Full January–December ledger, target assessment and downloadable Word report • {selected_year}",
    )

    section("MONTH-BY-MONTH FINANCIAL REPORT")
    st.dataframe(
        report_table,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Total Kgs": st.column_config.NumberColumn(format="%.2f"),
            "Buying Cost (KES)": st.column_config.NumberColumn(format="KES %.2f"),
            "Selling Value (KES)": st.column_config.NumberColumn(format="KES %.2f"),
            "Gross Profit (KES)": st.column_config.NumberColumn(format="KES %.2f"),
            "Expenses (KES)": st.column_config.NumberColumn(format="KES %.2f"),
            "Net Profit (KES)": st.column_config.NumberColumn(format="KES %.2f"),
            "Net Margin (%)": st.column_config.NumberColumn(format="%.1f%%"),
            "Net Profit Target (KES)": st.column_config.NumberColumn(format="KES %.2f"),
            "Target Variance (%)": st.column_config.NumberColumn(format="%+.1f%%"),
        },
    )

    section("YEAR-TO-DATE TOTALS")
    totals = st.columns(4)
    with totals[0]:
        st.metric("Total KGs", f'{summary["kgs"]:,.2f}')
    with totals[1]:
        st.metric("Buying cost", fmt_kes(summary["buying"]))
    with totals[2]:
        st.metric("Selling value", fmt_kes(summary["selling"]))
    with totals[3]:
        st.metric("Net profit", fmt_kes(summary["net"]))

    section("AUTOMATIC MANAGEMENT COMMENTARY")
    if summary["months"]:
        latest_month = year_data.iloc[-1]["Month"]
        commentary = (
            f"At the end of {latest_month} {selected_year}, "
            f"{summary['months']} month(s) have been recorded. "
            f"The business supplied {summary['kgs']:,.2f} kg of tea leaf, "
            f"with a total selling value of {fmt_kes(summary['selling'])} "
            f"against buying costs of {fmt_kes(summary['buying'])}. "
            f"Gross profit is {fmt_kes(summary['gross'])}. After expenses "
            f"of {fmt_kes(summary['expenses'])}, net profit is "
            f"{fmt_kes(summary['net'])}, giving a net margin of "
            f"{summary['margin']:.1f}%. Average buying price is "
            f"KES {summary['avg_buy']:.2f}/kg and average selling price is "
            f"KES {summary['avg_sell']:.2f}/kg."
        )
        st.write(commentary)
    else:
        st.info("No records are available to create commentary.")

    section("DOWNLOAD REPORTS")
    report_download_button("monthly_page_word_report")

    csv_data = report_table.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "⬇️ DOWNLOAD MONTHLY REPORT AS CSV",
        data=csv_data,
        file_name=f"SL_Tea_Ventures_Monthly_Report_{selected_year}.csv",
        mime="text/csv",
        use_container_width=True,
        key="monthly_csv",
    )

    full_csv = calculated.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "⬇️ DOWNLOAD COMPLETE BUSINESS LEDGER AS CSV",
        data=full_csv,
        file_name="SL_Tea_Ventures_Complete_Ledger.csv",
        mime="text/csv",
        use_container_width=True,
        key="full_ledger_csv",
    )


# ---------------------------------------------------------------------
# FOOTER
# ---------------------------------------------------------------------

st.markdown(
    """
    <div class="footer">
        🍃 <b>S&L TEA VENTURES</b><br>
        Tea Leaf Business Intelligence & Financial Performance Dashboard<br>
        All financial amounts are shown in Kenyan Shillings (KES).
    </div>
    """,
    unsafe_allow_html=True,
)
