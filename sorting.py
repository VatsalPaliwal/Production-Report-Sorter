if uploaded_file is not None:

    with st.spinner("Sorting report... Please wait."):

        sorted_df = process_sheet(uploaded_file)

        output = BytesIO()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            sorted_df.to_excel(
                writer,
                sheet_name="Sorted Data",
                startrow=4,
                index=False
            )

        output.seek(0)

    st.success("Sorting completed successfully.")

    st.dataframe(
        sorted_df,
        use_container_width=True,
        hide_index=True
    )

    st.download_button(
        label="Download Sorted Report",
        data=output,
        file_name="Sorted_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
