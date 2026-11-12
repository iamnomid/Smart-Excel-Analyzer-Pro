
import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Excel Merger Pro")

st.title("Multi-file Excel & CSV Merger - V2")

files = st.file_uploader(
    "Upload two or more files",
    type=["csv", "xlsx"],
    accept_multiple_files=True
)

if len(files) >= 2:
    try:
        frames = []

        for file in files:
            if file.name.lower().endswith(".csv"):
                df = pd.read_csv(file)
            else:
                df = pd.read_excel(file)

            df.columns = df.columns.astype(str).str.strip()

            if df.columns.duplicated().any():
                raise ValueError(
                    f"Duplicate column names in {file.name}"
                )

            frames.append((file.name, df))

        columns = list(frames[0][1].columns)

        if "Source_File" in columns or "Source_Row" in columns:
            raise ValueError(
                "Source_File and Source_Row are reserved names."
            )

        for name, df in frames:
            if set(df.columns) != set(columns):
                raise ValueError(
                    f"Column mismatch in {name}"
                )

        parts = []

        for name, df in frames:
            part = df.reindex(columns=columns).copy()
            part["Source_File"] = name
            part["Source_Row"] = range(2, len(df) + 2)
            parts.append(part)

        combined = pd.concat(parts, ignore_index=True)

        potential_duplicates = combined.duplicated(
            subset=columns,
            keep=False
        )

        rows_to_remove = combined.duplicated(
            subset=columns,
            keep="first"
        )

        st.subheader("Data Quality Report")

        st.metric("Original Rows", len(combined))
        st.metric(
            "Potential Duplicate Rows",
            int(potential_duplicates.sum())
        )

        remove_duplicates = st.checkbox(
            "Remove duplicate rows (keep first)"
        )

        if remove_duplicates:
            cleaned = combined.loc[~rows_to_remove].copy()
            change_log = combined.loc[rows_to_remove].copy()
            change_log["Action"] = "Removed exact duplicate"
        else:
            cleaned = combined.copy()
            change_log = pd.DataFrame(
                columns=list(combined.columns) + ["Action"]
            )

        st.subheader("Cleaning Results")

        col1, col2 = st.columns(2)

        col1.metric(
            "Rows Removed",
            len(combined) - len(cleaned)
        )

        col2.metric("Final Cleaned Rows", len(cleaned))

        st.subheader("Cleaned Data")
        st.dataframe(cleaned)

        st.subheader("Change Log")
        st.dataframe(change_log)

        summary = pd.DataFrame({
            "Metric": [
                "Original Rows",
                "Potential Duplicate Rows",
                "Rows Removed",
                "Final Cleaned Rows"
            ],
            "Value": [
                len(combined),
                int(potential_duplicates.sum()),
                len(change_log),
                len(cleaned)
            ]
        })

        output = BytesIO()

        with pd.ExcelWriter(
            output,
            engine="openpyxl"
        ) as writer:

            combined.to_excel(
                writer, sheet_name="All Data", index=False
            )

            cleaned.to_excel(
                writer, sheet_name="Cleaned Data", index=False
            )

            combined.loc[potential_duplicates].to_excel(
                writer,
                sheet_name="Potential Duplicates",
                index=False
            )

            change_log.to_excel(
                writer, sheet_name="Change Log", index=False
            )

            summary.to_excel(
                writer, sheet_name="Summary", index=False
            )

        for sheet in writer.book.worksheets:
            sheet.freeze_panes = "A2"
            sheet.auto_filter.ref = sheet.dimensions

            for column in sheet.columns:
                max_length = max(
                    len(str(cell.value or ""))
                    for cell in column
                )

                column_letter = column[0].column_letter

                sheet.column_dimensions[column_letter].width = min(
                    max(max_length + 3, 14),
                    45
                )
        st.download_button(
            "Download Professional Excel Report",
            data=output.getvalue(),
            file_name="merged_data_v2.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            )
        )

    except Exception as error:
        st.error(f"Error: {error}")

else:
    st.info("Please upload at least two files.")
