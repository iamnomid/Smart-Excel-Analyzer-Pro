import matplotlib.pyplot as plt
from reportlab.platypus import Image as PDFImage


from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle
)

import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(
    page_title="Smart Excel Analyzer Pro",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Smart Excel Analyzer Pro")

st.write(
    "Upload, analyze, clean and export your data."
)

uploaded_file = st.file_uploader(
    "Upload an Excel or CSV file",
    type=["xlsx", "csv"]
)

if uploaded_file is not None:

    try:

        # Read the uploaded file
        if uploaded_file.name.lower().endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)

        st.success("File uploaded successfully!")

        # Original data preview
        st.subheader("1. Original Data")

        st.dataframe(df.head(20))

        # Data quality metrics
        st.subheader("2. Data Quality Report")

        col1, col2, col3 = st.columns(3)

        col1.metric("Total Rows", len(df))

        col2.metric("Duplicate Rows", df.duplicated().sum())

        col3.metric(
            "Missing Values",
            df.isnull().sum().sum()
        )

        # Missing values by column
        st.subheader("3. Missing Values by Column")

        missing_report = pd.DataFrame({
            "Column": df.columns,
            "Missing Values": df.isnull().sum().values
        })

        st.dataframe(missing_report)

        # Cleaning options
        st.subheader("4. Data Cleaning Options")

        remove_duplicates = st.checkbox(
            "Remove duplicate rows"
        )

        fill_missing = st.checkbox(
            "Fill missing values"
        )

        cleaned_df = df.copy()

        # Remove duplicates
        if remove_duplicates:

            cleaned_df = cleaned_df.drop_duplicates()

        # Fill missing values
        if fill_missing:

            for column in cleaned_df.columns:

                if pd.api.types.is_numeric_dtype(
                    cleaned_df[column]
                ):

                    median_value = cleaned_df[column].median()

                    if pd.notna(median_value):

                        cleaned_df[column] = (
                            cleaned_df[column].fillna(
                                median_value
                            )
                        )

                elif (
                    pd.api.types.is_object_dtype(
                        cleaned_df[column]
                    )
                    or pd.api.types.is_string_dtype(
                        cleaned_df[column]
                    )
                ):

                    cleaned_df[column] = (
                        cleaned_df[column].fillna(
                            "Unknown"
                        )
                    )

        # Cleaning results
        st.subheader("5. Cleaning Results")

        col4, col5, col6 = st.columns(3)

        col4.metric(
            "Original Rows",
            len(df)
        )

        col5.metric(
            "Cleaned Rows",
            len(cleaned_df)
        )

        col6.metric(
            "Remaining Missing Values",
            int(cleaned_df.isnull().sum().sum())
        )

        st.subheader("6. Cleaned Data Preview")

        st.dataframe(cleaned_df.head(20))

        # Sales Dashboard
        st.subheader("7. Sales Dashboard")

        required = {"Product", "Price", "Quantity"}

        if required.issubset(cleaned_df.columns):

            sales_df = cleaned_df[
                ["Product", "Price", "Quantity"]
            ].copy()

            sales_df["Price"] = pd.to_numeric(
                sales_df["Price"], errors="coerce"
            )

            sales_df["Quantity"] = pd.to_numeric(
                sales_df["Quantity"], errors="coerce"
            )

            sales_df = sales_df.dropna()

            sales_df["Sales Value"] = (
                sales_df["Price"] * sales_df["Quantity"]
            )

            summary = sales_df.groupby(
                "Product", as_index=False
            )["Sales Value"].sum()

            total = summary["Sales Value"].sum()

            st.metric(
                "Total Sales Value",
                f"{total:,.2f}"
            )

            st.subheader("Sales Value by Product")

            st.bar_chart(
                summary.set_index("Product")["Sales Value"]
            )

        else:
            st.info(
                "Sales dashboard requires Product, "
                "Price and Quantity columns."
            )



        # Export Excel
        output = BytesIO()

        with pd.ExcelWriter(
            output,
            engine="openpyxl"
        ) as writer:

            cleaned_df.to_excel(
                writer,
                index=False,
                sheet_name="Cleaned Data"
            )

        st.download_button(
            label="Download Cleaned Excel",
            data=output.getvalue(),
            file_name="cleaned_data.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            )
        )

                # PDF Report
        st.subheader("8. PDF Report")

        pdf_buffer = BytesIO()

        doc = SimpleDocTemplate(
            pdf_buffer,
            pagesize=A4
        )

        styles = getSampleStyleSheet()
        elements = []

        elements.append(
            Paragraph(
                "Smart Excel Analyzer - Data Report",
                styles["Title"]
            )
        )

        elements.append(Spacer(1, 20))

        report_data = [
            ["Metric", "Value"],
            ["Original Rows", str(len(df))],
            ["Cleaned Rows", str(len(cleaned_df))],
            [
                "Remaining Missing Values",
                str(int(cleaned_df.isna().sum().sum()))
            ]
        ]

        table = Table(report_data)

        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.lightblue),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
            ("LEFTPADDING", (0, 0), (-1, -1), 8),
            ("RIGHTPADDING", (0, 0), (-1, -1), 8)
        ]))

        elements.append(table)
                # Add sales summary to PDF
        if required.issubset(cleaned_df.columns):

            elements.append(Spacer(1, 20))

            elements.append(
                Paragraph(
                    "Sales Summary",
                    styles["Heading2"]
                )
            )

            elements.append(
                Paragraph(
                    f"Total Sales Value: {total:,.2f}",
                    styles["Normal"]
                )
            )

            sales_data = [
                ["Product", "Sales Value"]
            ]

            for _, row in summary.head(10).iterrows():

                sales_data.append([
                    str(row["Product"]),
                    f'{row["Sales Value"]:,.2f}'
                ])

            sales_table = Table(
                sales_data,
                colWidths=[240, 140],
                repeatRows=1
            )

            sales_table.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.lightblue),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("TOPPADDING", (0, 0), (-1, -1), 8),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 8)
            ]))

            elements.append(Spacer(1, 10))
            elements.append(sales_table)

                        # Create sales chart for PDF
            chart_data = summary.sort_values(
                "Sales Value",
                ascending=False
            ).head(10)

            fig, ax = plt.subplots(figsize=(7, 3.5))

            ax.bar(
                chart_data["Product"].astype(str),
                chart_data["Sales Value"]
            )

            ax.set_title("Sales Value by Product")
            ax.set_ylabel("Sales Value")
            ax.tick_params(axis="x", rotation=30)

            fig.tight_layout()

            chart_buffer = BytesIO()

            fig.savefig(
                chart_buffer,
                format="png",
                dpi=150
            )

            plt.close(fig)
            chart_buffer.seek(0)

            elements.append(Spacer(1, 20))

            elements.append(
                Paragraph("Sales Chart", styles["Heading2"])
            )

            elements.append(
                PDFImage(
                    chart_buffer,
                    width=440,
                    height=220
                )
            )


        
        doc.build(elements)

        st.download_button(
            label="Download PDF Report",
            data=pdf_buffer.getvalue(),
            file_name="data_report.pdf",
            mime="application/pdf"
        )



    except Exception as error:

        st.error(f"Error: {error}")