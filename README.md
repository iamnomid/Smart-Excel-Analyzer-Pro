# Smart Excel Analyzer Pro

A Python-based data cleaning, analysis, and reporting application built with Streamlit.

## Project Overview

Smart Excel Analyzer Pro helps users explore Excel and CSV datasets, identify common data quality issues, perform optional cleaning operations, and generate downloadable reports.

This project is a portfolio demonstration of Python data analysis and reporting automation.
## Live Demo

Try the application online:

https://smart-excel-analyzer-pro.onrender.com

## GitHub Repository

https://github.com/iamnomid/Smart-Excel-Analyzer-Pro

## How to Test

1. Open the live demo.
2. Upload a sample Excel or CSV file.
3. Enable the data cleaning options.
4. Explore the sales dashboard.
5. Download the cleaned Excel file and PDF report.
## Features

- Upload Excel (.xlsx) and CSV files
- Preview and explore datasets
- Identify duplicate rows and missing values
- Optionally remove duplicate rows
- Optionally fill missing numerical values using the median
- Generate product sales summaries and charts
- Export cleaned datasets to Excel
- Generate PDF reports with tables and charts

## Technologies

Python, Pandas, Streamlit, OpenPyXL, Matplotlib, and ReportLab.

## Installation

Clone or download this repository, open its folder in a terminal, and create a Python virtual environment.

Install the dependencies using:

`python -m pip install -r requirements.txt`

Run the application using:

`python -m streamlit run app.py`

Open the local URL shown in the terminal.

## Sample Datasets

- `test_sales.csv`: Sample product sales data for testing cleaning and sales analysis.
- `test_customers.csv`: Sample customer data for testing general data quality reporting.

## Notes and Limitations

The sales dashboard requires the columns Product, Price, and Quantity.

Duplicate removal compares complete rows. Missing-value replacement uses simple default rules, which may not be appropriate for every dataset.

The application is a demonstration project and should be tested and adapted before processing production business data.

## Future Improvements

- Advanced data validation
- Configurable cleaning rules
- Automated tests
- Enhanced reporting and visualization
- Multi-file processing

## Project Purpose

Developed as a practical Python data analytics and automation portfolio project.