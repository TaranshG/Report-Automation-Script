#!/usr/bin/env python3
"""
Report Automation Script

This script automates the generation and distribution of weekly sales reports:
1. Loads raw CSV data containing individual orders.
2. Cleans and preprocesses data to ensure consistency and reliability.
3. Calculates key performance indicators (KPIs) such as total sales, average order value, and product breakdowns.
4. Creates visual summaries (time series and product bar charts) with Matplotlib.
5. Exports metrics and embedded charts into a structured Excel workbook.
6. Simulates email delivery to stakeholders with attachments.
"""

import os
from datetime import datetime
import pandas as pd
import matplotlib.pyplot as plt

# ----------------------
# Configuration Section
# ----------------------
# Base directories for input CSV files and output artifacts (reports & charts)
DATA_DIR = "data"
OUTPUT_DIR = "output"
CHART_DIR = os.path.join(OUTPUT_DIR, "charts")

# Input CSV file path containing raw sales orders data
REPORT_CSV = os.path.join(DATA_DIR, "sales_data.csv")
# Excel report will be named with the current date for versioning
EXCEL_REPORT = os.path.join(
    OUTPUT_DIR,
    f"weekly_report_{datetime.now().strftime('%Y%m%d')}.xlsx"
)

# Simulated email recipients list; in a real implementation this could be dynamic
STAKEHOLDERS = [
    "alice@example.com",
    "bob@example.com",
]

# ----------------------
# Function Definitions
# ----------------------

def load_data(filepath):
    """
    Load raw sales data from a CSV into a Pandas DataFrame.

    :param filepath: Path to the sales_data.csv file
    :return: DataFrame containing raw order records
    """
    # Read all rows and infer column types automatically
    df = pd.read_csv(filepath)
    return df


def clean_data(df):
    """
    Perform data cleaning steps on the raw DataFrame:
    - Convert 'order_date' strings to datetime objects for time series analysis.
    - Drop any rows missing critical information (order ID, date, or sales amount).

    :param df: Raw DataFrame loaded from CSV
    :return: Cleaned DataFrame ready for KPI calculations
    """
    # Convert order_date column from text to datetime for proper resampling later
    df['order_date'] = pd.to_datetime(df['order_date'], errors='coerce')

    # Eliminate records with missing essential fields to avoid skewing results
    df = df.dropna(subset=['order_id', 'order_date', 'sales'])
    return df


def calculate_kpis(df):
    """
    Compute high-level metrics to summarize sales performance:
    - Total sales revenue across all orders.
    - Average sales value per order.
    - Count of unique orders to gauge transaction volume.
    - Breakdown of revenue by product, sorted highest first.

    :param df: Preprocessed DataFrame
    :return: Dictionary of computed KPI values and breakdown series
    """
    total_sales = df['sales'].sum()  # Sum of all transaction amounts
    avg_order_value = df['sales'].mean()  # Mean sales amount per order
    num_orders = df['order_id'].nunique()  # Unique order count

    # Aggregate sales by product to identify top performers
    sales_by_product = df.groupby('product')['sales'] \
                         .sum() \
                         .sort_values(ascending=False)

    return {
        'total_sales': total_sales,
        'avg_order_value': avg_order_value,
        'num_orders': num_orders,
        'sales_by_product': sales_by_product,
    }


def generate_visuals(df, charts_dir):
    """
    Create and save chart images for use in the Excel report:
    - Weekly sales time series (line chart)
    - Top 10 products by total revenue (bar chart)

    :param df: Clean DataFrame with datetime index
    :param charts_dir: Directory path to save generated charts
    :return: List of saved image file paths
    """
    # Ensure the output directory for charts exists
    os.makedirs(charts_dir, exist_ok=True)
    chart_paths = []

    # Weekly sales over time: group orders by calendar week and sum sales
    weekly_sales = df.set_index('order_date') \
                    .resample('W')['sales'] \
                    .sum()
    plt.figure()
    weekly_sales.plot(
        title='Weekly Sales Over Time',
        xlabel='Week Ending',
        ylabel='Total Sales ($)'
    )
    timeseries_path = os.path.join(charts_dir, 'weekly_sales.png')
    plt.savefig(timeseries_path, bbox_inches='tight')
    plt.close()
    chart_paths.append(timeseries_path)

    # Identify and plot the top 10 revenue-generating products
    top_products = df.groupby('product')['sales'] \
                     .sum() \
                     .nlargest(10)
    plt.figure()
    top_products.plot(
        kind='bar',
        title='Top 10 Products by Revenue',
        xlabel='Product Name',
        ylabel='Revenue ($)'
    )
    top_products_path = os.path.join(charts_dir, 'top_products.png')
    plt.savefig(top_products_path, bbox_inches='tight')
    plt.close()
    chart_paths.append(top_products_path)

    return chart_paths


def export_to_excel(kpis, chart_files, output_file):
    """
    Compile KPIs and embed visual summaries into an Excel workbook.

    - Creates a 'Summary' sheet with key metrics.
    - Inserts chart images below the metrics for visual context.

    :param kpis: Dictionary of KPI values from calculate_kpis()
    :param chart_files: List of chart image paths
    :param output_file: Destination path for the Excel .xlsx report
    """
    # Ensure the directory structure for the report exists
    os.makedirs(os.path.dirname(output_file), exist_ok=True)

    # Use XlsxWriter engine to support image embedding
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        # Prepare a DataFrame summarizing core metrics
        summary_df = pd.DataFrame({
            'Metric': ['Total Sales', 'Average Order Value', 'Number of Orders'],
            'Value': [
                kpis['total_sales'],
                kpis['avg_order_value'],
                kpis['num_orders']
            ]
        })
        summary_df.to_excel(writer, sheet_name='Summary', index=False)

        workbook  = writer.book
        worksheet = writer.sheets['Summary']

        # Starting row for chart insertion (leaving space after the metrics table)
        start_row = len(summary_df) + 3
        for path in chart_files:
            # Insert each chart with appropriate scaling
            worksheet.insert_image(
                f'A{start_row}', path,
                {'x_scale': 0.8, 'y_scale': 0.8}
            )
            start_row += 20  # Move down to avoid overlapping charts

    print(f"Excel report successfully created at: {output_file}")


def simulate_email(recipients, subject, body, attachments=None):
    """
    Print a simulated email to the console to emulate distribution:
    - Shows 'To', 'Subject', and message body.
    - Lists any file attachments for stakeholder visibility.

    :param recipients: List of email addresses for notification
    :param subject: Email subject line text
    :param body: Body text of the email
    :param attachments: Optional list of file paths to attach
    """
    print("\n--- Simulated Email Delivery ---")
    print(f"To: {', '.join(recipients)}")
    print(f"Subject: {subject}\n")
    print(body)
    if attachments:
        print("\nAttachments:")
        for f in attachments:
            print(f" - {f}")
    print("--- End of Email ---\n")


def main():
    """
    Orchestrate the end-to-end report generation workflow:
    1. Load raw sales data
    2. Clean and validate the dataset
    3. Calculate and retrieve KPIs
    4. Generate chart images for insights
    5. Export a consolidated Excel report
    6. Simulate notifying stakeholders via console
    """
    # Step 1: Read data from CSV
    raw_data = load_data(REPORT_CSV)

    # Step 2: Clean and prepare data for analysis
    clean_df = clean_data(raw_data)

    # Step 3: Compute key performance indicators
    kpis = calculate_kpis(clean_df)

    # Step 4: Create visual summaries and save to disk
    chart_files = generate_visuals(clean_df, CHART_DIR)

    # Step 5: Build and save Excel report
    export_to_excel(kpis, chart_files, EXCEL_REPORT)

    # Step 6: Simulated email notification to stakeholders
    subject = f"Weekly Sales Report - {datetime.now().strftime('%Y-%m-%d')}"
    email_body = (
        "Hello Team,\n\n"
        "Attached is this week’s sales report with key insights and charts. "
        "Please review and let me know if you have any questions.\n\n"
        "Best regards,\nAutomation Bot"
    )
    simulate_email(
        STAKEHOLDERS,
        subject,
        email_body,
        attachments=[EXCEL_REPORT] + chart_files
    )


if __name__ == "__main__":
    main()
