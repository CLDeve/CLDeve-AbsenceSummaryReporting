import pandas as pd
import streamlit as st
from datetime import datetime, timedelta

# Streamlit App Title
st.title("Absence Summary Dashboard")

# File Upload
uploaded_file = st.file_uploader("Upload an Excel File", type=["xlsx"])

if uploaded_file:
    # Load the Excel File
    df = pd.ExcelFile(uploaded_file).parse(0)

    # Focus on relevant columns
    df_filtered = df[['Personnel No.', 'Employee Name', 'Org.Unit', 'Leave From', 'Quota Days', 'Absence Type']]
    df_filtered.columns = ['Staff ID', 'Name', 'Org Unit', 'Leave Date', 'Quota Days', 'Absence Type']

    # Convert 'Staff ID' to integers
    df_filtered['Staff ID'] = pd.to_numeric(df_filtered['Staff ID'], errors='coerce').fillna(0).astype(int)

    # Convert 'Leave Date' to datetime
    df_filtered['Leave Date'] = pd.to_datetime(df_filtered['Leave Date'], errors='coerce')

    # Filter for relevant absence types
    relevant_absences = ['Absent', 'Absent Without Leave', 'Sick Leave', 'Unpaid Medical leave']
    df_filtered = df_filtered[df_filtered['Absence Type'].isin(relevant_absences)]

    # Expand absence days based on 'Leave From' and 'Quota Days'
    expanded_rows = []
    for _, row in df_filtered.iterrows():
        start_date = row['Leave Date']
        quota_days = int(row['Quota Days'])
        
        for day in range(quota_days):
            current_date = start_date + timedelta(days=day)
            expanded_rows.append({
                'Staff ID': row['Staff ID'],
                'Name': row['Name'],
                'Org Unit': row['Org Unit'],
                'Date': current_date,
                'Month': current_date.month,
                'Absence Days': 1  # Each day counts as 1 absence day
            })

    # Create expanded DataFrame
    df_expanded = pd.DataFrame(expanded_rows)

    # Group by Staff ID, Name, Org Unit, and numeric month
    monthly_absence = df_expanded.groupby(['Staff ID', 'Name', 'Org Unit', 'Month'])['Absence Days'].sum().reset_index()

    # Pivot table to format months as columns
    monthly_absence_pivot = monthly_absence.pivot(index=['Staff ID', 'Name', 'Org Unit'], columns='Month', values='Absence Days').fillna(0)

    # Calculate totals
    monthly_absence_pivot['Grand Total'] = monthly_absence_pivot.sum(axis=1)

    # Filter data for the last 3 and 6 months based on actual dates
    today = datetime.now()
    last_3_months = today - timedelta(days=90)
    last_6_months = today - timedelta(days=180)

    recent_3_months = df_expanded[df_expanded['Date'] >= last_3_months]
    recent_6_months = df_expanded[df_expanded['Date'] >= last_6_months]

    summary_3_months = recent_3_months.groupby(['Staff ID', 'Name', 'Org Unit'])['Absence Days'].sum().reset_index(name='Absence Days (Last 3 Months)')
    summary_6_months = recent_6_months.groupby(['Staff ID', 'Name', 'Org Unit'])['Absence Days'].sum().reset_index(name='Absence Days (Last 6 Months)')

    # Merge the summaries with the pivoted data
    final_summary = pd.merge(monthly_absence_pivot, summary_3_months, on=['Staff ID', 'Name', 'Org Unit'], how='left')
    final_summary = pd.merge(final_summary, summary_6_months, on=['Staff ID', 'Name', 'Org Unit'], how='left')

    # Fill NaN values with 0
    final_summary.fillna(0, inplace=True)

    # Reset column order to match desired format
    final_summary = final_summary.reset_index()
    column_order = ['Staff ID', 'Name', 'Grand Total'] + list(range(1, 13)) + ['Absence Days (Last 3 Months)', 'Absence Days (Last 6 Months)']
    for month in range(1, 13):
        if month not in final_summary.columns:
            final_summary[month] = 0  # Add missing month columns if they don't exist
    final_summary = final_summary[column_order]

    # Display Data in Streamlit
    st.subheader("Absence Summary Table")
    st.dataframe(final_summary)

    # Downloadable Excel Output
    output_file = "Absence_Summary.xlsx"
    final_summary.to_excel(output_file, index=False)

    with open(output_file, "rb") as file:
        st.download_button(
            label="Download Absence Summary",
            data=file,
            file_name="Absence_Summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
