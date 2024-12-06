import pandas as pd
import streamlit as st
from datetime import datetime

# Streamlit App Title
st.title("Absence Summary Dashboard")

# Display the counted items table
st.subheader("Counted Absence Types")
counted_items = pd.DataFrame({
    "Absence Type": ['Absent', 'Absent Without Leave', 'Sick Leave', 'Unpaid Medical leave'],
    "Description": [
        "Days marked as absent",
        "Days absent without prior approval",
        "Days taken for medical reasons with a medical certificate",
        "Days taken for medical reasons without pay"
    ]
})
st.table(counted_items)

# File Upload
uploaded_file = st.file_uploader("Upload an Excel File", type=["xlsx"])

if uploaded_file:
    try:
        # Load the uploaded Excel File
        df = pd.ExcelFile(uploaded_file).parse(0)

        # Ensure `Leave From` is in datetime format
        df['Leave From'] = pd.to_datetime(df['Leave From'], errors='coerce')

        # Filter for relevant absence types
        relevant_absences = ['Absent', 'Absent Without Leave', 'Sick Leave', 'Unpaid Medical leave']
        df_filtered = df[df['Absence Type'].isin(relevant_absences)]

        # Extract the month names (Jan, Feb, ...) from `Leave From`
        df_filtered['Month'] = df_filtered['Leave From'].dt.strftime('%b').str.lower()

        # Create a pivot table to calculate absences by month
        monthly_absence = (
            df_filtered.pivot_table(
                index=['Personnel No.', 'Employee Name'],
                columns='Month',
                values='Quota Days',
                aggfunc='sum',
                fill_value=0
            )
        )

        # Rename columns to match full month names
        monthly_absence.columns = [col.capitalize() for col in monthly_absence.columns]

        # Add missing months (if not present) with zero absences
        full_months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
        for month in full_months:
            if month not in monthly_absence.columns:
                monthly_absence[month] = 0

        # Reorder columns to match calendar order
        monthly_absence = monthly_absence[full_months]

        # Calculate Grand Total
        monthly_absence['Grand Total'] = monthly_absence.sum(axis=1)

        # Calculate Last 6 months (Jul–Dec) and Last 3 months (Oct–Dec)
        monthly_absence['Absence Days (Last 6 Months)'] = monthly_absence[['Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']].sum(axis=1)
        monthly_absence['Absence Days (Last 3 Months)'] = monthly_absence[['Oct', 'Nov', 'Dec']].sum(axis=1)

        # Reset index for output
        final_summary = monthly_absence.reset_index()

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
    except KeyError as e:
        st.error(f"Error: Missing expected columns in the file. Details: {e}")
    except Exception as e:
        st.error(f"An unexpected error occurred. Details: {e}")
