import pandas as pd
import streamlit as st

# Streamlit App Title
st.title("Absence Summary Dashboard")

# File Upload
uploaded_file = st.file_uploader("Upload an Excel File", type=["xlsx"])

if uploaded_file:
    # Load the uploaded Excel File
    df = pd.ExcelFile(uploaded_file).parse(0)

    # Debug: Display column names
    st.write("Columns in the uploaded file:")
    st.write(df.columns.tolist())

    # Ensure monthly columns are correctly identified (expected columns 1–12)
    try:
        # Dynamically detect monthly columns (1–12) by their positions or expected names
        monthly_columns = df.columns[3:15]  # Adjust based on dataset structure
        monthly_columns = [col for col in monthly_columns if col in range(1, 13) or str(col).isdigit()]

        # Convert monthly columns to numeric
        df[monthly_columns] = df[monthly_columns].apply(pd.to_numeric, errors='coerce').fillna(0)

        # Calculate Grand Total from monthly columns
        df['Grand Total'] = df[monthly_columns].sum(axis=1)

        # Calculate Last 6 months (July–December) and Last 3 months (October–December)
        last_6_months = [col for col in monthly_columns if col in [7, 8, 9, 10, 11, 12]]
        last_3_months = [col for col in monthly_columns if col in [10, 11, 12]]

        df['Absence Days (Last 6 Months)'] = df[last_6_months].sum(axis=1)
        df['Absence Days (Last 3 Months)'] = df[last_3_months].sum(axis=1)

        # Display Data in Streamlit
        st.subheader("Absence Summary Table")
        st.dataframe(df)

        # Downloadable Excel Output
        output_file = "Absence_Summary.xlsx"
        df.to_excel(output_file, index=False)

        with open(output_file, "rb") as file:
            st.download_button(
                label="Download Absence Summary",
                data=file,
                file_name="Absence_Summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    except Exception as e:
        st.error(f"Error processing the file: {e}")
        st.write("Please ensure the uploaded file contains valid monthly columns (1–12).")
