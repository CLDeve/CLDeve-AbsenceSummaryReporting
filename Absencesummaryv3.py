import pandas as pd
import streamlit as st

# Streamlit App Title
st.title("Absence Summary Dashboard")

# File Upload
uploaded_file = st.file_uploader("Upload an Excel File", type=["xlsx"])

if uploaded_file:
    try:
        # Load the uploaded Excel File
        df = pd.ExcelFile(uploaded_file).parse(0)

        # Identify monthly columns dynamically
        monthly_columns = list(range(1, 13))  # Expected month columns (1–12)

        # Ensure monthly columns are numeric
        df[monthly_columns] = df[monthly_columns].apply(pd.to_numeric, errors='coerce').fillna(0)

        # Calculate Grand Total from monthly columns
        df['Grand Total'] = df[monthly_columns].sum(axis=1)

        # Calculate Last 6 months (July–December) and Last 3 months (October–December)
        last_6_months = [7, 8, 9, 10, 11, 12]
        last_3_months = [10, 11, 12]

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
    except KeyError as e:
        st.error(f"Error: Missing expected columns in the file. Details: {e}")
    except Exception as e:
        st.error(f"An unexpected error occurred. Details: {e}")
