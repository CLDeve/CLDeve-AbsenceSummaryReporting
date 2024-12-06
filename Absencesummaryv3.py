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

        # Clean and normalize column names
        df.columns = df.columns.str.strip().str.replace(" ", "_").str.lower()

        # Define monthly columns as Jan to Dec
        monthly_columns = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec']
        if not all(col in df.columns for col in monthly_columns):
            raise KeyError(f"Expected monthly columns {monthly_columns} are missing or misnamed in the uploaded file.")

        # Ensure monthly columns are numeric
        df[monthly_columns] = df[monthly_columns].apply(pd.to_numeric, errors='coerce').fillna(0)

        # Calculate Grand Total from monthly columns
        df['Grand_Total'] = df[monthly_columns].sum(axis=1)

        # Calculate Last 6 months (Jul–Dec) and Last 3 months (Oct–Dec)
        last_6_months = ['jul', 'aug', 'sep', 'oct', 'nov', 'dec']
        last_3_months = ['oct', 'nov', 'dec']

        df['Absence_Days_Last_6_Months'] = df[last_6_months].sum(axis=1)
        df['Absence_Days_Last_3_Months'] = df[last_3_months].sum(axis=1)

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
