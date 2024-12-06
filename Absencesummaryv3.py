import pandas as pd
import streamlit as st

# Streamlit App Title
st.title("Absence Summary Dashboard")

# File Upload
uploaded_file = st.file_uploader("Upload an Excel File", type=["xlsx"])

if uploaded_file:
    # Load the uploaded Excel File
    df = pd.ExcelFile(uploaded_file).parse(0)

    # Debugging: Display column names
    st.write("Columns in the uploaded file:")
    st.write(df.columns.tolist())

    # Clean column names
    df.columns = df.columns.str.strip().str.replace(" ", "_").astype(str)

    # Dynamically identify columns that correspond to months
    # This assumes the month columns are integers or strings like "1", "2", ..., "12"
    monthly_columns = [col for col in df.columns if col.isdigit() and 1 <= int(col) <= 12]

    if not monthly_columns:
        st.error("No valid monthly columns (1–12) found in the uploaded file.")
    else:
        # Ensure monthly columns are numeric
        df[monthly_columns] = df[monthly_columns].apply(pd.to_numeric, errors='coerce').fillna(0)

        # Calculate Grand Total from monthly columns
        df['Grand_Total'] = df[monthly_columns].sum(axis=1)

        # Calculate Last 6 months (July–December) and Last 3 months (October–December)
        last_6_months = [col for col in monthly_columns if int(col) >= 7]
        last_3_months = [col for col in monthly_columns if int(col) >= 10]

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
