import pandas as pd
import streamlit as st

# Streamlit App Title
st.title("Absence Summary Dashboard")

# File Upload
uploaded_file = st.file_uploader("Upload an Excel File", type=["xlsx"])

if uploaded_file:
    # Load the uploaded Excel File
    df = pd.ExcelFile(uploaded_file).parse(0)

    # Convert monthly columns (1–12) to numeric
    monthly_columns = list(range(1, 13))
    df[monthly_columns] = df[monthly_columns].apply(pd.to_numeric, errors='coerce').fillna(0)

    # Calculate Grand Total from monthly columns
    df['Grand Total'] = df[monthly_columns].sum(axis=1)

    # Calculate Last 6 months (July–December) and Last 3 months (October–December)
    df['Absence Days (Last 6 Months)'] = df[[7, 8, 9, 10, 11, 12]].sum(axis=1)
    df['Absence Days (Last 3 Months)'] = df[[10, 11, 12]].sum(axis=1)

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
