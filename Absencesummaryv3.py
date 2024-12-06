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

    # Clean column names (optional, depending on inspection results)
    df.columns = df.columns.str.strip().str.replace(" ", "_").astype(str)

    # Define monthly columns dynamically based on cleaned column names
    try:
        monthly_columns = [str(i) for i in range(1, 13)]  # Use strings if column names are like "1", "2", ...
        df[monthly_columns] = df[monthly_columns].apply(pd.to_numeric, errors='coerce').fillna(0)

        # Calculate Grand Total from monthly columns (1–12)
        df['Grand_Total'] = df[monthly_columns].sum(axis=1)

        # Calculate Last 6 months (July–December) and Last 3 months (October–December)
        df['Absence_Days_Last_6_Months'] = df[['7', '8', '9', '10', '11', '12']].sum(axis=1)
        df['Absence_Days_Last_3_Months'] = df[['10', '11', '12']].sum(axis=1)

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
        st.error(f"Missing expected columns: {e}")
        st.write("Please check the uploaded file and ensure it contains all required columns (1–12, etc.).")
