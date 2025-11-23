import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Excel Cleaner", page_icon="📊")

st.title("📊 Excel Cleaner")
st.write("Upload your Excel or CSV file to clean it automatically.")

uploaded_file = st.file_uploader("Choose a file", type=['xlsx', 'xls', 'csv'])

if uploaded_file is not None:
    try:
        # Read file
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
            
        st.write("### Original Data Preview")
        st.dataframe(df.head())
        
        st.write(f"**Shape:** {df.shape[0]} rows, {df.shape[1]} columns")
        
        # Cleaning Options
        st.sidebar.header("Cleaning Options")
        
        drop_duplicates = st.sidebar.checkbox("Remove Duplicates", value=True)
        fill_na = st.sidebar.checkbox("Fill Missing Values", value=False)
        
        if drop_duplicates:
            df = df.drop_duplicates()
            
        if fill_na:
            fill_value = st.sidebar.text_input("Fill value (e.g., 0, N/A)", "N/A")
            df = df.fillna(fill_value)
            
        # Column selection
        st.write("### Select Columns to Keep")
        all_columns = df.columns.tolist()
        selected_columns = st.multiselect("Columns", all_columns, default=all_columns)
        
        if selected_columns:
            df_cleaned = df[selected_columns]
            
            st.write("### Cleaned Data Preview")
            st.dataframe(df_cleaned.head())
            
            # Download
            st.write("### Download")
            
            output = io.BytesIO()
            if uploaded_file.name.endswith('.csv'):
                df_cleaned.to_csv(output, index=False)
                mime = "text/csv"
                ext = "csv"
            else:
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_cleaned.to_excel(writer, index=False)
                mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                ext = "xlsx"
                
            st.download_button(
                label=f"Download Cleaned File",
                data=output.getvalue(),
                file_name=f"cleaned_{uploaded_file.name}",
                mime=mime
            )
            
    except Exception as e:
        st.error(f"Error processing file: {e}")
