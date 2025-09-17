import streamlit as st
import pandas as pd
import io
from datetime import datetime
import xlsxwriter
import numpy as np

# Set page configuration
st.set_page_config(
    page_title="Goods Receipt Note Pivot Table Generator",
    page_icon="📊",
    layout="wide"
)

# Title and description
st.title("📊 Goods Receipt Note Pivot Table Generator")
st.markdown("Upload your GOODS_RECEIPT_NOTE Excel file to generate a pivot table with Party Name, Karat, and summary statistics.")

# File upload section
uploaded_file = st.file_uploader(
    "Choose an Excel file", 
    type=['xlsx'],
    help="Upload your GOODS_RECEIPT_NOTE Excel file"
)

if uploaded_file is not None:
    try:
        # Read the Excel file with improved error handling
        df = pd.read_excel(uploaded_file, sheet_name='Report', skiprows=3)
        
        # Clean column names by removing extra spaces
        df.columns = df.columns.str.strip()
        
        # Display original data
        st.subheader("Original Data Preview")
        st.dataframe(df.head())
        
        # Check if required columns exist (with case-insensitive matching)
        required_columns = ['Party Name', 'Karat', 'Net Wt', 'Pg Wt', 'Wastage Perc', 'Pg Wastage Wt']
        available_columns = [col.strip() for col in df.columns]
        
        # Find matching columns (case-insensitive)
        matched_columns = []
        for req_col in required_columns:
            found = False
            for avail_col in available_columns:
                if req_col.lower() == avail_col.lower():
                    matched_columns.append(avail_col)
                    found = True
                    break
            if not found:
                st.error(f"Required column '{req_col}' is missing from the uploaded file.")
                st.stop()
        
        # Rename columns to standard names for processing
        column_mapping = {}
        for req_col in required_columns:
            for avail_col in available_columns:
                if req_col.lower() == avail_col.lower():
                    column_mapping[avail_col] = req_col
        
        df = df.rename(columns=column_mapping)
        
        # Convert numeric columns to appropriate data types
        numeric_columns = ['Net Wt', 'Pg Wt', 'Wastage Perc', 'Pg Wastage Wt']
        for col in numeric_columns:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce')
        
        # Create pivot table with margins for total row
        pivot_table = pd.pivot_table(
            df,
            values=['Net Wt', 'Pg Wt', 'Wastage Perc', 'Pg Wastage Wt'],
            index=['Party Name', 'Karat'],
            aggfunc='sum',
            fill_value=0,
            margins=True,
            margins_name='TOTAL'
        ).reset_index()
        
        # Rename columns for clarity
        pivot_table.columns = ['Party Name', 'KARAT', 'Sum of Net Wt', 'Sum of Pg Wt', 
                              'Sum of Wastage Perc', 'Sum of Pg Wastage Wt']
        
        # Format numeric columns
        for col in pivot_table.columns[2:]:
            pivot_table[col] = pivot_table[col].round(2)
        
        # Display pivot table with total
        st.subheader("Pivot Table with Total Row")
        st.dataframe(pivot_table)
        
        # Download section
        st.subheader("Download Pivot Table")
        
        # Create a buffer for the Excel file
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            pivot_table.to_excel(writer, sheet_name='Pivot Table', index=False)
            
            # Get the workbook and worksheet objects
            workbook = writer.book
            worksheet = writer.sheets['Pivot Table']
            
            # Define formats
            header_format = workbook.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'top',
                'border': 1,
                'bg_color': '#D9E1F2'
            })
            
            cell_format = workbook.add_format({
                'border': 1,
                'num_format': '#,##0.00'
            })
            
            total_format = workbook.add_format({
                'bold': True,
                'border': 1,
                'bg_color': '#FCE4D6',
                'num_format': '#,##0.00'
            })
            
            # Apply header format
            for col_num, value in enumerate(pivot_table.columns.values):
                worksheet.write(0, col_num, value, header_format)
            
            # Apply cell format to data cells
            for row in range(1, len(pivot_table) + 1):
                for col in range(len(pivot_table.columns)):
                    if row == len(pivot_table):  # Total row
                        worksheet.write(row, col, pivot_table.iloc[row-1, col], total_format)
                    else:
                        worksheet.write(row, col, pivot_table.iloc[row-1, col], cell_format)
            
            # Auto-adjust columns' width
            for idx, col in enumerate(pivot_table.columns):
                max_len = max(
                    pivot_table[col].astype(str).map(len).max(),
                    len(col)
                ) + 2
                worksheet.set_column(idx, idx, max_len)
        
        # Create download button
        st.download_button(
            label="📥 Download Pivot Table as Excel",
            data=output.getvalue(),
            file_name=f"goods_receipt_pivot_table_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.ms-excel"
        )
        
    except Exception as e:
        st.error(f"An error occurred while processing the file: {str(e)}")
        st.info("Please make sure the Excel file has the correct format with a 'Report' sheet.")
else:
    st.info("Please upload an Excel file to get started.")

# Add instructions
with st.expander("ℹ️ Instructions"):
    st.markdown("""
    ### How to use this tool:
    1. Upload your GOODS_RECEIPT_NOTE Excel file
    2. The app will automatically read the data and generate a pivot table
    3. The pivot table will have the following columns:
       - Party Name
       - KARAT
       - Sum of Net Wt
       - Sum of Pg Wt
       - Sum of Wastage Perc
       - Sum of Pg Wastage Wt
    4. A total row will be added at the bottom of the table
    5. Download the pivot table as an Excel file using the download button
    
    ### Requirements:
    - Your Excel file should have a sheet named 'Report'
    - The data should start from row 4 (first 3 rows are for headers/metadata)
    - The following columns must be present (case-insensitive):
      - Party Name
      - Karat
      - Net Wt
      - Pg Wt
      - Wastage Perc
      - Pg Wastage Wt
    """)