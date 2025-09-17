import streamlit as st
import pandas as pd
from io import BytesIO
import re
import math

# Set page configuration
st.set_page_config(
    page_title="Lot Rejection Report Converter",
    page_icon="📊",
    layout="wide"
)

# Fixed columns
FIXED_COLUMNS = [
    "Trans Date", "Order No", "Group No", "Style Name", "Karat", 
    "Wt", "Operation Name", "Wc Name", "User Name", "Remark"
]

def find_header_row(df):
    """Find the row index where the required columns are found"""
    for idx, row in df.iterrows():
        # Convert row values to strings and clean them
        row_values = [str(cell).strip().lower() for cell in row.values if pd.notna(cell)]
        
        # Check if this row contains our required columns
        found_columns = 0
        for col in FIXED_COLUMNS:
            col_lower = col.lower()
            # Check if any cell in the row contains the column name
            if any(col_lower in cell_value for cell_value in row_values):
                found_columns += 1
        
        # If we found most of our columns, this is likely the header row
        if found_columns >= len(FIXED_COLUMNS) * 0.6:  # At least 60% match
            return idx
    
    return 0  # Default to first row if not found

def extract_grand_total_row(df):
    """Extract the grand total row from the original dataframe"""
    for idx, row in df.iterrows():
        # Check if this row contains "grand total" (case insensitive)
        if any('grand total' in str(cell).lower() for cell in row.values if pd.notna(cell)):
            return row
    return None

def process_excel(uploaded_file):
    """Process the uploaded Excel file and extract required data"""
    try:
        # Read Excel file
        df = pd.read_excel(uploaded_file, header=None)
        
        # Extract grand total row before processing
        grand_total_row = extract_grand_total_row(df)
        
        # Find the header row
        header_row_idx = find_header_row(df)
        
        # Read the file again with the correct header row
        df = pd.read_excel(uploaded_file, header=header_row_idx)
        
        # Clean column names
        df.columns = [str(col).strip() for col in df.columns]
        
        # Map similar column names to our fixed columns
        column_mapping = {}
        for col in df.columns:
            col_lower = col.lower()
            for fixed_col in FIXED_COLUMNS:
                fixed_col_lower = fixed_col.lower()
                
                # Check for different variations of column names
                if (fixed_col_lower in col_lower or 
                    col_lower in fixed_col_lower or
                    re.sub(r'[^a-z]', '', fixed_col_lower) in re.sub(r'[^a-z]', '', col_lower)):
                    column_mapping[col] = fixed_col
                    break
        
        # Rename columns
        df = df.rename(columns=column_mapping)
        
        # Keep only the columns we need
        available_cols = [col for col in FIXED_COLUMNS if col in df.columns]
        df = df[available_cols]
        
        # Clean the data
        for col in df.columns:
            if df[col].dtype == 'object':
                df[col] = df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)
        
        # Remove empty rows
        df = df.dropna(how='all')
        
        # Fill missing values with empty strings
        df = df.fillna('')
        
        # Remove any rows that might be grand total rows (we'll add the proper one later)
        df = df[~df.iloc[:, 0].astype(str).str.contains('grand total', case=False, na=False)]
        
        return df, grand_total_row
    
    except Exception as e:
        st.error(f"Error processing file: {str(e)}")
        return pd.DataFrame(columns=FIXED_COLUMNS), None

def safe_write_cell(worksheet, row, col, value, cell_format=None):
    """Safely write a cell value, handling NaN/INF values"""
    if pd.isna(value) or value is None:
        worksheet.write(row, col, "", cell_format)
    elif isinstance(value, (int, float)) and (math.isnan(value) or math.isinf(value)):
        worksheet.write(row, col, "", cell_format)
    else:
        # Convert to string if it's a complex type that might cause issues
        if not isinstance(value, (str, int, float, bool)):
            value = str(value)
        worksheet.write(row, col, value, cell_format)

def create_download_file(df, grand_total_row, from_date, to_date):
    """Create an Excel file for download with proper formatting"""
    output = BytesIO()
    
    # Use nan_inf_to_errors option to handle NaN/INF values
    with pd.ExcelWriter(output, engine='xlsxwriter', engine_kwargs={'options': {'nan_inf_to_errors': True}}) as writer:
        # Create a header DataFrame
        header_df = pd.DataFrame({'Lot Rejection Report': []})
        header_df.to_excel(writer, sheet_name='Report', index=False)
        
        # Add date range
        date_df = pd.DataFrame({f'From Date: {from_date} - To Date: {to_date}': []})
        date_df.to_excel(writer, sheet_name='Report', index=False, startrow=1)
        
        # Write the data
        df.to_excel(writer, sheet_name='Report', index=False, startrow=3)
        
        # Get workbook and worksheet objects
        workbook = writer.book
        worksheet = writer.sheets['Report']
        
        # Add formatting
        header_format = workbook.add_format({
            'bold': True, 
            'font_size': 16,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        date_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        column_header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#366092',
            'font_color': 'white',
            'text_wrap': True,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        data_format = workbook.add_format({
            'text_wrap': True,
            'border': 1,
            'valign': 'vcenter'
        })
        
        grand_total_format = workbook.add_format({
            'bold': True,
            'bg_color': '#FFFF00',  # Yellow background
            'border': 1,
            'valign': 'vcenter',
            'align': 'center',
            'text_wrap': True
        })
        
        # Apply formatting
        worksheet.merge_range('A1:J1', 'Lot Rejection Report', header_format)
        worksheet.merge_range('A2:J2', f'From Date: {from_date} - To Date: {to_date}', date_format)
        
        # Format column headers
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(3, col_num, value, column_header_format)
        
        # Format data rows
        for row_num in range(4, len(df) + 4):
            # Set row height
            worksheet.set_row(row_num, 25)
            
            # Apply data formatting to all cells using safe_write
            for col_num in range(len(df.columns)):
                cell_value = df.iloc[row_num-4, col_num]
                safe_write_cell(worksheet, row_num, col_num, cell_value, data_format)
        
        # Add grand total row if it exists
        if grand_total_row is not None:
            grand_total_row_num = len(df) + 4
            worksheet.set_row(grand_total_row_num, 25)  # Set height to 25
            
            # Write grand total values using safe_write
            for col_num in range(len(df.columns)):
                if col_num < len(grand_total_row):
                    cell_value = grand_total_row.iloc[col_num] if col_num < len(grand_total_row) else ''
                    safe_write_cell(worksheet, grand_total_row_num, col_num, cell_value, grand_total_format)
                else:
                    safe_write_cell(worksheet, grand_total_row_num, col_num, '', grand_total_format)
        
        # Set column widths
        if 'Trans Date' in df.columns:
            trans_date_idx = df.columns.get_loc('Trans Date')
            worksheet.set_column(trans_date_idx, trans_date_idx, 10)  # Trans Date width = 10
        
        if 'Style Name' in df.columns:
            style_name_idx = df.columns.get_loc('Style Name')
            worksheet.set_column(style_name_idx, style_name_idx, 30)  # Style Name width = 30
        
        # Set default width for other columns
        for idx, col in enumerate(df.columns):
            if col not in ['Trans Date', 'Style Name']:
                max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
                worksheet.set_column(idx, idx, min(max_len, 20))
    
    processed_data = output.getvalue()
    return processed_data

def main():
    st.title("📊 Lot Rejection Report Converter")
    st.markdown("Upload your raw Excel file to convert it to the proper Lot Rejection Report format")
    
    # File upload
    uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx', 'xls'])
    
    if uploaded_file is not None:
        # Process the file
        df, grand_total_row = process_excel(uploaded_file)
        
        if not df.empty:
            st.success("File processed successfully!")
            
            if grand_total_row is not None:
                st.info("Grand Total row found and will be included in the output.")
            
            # Date inputs
            col1, col2 = st.columns(2)
            with col1:
                from_date = st.date_input("From Date")
            with col2:
                to_date = st.date_input("To Date")
            
            # Display the data
            st.subheader("Processed Data")
            st.dataframe(df)
            
            # Show grand total row if available
            if grand_total_row is not None:
                st.subheader("Grand Total Row (From Original File)")
                st.dataframe(pd.DataFrame([grand_total_row]))
            
            # Create download file
            excel_data = create_download_file(df, grand_total_row, from_date.strftime("%d/%m/%Y"), to_date.strftime("%d/%m/%Y"))
            
            # Download button
            st.download_button(
                label="📥 Download Formatted Excel",
                data=excel_data,
                file_name="Lot_Rejection_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("No data could be extracted. Please check the file format.")
            
            # Debug information
            with st.expander("Debug Information"):
                st.write("Trying to read the file without header detection...")
                try:
                    raw_df = pd.read_excel(uploaded_file)
                    st.write("Raw DataFrame columns:")
                    st.write(list(raw_df.columns))
                    st.write("First 5 rows:")
                    st.write(raw_df.head())
                    
                    # Check for grand total row
                    gt_row = extract_grand_total_row(raw_df)
                    if gt_row is not None:
                        st.write("Grand Total row found:")
                        st.write(gt_row)
                    else:
                        st.write("No Grand Total row found in the original file.")
                        
                except Exception as e:
                    st.error(f"Error reading file: {e}")

if __name__ == "__main__":
    main()