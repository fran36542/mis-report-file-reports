import streamlit as st
import pandas as pd
from io import BytesIO
import re
import math
from datetime import datetime

# Set page configuration
st.set_page_config(
    page_title="Trans Type :- GOODS RECEIPT NOTE Factory Outward Report",
    page_icon="📊",
    layout="wide"
)

# Fixed columns for Goods Receipt Note (with Karat added)
FIXED_COLUMNS = [
    "Trans Date", "Doc No", "Party Name", "Style No", "Karat", "Variant Name", 
    "Net Wt", "Wastage Perc", "Pg Wt", "Pg Wastage Wt", "Line Remark"
]

# Custom column widths
COLUMN_WIDTHS = {
    "Trans Date": 18.5,
    "Doc No": 15,
    "Party Name": 23,
    "Style No": 20,
    "Karat": 13,
    "Variant Name": 14,
    "Net Wt": 10,
    "Wastage Perc": 8,
    "Pg Wt": 8,
    "Pg Wastage Wt": 10,
    "Line Remark": 20
}

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

def extract_title_from_excel(df):
    """Extract title from the Excel file (e.g., SALE INVOICE)"""
    # Look for common title patterns in the first few rows
    for idx, row in df.head(5).iterrows():
        for cell in row:
            if pd.notna(cell):
                cell_str = str(cell).strip()
                # Check for SALE INVOICE or similar patterns
                if 'sale invoice' in cell_str.lower():
                    return "SALE INVOICE"
                elif 'goods receipt' in cell_str.lower():
                    return "GOODS RECEIPT NOTE"
                elif 'trans type' in cell_str.lower():
                    # Extract the value after Trans Type
                    match = re.search(r'Trans Type\s*[:-]*\s*(.+)', cell_str, re.IGNORECASE)
                    if match:
                        return match.group(1).strip()
    
    # Default title if not found
    return "GOODS RECEIPT NOTE"

def extract_grand_total_row(df):
    """Extract the grand total row from the original dataframe"""
    for idx, row in df.iterrows():
        # Check if this row contains "grand total" (case insensitive)
        if any('grand total' in str(cell).lower() for cell in row.values if pd.notna(cell)):
            return row
    return None

def extract_dates_from_data(df):
    """Extract From Date and To Date from the Trans Date column"""
    if 'Trans Date' not in df.columns:
        return "01/09/2025", "02/09/2025"  # Default dates
    
    try:
        # Convert to datetime and find min and max
        date_series = pd.to_datetime(df['Trans Date'], errors='coerce')
        valid_dates = date_series.dropna()
        
        if len(valid_dates) > 0:
            from_date = valid_dates.min().strftime("%d/%m/%Y")
            to_date = valid_dates.max().strftime("%d/%m/%Y")
            return from_date, to_date
    except:
        pass
    
    # Try to extract dates from the original Excel data
    try:
        raw_df = pd.read_excel(uploaded_file, header=None)
        for idx, row in raw_df.head(10).iterrows():
            for cell in row:
                if pd.notna(cell):
                    cell_str = str(cell).lower()
                    # Look for date patterns
                    if 'from date' in cell_str or 'to date' in cell_str:
                        # Try to extract dates using regex
                        date_match = re.search(r'(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})', str(cell))
                        if date_match:
                            return date_match.group(1), date_match.group(1) if 'from date' in cell_str and 'to date' in cell_str else "02/09/2025"
    except:
        pass
    
    return "01/09/2025", "02/09/2025"  # Default dates

def clean_trans_date(date_value):
    """Remove time portion from Trans Date values"""
    if pd.isna(date_value):
        return ""
    
    date_str = str(date_value)
    
    # If it contains time portion, remove it
    if ' ' in date_str and (':' in date_str or '00:00:00' in date_str):
        return date_str.split(' ')[0]
    
    return date_str

def extract_karat(style_no):
    """Extract Karat value from Style No text"""
    if pd.isna(style_no) or not isinstance(style_no, str):
        return ""
    
    # Look for karat patterns (18KT, 22KT, 24KT)
    karat_pattern = r'\b(18KT|22KT|24KT)\b'
    match = re.search(karat_pattern, style_no, re.IGNORECASE)
    
    if match:
        return match.group(1).upper()
    
    return ""

def process_excel(uploaded_file):
    """Process the uploaded Excel file and extract required data"""
    try:
        # Read Excel file
        df = pd.read_excel(uploaded_file, header=None)
        
        # Extract title from Excel
        title = extract_title_from_excel(df)
        
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
        available_cols = [col for col in FIXED_COLUMNS if col in df.columns and col != "Karat"]
        df = df[available_cols]
        
        # Add Karat column after Style No
        if "Style No" in df.columns:
            # Extract karat values from Style No
            karat_values = df["Style No"].apply(extract_karat)
            
            # Insert Karat column after Style No
            style_no_idx = df.columns.get_loc("Style No")
            df.insert(style_no_idx + 1, "Karat", karat_values)
        
        # Clean the data - remove time from Trans Date
        if 'Trans Date' in df.columns:
            df['Trans Date'] = df['Trans Date'].apply(clean_trans_date)
        
        # Clean other columns
        for col in df.columns:
            if df[col].dtype == 'object':
                df[col] = df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)
        
        # Remove empty rows
        df = df.dropna(how='all')
        
        # Fill missing values with empty strings
        df = df.fillna('')
        
        # Remove any rows that might be grand total rows (we'll add the proper one later)
        df = df[~df.iloc[:, 0].astype(str).str.contains('grand total', case=False, na=False)]
        
        return df, grand_total_row, title
    
    except Exception as e:
        st.error(f"Error processing file: {str(e)}")
        return pd.DataFrame(columns=FIXED_COLUMNS), None, "GOODS RECEIPT NOTE"

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

def create_download_file(df, grand_total_row, from_date, to_date, title):
    """Create an Excel file for download with proper formatting"""
    output = BytesIO()
    
    # Use nan_inf_to_errors option to handle NaN/INF values
    with pd.ExcelWriter(output, engine='xlsxwriter', engine_kwargs={'options': {'nan_inf_to_errors': True}}) as writer:
        # Create header rows
        header_df = pd.DataFrame({f'Trans Type :- {title}': []})
        header_df.to_excel(writer, sheet_name='Report', index=False)
        
        date_df = pd.DataFrame({f'From Date :- {from_date}   To Date :- {to_date}': []})
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
        
        line_remark_format = workbook.add_format({
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
        worksheet.merge_range('A1:K1', f'Trans Type :- {title}', header_format)
        worksheet.merge_range('A2:K2', f'From Date :- {from_date}   To Date :- {to_date}', date_format)
        
        # Format column headers
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(3, col_num, value, column_header_format)
        
        # Format data rows
        for row_num in range(4, len(df) + 4):
            # Set row height to 25
            worksheet.set_row(row_num, 25)
            
            # Apply data formatting to all cells using safe_write
            for col_num in range(len(df.columns)):
                cell_value = df.iloc[row_num-4, col_num]
                
                # Special formatting for Line Remark column
                if df.columns[col_num] == 'Line Remark':
                    safe_write_cell(worksheet, row_num, col_num, cell_value, line_remark_format)
                else:
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
        
        # Set custom column widths
        for idx, col in enumerate(df.columns):
            if col in COLUMN_WIDTHS:
                worksheet.set_column(idx, idx, COLUMN_WIDTHS[col])
            else:
                # Default width for columns not specified
                max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
                worksheet.set_column(idx, idx, min(max_len, 15))
    
    processed_data = output.getvalue()
    return processed_data

def main():
    st.title("📊 Trans Type :- SALE INVOICE")
    st.markdown("Upload your Excel file to convert it to the proper format")
    
    # File upload
    uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx', 'xls'], key="file_uploader")
    
    if uploaded_file is not None:
        # Process the file
        df, grand_total_row, title = process_excel(uploaded_file)
        
        if not df.empty:
            st.success("File processed successfully!")
            
            # Extract dates from data
            from_date, to_date = extract_dates_from_data(df)
            
            if grand_total_row is not None:
                st.info("Grand Total row found and will be included in the output.")
            
            # Display the title
            st.info(f"Detected Title: {title}")
            
            # Display the data
            st.subheader("Processed Data")
            st.dataframe(df)
            
            # Show karat extraction examples
            if "Style No" in df.columns and "Karat" in df.columns:
                st.subheader("Karat Extraction Examples")
                example_df = df[["Style No", "Karat"]].head(3)
                st.dataframe(example_df)
            
            # Show grand total row if available
            if grand_total_row is not None:
                st.subheader("Grand Total Row (From Original File)")
                st.dataframe(pd.DataFrame([grand_total_row]))
            
            # Show extracted dates
            st.info(f"Extracted Dates: From {from_date} To {to_date}")
            
            # Create download file
            excel_data = create_download_file(df, grand_total_row, from_date, to_date, title)
            
            # Download button
            st.download_button(
                label="📥 Download Formatted Excel",
                data=excel_data,
                file_name=f"{title.replace(' ', '_')}.xlsx",
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