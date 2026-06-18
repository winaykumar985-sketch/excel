import pandas as pd
import streamlit as st
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io
import re

st.set_page_config(page_title="Autonomous Data Beautifier", layout="wide")
st.title("📊 The Messy Data Cleanup & Styling Engine")
st.write("Upload any chaotic Excel or CSV sheet and get a perfectly formatted, corporate-ready Excel file instantly.")

# File Upload Widget
uploaded_file = st.file_uploader("Choose a messy Excel or CSV file", type=['xlsx', 'csv'])

if uploaded_file is not None:
    try:
        # Load and clean up raw file structures
        if uploaded_file.name.endswith('.csv'):
            raw_bytes = uploaded_file.getvalue()
            # Decode file cleanly
            raw_text = raw_bytes.decode('utf-8', errors='ignore')
            lines = raw_text.splitlines()
            
            # Find where the actual table headers start dynamically
            header_index = 0
            for i, line in enumerate(lines):
                cleaned_line = line.strip().upper()
                # A real header row must contain at least TWO valid data columns
                matches = sum(1 for k in ['ID', 'INVOICE', 'NAME', 'DATE', 'PRICE', 'AMOUNT', 'CATEGORY', 'COUNTRY'] if k in cleaned_line)
                if matches >= 2:
                    header_index = i
                    break
            
            # Filter out all lines before the real headers
            valid_lines = lines[header_index:]
            
            # Reconstruct the clean text data buffer
            clean_csv_data = "\n".join(valid_lines)
            
            # Read cleanly, automatically bypassing empty metadata rows
            df = pd.read_csv(io.StringIO(clean_csv_data), on_bad_lines='skip')
        else:
            df = pd.read_excel(uploaded_file)

        # -------------------------------------------------------------
        # CORE CLEANING PIPELINE
        # -------------------------------------------------------------
        # Drop rows that are completely empty or filled with NaNs
        df = df.dropna(how='all')
        
        # Clean up column headers (strip spaces, resolve Unnamed indicators)
        df.columns = [str(col).strip() for col in df.columns]
        df = df.loc[:, ~df.columns.str.contains('^Unnamed', case=False, na=False)]
        df.columns = [col.title() for col in df.columns]
        
        # Row-level string cleaning and processing
        for col in df.columns:
            # 1. Clean up messy text/spaces in text columns
            if df[col].dtype == 'object':
                df[col] = df[col].astype(str).str.strip()
                df[col] = df[col].replace({'nan': None, 'N/A': None, 'None': None, 'NaN': None})
            
            col_lower = col.lower()
            
            # 2. Smart Date Unification Engine
            if 'date' in col_lower:
                df[col] = pd.to_datetime(df[col], errors='coerce').dt.strftime('%Y-%m-%d')
                df[col] = df[col].replace({'NaT': None})
                
            # 3. Currency and Number Standardization
            elif any(k in col_lower for k in ['price', 'amount', 'sales', 'revenue', 'cost']):
                # Remove currency formatting text symbols safely (₹, $, commas)
                df[col] = df[col].astype(str).str.replace(r'[₹$,\s]', '', regex=True)
                df[col] = pd.to_numeric(df[col], errors='coerce')

        # Drop any remaining rows that are completely blank after cleaning data types
        df = df.dropna(how='all')

        # Display Data View Preview
        st.subheader("👀 Preview of Perfectly Parsed & Cleaned Data")
        st.dataframe(df)

        # -------------------------------------------------------------
        # EXCEL STYLING & BEAUTIFICATION ENGINE
        # -------------------------------------------------------------
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Clean Data', index=False)
            
            workbook = writer.book
            worksheet = writer.sheets['Clean Data']
            
            # Professional Corporate Theme Configuration (Classic Navy Slate)
            header_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid') 
            header_font = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
            
            data_font = Font(name='Calibri', size=11)
            center_align = Alignment(horizontal='center', vertical='center')
            left_align = Alignment(horizontal='left', vertical='center')
            right_align = Alignment(horizontal='right', vertical='center')
            
            thin_border_side = Side(border_style="thin", color="D9D9D9")
            cell_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
            
            # Style column headers cleanly
            for col_num in range(1, len(df.columns) + 1):
                cell = worksheet.cell(row=1, column=col_num)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = center_align
            
            # Format and align data rows intelligently based on content type
            for row in range(2, worksheet.max_row + 1):
                for col in range(1, worksheet.max_column + 1):
                    cell = worksheet.cell(row=row, column=col)
                    cell.font = data_font
                    cell.border = cell_border
                    
                    # Align metrics right, keys/dates center, text left
                    val_str = str(cell.value or '').lower()
                    if isinstance(cell.value, (int, float)):
                        cell.alignment = right_align
                        # Apply clear accounting/number format
                        cell.number_format = '#,##0.00'
                    elif re.match(r'^\d{4}-\d{2}-\d{2}$', val_str) or 'inv-' in val_str or 'trx-' in val_str:
                        cell.alignment = center_align
                    else:
                        cell.alignment = left_align
            
            # Set gridlines visibility explicitly
            worksheet.views.sheetView[0].showGridLines = True
            
            # Auto-fit column layout sizes safely to handle unexpected widths
            for col in worksheet.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = get_column_letter(col[0].column)
                worksheet.column_dimensions[col_letter].width = max(max_len + 4, 13)
        
        processed_data = output.getvalue()
        
        st.success("🎉 Your data has been cleaned and styled beautifully!")
        st.download_button(
            label="📥 Download Beautiful Clean Sheet",
            data=processed_data,
            file_name=f"Cleaned_{uploaded_file.name if uploaded_file.name.endswith('.xlsx') else 'Data.xlsx'}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"Error processing file: {e}")
