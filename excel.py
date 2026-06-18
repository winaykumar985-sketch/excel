import pandas as pd
import streamlit as st
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import io

st.set_page_config(page_title="Autonomous Data Beautifier", layout="wide")
st.title("📊 The Messy Data Cleanup & Styling Engine")
st.write("Upload any chaotic Excel or CSV sheet and get a perfectly formatted, corporate-ready Excel file instantly.")

# File Upload Widget
uploaded_file = st.file_uploader("Choose a messy Excel or CSV file", type=['xlsx', 'csv'])

if uploaded_file is not None:
    try:
        # Load data based on file type
        if uploaded_file.name.endswith('.csv'):
            # Read raw lines first to find where the actual table headers start
            raw_bytes = uploaded_file.getvalue()
            lines = raw_bytes.decode('utf-8', errors='ignore').split('\n')
            
            skip_rows = 0
            for i, line in enumerate(lines):
                # If a line contains actual key column identifiers, that's our header row!
                if 'INVOICE ID' in line or 'Invoice' in line or 'Transaction' in line or 'ID' in line:
                    skip_rows = i
                    break
            
            # Reload properly skipping the junk metadata headers
            uploaded_file.seek(0)
            df = pd.read_csv(uploaded_file, skiprows=skip_rows)
        else:
            df = pd.read_excel(uploaded_file)

        # Core Cleaning Pipeline
        df = df.dropna(how='all')
        df.columns = [str(col).strip().title() for col in df.columns]
        
        for col in df.columns:
            if 'date' in col.lower():
                df[col] = pd.to_datetime(df[col], errors='coerce').dt.strftime('%Y-%m-%d')
        
        # Excel Styling and Beautification
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Clean Data', index=False)
            
            workbook = writer.book
            worksheet = writer.sheets['Clean Data']
            
            header_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid') 
            header_font = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
            center_align = Alignment(horizontal='center', vertical='center')
            
            for col_num in range(1, len(df.columns) + 1):
                cell = worksheet.cell(row=1, column=col_num)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = center_align
            
            for col in worksheet.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = get_column_letter(col[0].column)
                worksheet.column_dimensions[col_letter].width = max(max_len + 4, 12)
        
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
