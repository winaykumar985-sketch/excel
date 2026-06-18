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
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
        
        st.subheader("👀 Preview of Raw Data")
        st.dataframe(df.head(10))

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
