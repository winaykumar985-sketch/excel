import pandas as pd
import streamlit as st
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference
import io
import re

st.set_page_config(page_title="Corporate Data Engine & Visualizer", layout="wide")
st.title("📊 Administrative Data Engine & Visualizer")
st.write("Clean, parse, and visually aggregate standard public records and corporate sheets instantly.")

# File Upload Widget
uploaded_file = st.file_uploader("Choose an administrative Excel or CSV file", type=['xlsx', 'csv'])

if uploaded_file is not None:
    try:
        # Load and clean up raw file structures
        if uploaded_file.name.endswith('.csv'):
            raw_bytes = uploaded_file.getvalue()
            raw_text = raw_bytes.decode('utf-8', errors='ignore')
            lines = raw_text.splitlines()
            
            # Find where the data grid matrix headers actually begin dynamically
            header_index = 0
            for i, line in enumerate(lines):
                cleaned_line = line.strip().upper()
                # Expanded structural flags matching public administrative/civil rights criteria
                matches = sum(1 for k in ['ID', 'NAME', 'ALLEGATION', 'TOTAL', 'SEX', 'RACE', 'DISABILITY', 'BULLYING'] if k in cleaned_line)
                if matches >= 1:
                    header_index = i
                    break
            
            valid_lines = lines[header_index:]
            clean_csv_data = "\n".join(valid_lines)
            df = pd.read_csv(io.StringIO(clean_csv_data), on_bad_lines='skip')
        else:
            # Force read excel seamlessly, managing sheet index arrays safely
            df = pd.read_excel(uploaded_file)

        # -------------------------------------------------------------
        # CORE ADMINISTRATIVE DATA CLEANING PIPELINE
        # -------------------------------------------------------------
        df = df.dropna(how='all')
        df.columns = [str(col).strip() for col in df.columns]
        df = df.loc[:, ~df.columns.str.contains('^Unnamed', case=False, na=False)]
        
        # Format columns nicely
        df.columns = [col.title() if not col.isupper() else col for col in df.columns]
        
        for col in df.columns:
            # Handle privacy indicators or masking dots typical in civil reports
            if df[col].dtype == 'object':
                df[col] = df[col].astype(str).str.strip()
                df[col] = df[col].replace({r'^\.$': '0', 'nan': None, 'N/A': None, 'None': None, 'NaN': None, '': None}, regex=True)
            
            col_lower = col.lower()
            
            # Numeric conversion targeting metrics, totals, or target values
            if any(k in col_lower for k in ['total', 'count', 'sum', 'price', 'amount', 'allegation', 'basis', 'sex', 'race', 'disability']):
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        df = df.dropna(how='all')

        # Discover category layouts dynamically
        numeric_cols = [col for col in df.columns if df[col].dtype in ['int64', 'float64'] and not any(k in col.lower() for k in ['id', 'year', 'code'])]
        text_cols = [col for col in df.columns if df[col].dtype == 'object' and not any(k in col.lower() for k in ['id', 'code'])]

        # -------------------------------------------------------------
        # STREAMLIT USER INTERFACE METRICS & TABS
        # -------------------------------------------------------------
        st.write("---")
        st.subheader("📈 Core Matrix Performance Analytics")
        
        kpi1, kpi2, kpi3 = st.columns(3)
        kpi1.metric("Total Rows Evaluated", f"{len(df)}")
        
        if numeric_cols:
            kpi2.metric(f"Cumulative Sum ({numeric_cols[0]})", f"{df[numeric_cols[0]].sum():,g}")
            if len(numeric_cols) > i:
                kpi3.metric(f"Secondary Category Sum", f"{df[numeric_cols[-1]].sum():,g}")
            else:
                kpi3.metric("Average System Metric", f"{df[numeric_cols[0]].mean():.2f}")
        else:
            kpi2.metric("Metric Values", "0")
            kpi3.metric("Metric Mean", "0.00")

        tab1, tab2 = st.tabs(["👀 Formatted Grid Summary", "📊 Interactive Matrix Charts"])
        with tab1:
            st.dataframe(df, use_container_width=True)
        with tab2:
            if numeric_cols:
                st.write(f"#### 🍩 Distribution Level Breakdown across Top Registered Measures")
                # Melt wide tables into clean category plots if columns represent breakdown indices
                if len(numeric_cols) > 1 and len(df) < 50:
                    melted_df = df[numeric_cols].sum().reset_index()
                    melted_df.columns = ['Reporting Category Component', 'Aggregated Metrics']
                    st.bar_chart(data=melted_df, x='Reporting Category Component', y='Aggregated Metrics', use_container_width=True)
                elif text_cols:
                    chart_data = df.groupby(text_cols[0])[numeric_cols[0]].sum().reset_index()
                    st.bar_chart(data=chart_data, x=text_cols[0], y=numeric_cols[0], use_container_width=True)
            else:
                st.info("Unlock automated visualization vectors by feeding numerical array frames.")

        # -------------------------------------------------------------
        # PROFESSIONAL EXCEL TAB COMPILING & EMBEDDED CHART INJECTION
        # -------------------------------------------------------------
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Clean Data', index=False)
            workbook = writer.book
            data_sheet = workbook['Clean Data']
            
            dashboard_sheet = workbook.create_sheet(title="Executive Dashboard", index=0)
            dashboard_sheet.views.sheetView[0].showGridLines = True
            
            # Color Palettes (Corporate Navy Slate Theme)
            header_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid') 
            header_font = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
            data_font = Font(name='Calibri', size=11)
            center_align = Alignment(horizontal='center', vertical='center')
            left_align = Alignment(horizontal='left', vertical='center')
            right_align = Alignment(horizontal='right', vertical='center')
            thin_side = Side(border_style="thin", color="D9D9D9")
            cell_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
            
            # Format Raw Sheet View
            for col_num in range(1, len(df.columns) + 1):
                cell = data_sheet.cell(row=1, column=col_num)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = center_align
            
            for row in range(2, data_sheet.max_row + 1):
                for col in range(1, data_sheet.max_column + 1):
                    cell = data_sheet.cell(row=row, column=col)
                    cell.font = data_font
                    cell.border = cell_border
                    
                    if isinstance(cell.value, (int, float)):
                        cell.alignment = right_align
                        cell.number_format = '#,##0'
                    else:
                        cell.alignment = left_align

            for col in data_sheet.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = get_column_letter(col[0].column)
                data_sheet.column_dimensions[col_letter].width = max(max_len + 4, 13)

            # --- WRITE DASHBOARD MATRIX SUMMARY FOR EXCEL GRAPH PLOTTING ---
            dashboard_sheet["A2"] = "Administrative Metrics Summary"
            dashboard_sheet["A2"].font = Font(name='Calibri', size=16, bold=True, color='1F4E78')
            
            dashboard_sheet["A4"] = "Metric Segment Key"
            dashboard_sheet["B4"] = "Aggregated Values"
            dashboard_sheet["A4"].font = header_font
            dashboard_sheet["A4"].fill = header_fill
            dashboard_sheet["B4"].font = header_font
            dashboard_sheet["B4"].fill = header_fill
            
            # Build an internal summary lookup sheet grid
            if numeric_cols:
                summary_dict = {col: df[col].sum() for col in numeric_cols}
                for idx, (k, v) in enumerate(summary_dict.items()):
                    r = 5 + idx
                    dashboard_sheet.cell(row=r, column=1, value=str(k)).alignment = left_align
                    val_cell = dashboard_sheet.cell(row=r, column=2, value=float(v))
                    val_cell.number_format = '#,##0'
                    val_cell.alignment = right_align
                
                # Plot native layout charts right into the dashboard
                chart = BarChart()
                chart.type = "col"
                chart.style = 10
                chart.title = "Aggregated Factor Distribution Overview"
                chart.y_axis.title = "Reported Quantities"
                chart.x_axis.title = "Metrics Classification Category"
                
                data_ref = Reference(dashboard_sheet, min_col=2, min_row=4, max_row=4+len(summary_dict))
                cats_ref = Reference(dashboard_sheet, min_col=1, min_row=5, max_row=4+len(summary_dict))
                
                chart.add_data(data_ref, titles_from_data=True)
                chart.set_categories(cats_ref)
                chart.height = 14
                chart.width = 22
                dashboard_sheet.add_chart(chart, "D4")
            else:
                dashboard_sheet["A5"] = "Provide numeric arrays to populate visual layouts."

        processed_data = output.getvalue()
        
        st.write("---")
        st.success("🎉 Comprehensive metric layers and matching native charts successfully prepared!")
        st.download_button(
            label="📥 Download Structured Document with Charts",
            data=processed_data,
            file_name=f"Beautified_{uploaded_file.name if uploaded_file.name.endswith('.xlsx') else 'Document.xlsx'}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"Error processing file: {e}")
