import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# 1. Suppose Python has already cleaned the messy data into this dataframe
df = pd.DataFrame({
    'Transaction Date': ['2026-06-18', '2026-06-19'],
    'Customer Name': ['Vinay Kumar', 'Amit Sharma'],
    'Amount': [15000.00, 2450.50]
})

# 2. Write to Excel using openpyxl engine
with pd.ExcelWriter('Beautiful_Clean_Sheet.xlsx', engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='Clean Data', index=False)
    
    # Get the workbook and active sheet to apply styles
    workbook = writer.book
    worksheet = writer.sheets['Clean Data']
    
    # Define corporate styles
    header_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
    header_font = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
    data_font = Font(name='Calibri', size=11)
    center_align = Alignment(horizontal='center', vertical='center')
    
    # Style the Header Row
    for col_num in range(1, len(df.columns) + 1):
        cell = worksheet.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align

    # Auto-adjust column widths so nothing is cut off
    for col in worksheet.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        worksheet.column_dimensions[col_letter].width = max(max_len + 3, 12)

print("Beautiful sheet generated successfully!")
