import xlwings as xw
import pandas as pd
import seaborn as sns 
import matplotlib.pyplot as plt
from datetime import datetime
import os


def create_pivot_table(data, aggfunc, values, index, columns=None):
    pivot_df = pd.pivot_table(data, 
                          values = values, # Các chỉ tiêu nhóm theo hàng
                          index = index, # Các chỉ tiêu nhóm theo cột
                          columns=columns,
                          aggfunc=aggfunc, # Hàm tổng hợp
                          fill_value=0 # That thế các giá trị NaN bằng 0
                          )

    return pivot_df
    

def convert_to_ddmmyyyy(date_str):
    # Định dạng ngày tháng có thể gặp phải
    possible_formats = [
        "%d/%m/%y", "%d/%m/%Y",  # dd/mm/yy, dd/mm/yyyy
        "%m/%d/%y", "%m/%d/%Y",  # mm/dd/yy, mm/dd/yyyy
        "%d-%m-%y", "%d-%m-%Y",  # dd-mm-yy, dd-mm-yyyy
        "%m-%d-%y", "%m-%d-%Y",  # mm-dd-yy, mm-dd-yyyy
        "%Y-%m-%d",              # yyyy-mm-dd
        "%d %b %Y", "%b %d %Y"    # dd Mmm yyyy, Mmm dd yyyy (tháng viết tắt)
    ]
    
    for fmt in possible_formats:
        try:
            # Cố gắng chuyển đổi chuỗi ngày tháng sang datetime với các định dạng thử
            date_obj = datetime.strptime(date_str, fmt)
            # Trả về ngày tháng theo format dd/mm/yyyy
            return date_obj.strftime("%d/%m/%Y")
        except ValueError:
            continue
    
    # Nếu không thể chuyển đổi, trả về thông báo lỗi
    return "Invalid date format"
    
def draw_bar_chart(data, x_col, y_col, legend=None):
    fig = plt.figure(figsize=(10, 6))
    if legend is None:
        sns.barplot(data=data, x=x_col, y=y_col)
    else:
        sns.barplot(data=data, x=x_col, y=y_col, hue=legend)
    
    return fig

def draw_pie_chart(value_counts, labels, title):
    fig = plt.figure(figsize=(10, 6))
    plt.pie(value_counts, labels=labels, autopct="%1.1f%%", colors=['pink', 'green'], startangle=90)
    plt.title(title)
    
    return fig


def overwrite_sheet(file_path, sheet_name, data):
    """
    Overwrite an Excel sheet or create it if it doesn't exist.
    
    Args:
    - file_path (str): Path to the Excel file.
    - sheet_name (str): Name of the sheet to overwrite.
    - data (list of lists): Data to write to the sheet.
    """
    # Open the workbook
    workbook = xw.Book(file_path)

    # Check if the sheet exists
    if sheet_name in [sheet.name for sheet in workbook.sheets]:
        # Get the existing sheet
        sheet = workbook.sheets[sheet_name]
        # Clear the sheet contents
        sheet.clear_contents()
    else:
        # Add a new sheet
        sheet = workbook.sheets.add(after=workbook.sheets[-1])
        sheet.name = sheet_name

    # Write data to the sheet
    sheet.range('A1').value = data
    
    return sheet


def remove_chart(file_path):
    if os.path.isfile(file_path):
        os.remove(file_path)