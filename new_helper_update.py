import streamlit as st 
import matplotlib.pyplot as plt
import seaborn as sns
import pandas as pd
import PIL.Image
from dotenv import load_dotenv
import os
import xlwings as xw 
import google.generativeai as genai


load_dotenv()

def plot_chart(folder_path, chart_type, data, x_col, y_col):
    if chart_type == "Line Chart":
        chart = sns.lineplot(data=data, x=x_col, y=y_col, markers='o')
        chart_fig = chart.get_figure()
        chart_name = f"{chart_type}_{x_col}_by_{y_col}.png"
        chart_fig.savefig(f"./{folder_path}/{chart_name}")
        # st.pyplot()
    elif chart_type == "Bar Chart":
        chart = sns.barplot(data=data, x=x_col, y=y_col)
        chart_fig = chart.get_figure()
        chart_name = f"{chart_type}_{x_col}_by_{y_col}.png"
        chart_fig.savefig(f"./{folder_path}/{chart_name}")
        # st.pyplot()
    elif chart_type == "Scatter Plot":
        chart = sns.scatterplot(data=data, x=x_col, y=y_col)
        chart_fig = chart.get_figure()
        chart_name = f"{chart_type}_{x_col}_by_{y_col}.png"
        chart_fig.savefig(f"./{folder_path}/{chart_name}")
        # st.pyplot()
    elif chart_type == "Pie Chart":
        pie_data = data.groupby(x_col)[y_col].sum()
        chart = pie_data.plot.pie(autopct='%1.1f%%', startangle=90)
        chart_fig = chart.get_figure()
        chart_name = f"{chart_type}_{x_col}_by_{y_col}.png"
        chart_fig.savefig(f"./{folder_path}/{chart_name}")
        # st.pyplot()
    
    return f"/Users/vietannguyen/TechX/demo_apps/automation_xlwings/{folder_path}/{chart_name}", chart_name


def remove_chart(file_path):
    if os.path.isfile(file_path):
        os.remove(file_path)
        

def create_pivot_table(data, aggfunc, values, index, columns=None):
    # data[index] = pd.to_numeric(df['Sales'], errors='coerce')
    pivot_df = pd.pivot_table(data, 
                          values = values, # Các chỉ tiêu nhóm theo hàng
                          index = index, # Các chỉ tiêu nhóm theo cột
                          columns=columns,
                          aggfunc=aggfunc, # Hàm tổng hợp
                          fill_value=0 # That thế các giá trị NaN bằng 0
                          )

    return pivot_df


def generate_report_from_chart(chart_folder, chart_name):
    genai.configure(api_key="Thay bằng key của các bạn")
    model = genai.GenerativeModel("gemini-1.5-flash")
    
    if chart_name.endswith(('.png', '.jpg', '.jpeg')):
        file_path = os.path.join(chart_folder, chart_name)
        organ = PIL.Image.open(file_path)
                
        response = model.generate_content(["According to the chart, generate a comprehensive report. Thus, what insight could be taken from it. Write at most 100 words", organ])
        bot_response = response.text.replace("*", "")
            
    return bot_response


def overwrite_sheet(workbook, sheet_name):
    """
    Overwrite an Excel sheet or create it if it doesn't exist.
    
    Args:
    - file_path (str): Path to the Excel file.
    - sheet_name (str): Name of the sheet to overwrite.
    - data (list of lists): Data to write to the sheet.
    """
    # Open the workbook
    # workbook = xw.Book(wb_name)

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
    
    return sheet
    

def generate_excel_report(data, reports, report_name):
    """
    Generates an Excel report containing:
    - Original dataset
    - Pivot tables from aggregation
    - Charts (saved as images)
    - AI-generated insights

    Args:
    - data (pd.DataFrame): The original dataset.
    - reports (list): A list of reports with pivot tables, charts, and insights.
    - report_name (str): The output Excel file name (without extension).

    Returns:
    - str: The absolute path to the generated report.
    """

    report_filename = f"{report_name}.xlsx"
    report_path = os.path.abspath(report_filename)

    # Use Pandas' ExcelWriter with XlsxWriter as the engine
    with pd.ExcelWriter(report_path, engine='xlsxwriter') as writer:
        workbook = writer.book

        # Save the original dataset in the first sheet
        data.to_excel(writer, sheet_name="Datasource", index=False)

        # Process each report
        for report in reports:
            sheet_name = report["sheet_name"]

            # Save the pivot table to Excel
            report["pivot_table"].to_excel(writer, sheet_name=sheet_name, index=False)

            # Get the worksheet to add images and insights
            worksheet = writer.sheets[sheet_name]

            # Add Chart Image if Exists
            chart_path = report["chart_path"]
            if os.path.exists(chart_path):
                worksheet.insert_image("F1", chart_path, {"x_scale": 0.5, "y_scale": 0.5})

            # Add AI-Generated Insight in Cell F24
            worksheet.write("F24", report["insight"])

def refresh_session(folder_path):
    # if "refresh_triggered" not in st.session_state:
    #     st.session_state.refresh_triggered = False
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
    else:
        for filename in os.listdir(folder_path):
            file_path = os.path.join(folder_path, filename)
            
            if os.path.isfile(file_path):
                os.remove(file_path)