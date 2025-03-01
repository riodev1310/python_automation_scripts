import streamlit as st 
import pandas as pd
import os
from new_helper import *
from PIL import Image
from streamlit_js_eval import streamlit_js_eval


st.set_page_config(page_title="Report Generator", page_icon=":robot_face:")

st.title(":robot_face: Report Generator")

file = st.sidebar.file_uploader(":file_folder: Upload a file", type=(["csv"]))

if "reports" not in st.session_state:
    st.session_state.reports = []

if "no" not in st.session_state:
    st.session_state.no = 0

chart_folder_name = "charts"
chart_folder_path = f"./{chart_folder_name}"


element_num = 0

if not os.path.exists(chart_folder_name):
    os.makedirs(chart_folder_name)

if file is not None:
    data = pd.read_csv(file)
    if st.button("Refresh Session"):
        refresh_session(chart_folder_name)
        streamlit_js_eval(js_expressions="parent.window.location.reload()")
    
    st.subheader("Aggregation Options")
    # Categorical values
    categorical_columns = data.select_dtypes(include=["object"]).columns.tolist()
    # Numerical values
    numeric_columns = data.select_dtypes(include=["float64", "int64"]).columns.tolist()
    
    if len(categorical_columns) > 0 and len(numeric_columns) > 0:
        category_col = st.selectbox("Choose a categorical column for grouping:", categorical_columns, key="category")
        
        numeric_col = st.selectbox("Choose a numeric column for aggregation:", numeric_columns, key="numeric")
        
        aggregation_function = st.selectbox(
            "Aggregation function: ",
            ["sum", "mean", "count", "min", "max"]
        )
        
        if "category" not in st.session_state:
            st.session_state.category = categorical_columns[0]

        if "numeric" not in st.session_state:
            st.session_state.numeric = numeric_col[0]
        
        if aggregation_function == "sum":
            aggregated_data = data.groupby(category_col)[numeric_col].sum().reset_index()
        elif aggregation_function == "mean":
            aggregated_data = data.groupby(category_col)[numeric_col].mean().reset_index()
        elif aggregation_function == "count":
            aggregated_data = data.groupby(category_col)[numeric_col].count().reset_index()
        elif aggregation_function == "min":
            aggregated_data = data.groupby(category_col)[numeric_col].min().reset_index()
        elif aggregation_function == "max":
            aggregated_data = data.groupby(category_col)[numeric_col].max().reset_index()
        
        # if aggregated_data is not None:
        st.subheader(f"Aggregated Data: {aggregation_function} of {numeric_col} by {category_col}")
        st.dataframe(aggregated_data.style.background_gradient(cmap="viridis"))
        
        st.subheader("Visualize your data")
        
        chart_type = st.selectbox(
            "Choose chart type",
            ["Line Chart", "Bar Chart", "Scatter Plot", "Pie Chart"]
        )
        
        if st.button("Plot Graph"):
            chart_path, chart_name = plot_chart(chart_folder_path, chart_type, aggregated_data, st.session_state.category, st.session_state.numeric)
            # pivot_table = create_pivot_table(aggregated_data, aggregation_function, st.session_state.numeric, st.session_state.category)
            response = generate_report_from_chart(chart_folder_path, chart_name)
            
            report = {
                "pivot_table": aggregated_data,
                "chart_path": chart_path,
                "sheet_name": f"Sheet {st.session_state.no}",
                "insight": response
            }
            
            st.session_state.reports.append(report)
            st.session_state.no += 1
            # st.session_state.category = categorical_columns[0]
            # st.session_state.numeric = numeric_col[0]
            
            
        # for filename in os.listdir(chart_folder_name):
        #     if filename.endswith(('.png', '.jpg', '.jpeg')):
        #         file_path = os.path.join(chart_folder_path, filename)
                
        #         image = Image.open(file_path)
        #         st.image(image, caption=filename, use_column_width=True)
        #         element_num += 1
        #         if st.button("Remove chart", key=element_num):
        #             remove_chart(file_path)
        #             st.session_state.remove()
        
        for report in st.session_state.reports:
            filename = report["chart_path"]
            if filename.endswith(('.png', '.jpg', '.jpeg')):
                file_path = os.path.join(chart_folder_path, filename)
                
                image = Image.open(file_path)
                st.image(image, caption=filename, use_column_width=True)
                element_num += 1
                if st.button("Remove chart", key=element_num):
                    remove_chart(file_path)
                    st.session_state.reports.remove(report)
        
        if st.button("Generate Report"):
            # report_name = st.text_input("Report name: ")
            # if report_name is not None:
            generate_excel_report(data, st.session_state.reports, "report")
            with open("report.xlsx", "rb") as file:
                excel_data = file.read()
            st.download_button(label="Download", data=excel_data, file_name="report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")