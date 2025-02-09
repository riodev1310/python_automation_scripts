import streamlit as st
import pandas as pd
import os
from openpyxl import load_workbook

def save_to_excel(file_path, df):
    """Lưu DataFrame vào chính tệp Excel đã tải lên."""
    with pd.ExcelWriter(file_path, engine='openpyxl', mode='w') as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")

def search_data(df, column, query):
    """Tìm kiếm thông tin trong DataFrame dựa trên cột và giá trị cần tìm."""
    return df[df[column].astype(str).str.contains(query, case=False, na=False)]

def delete_searched_data(df, column, query, file_path):
    """Xóa tất cả dữ liệu khớp với kết quả tìm kiếm và lưu lại tệp Excel."""
    df = df[~df[column].astype(str).str.contains(query, case=False, na=False)]
    save_to_excel(file_path, df)
    return df

def convert_value(value, dtype):
    """Chuyển đổi giá trị sang kiểu dữ liệu phù hợp"""
    try:
        if pd.api.types.is_integer_dtype(dtype):
            return int(value)
        elif pd.api.types.is_float_dtype(dtype):
            return float(value)
        elif pd.api.types.is_bool_dtype(dtype):
            return str(value).strip().lower() in ['true', '1', 'yes']
        else:
            return str(value)
    except ValueError:
        return value

def update_data_by_column(df, search_column, search_value, update_column, new_value, file_path):
    """Cập nhật dữ liệu trong một cột được chọn dựa trên giá trị tìm kiếm"""
    mask = df[search_column].astype(str).str.contains(str(search_value), na=False)
    if not mask.any():
        st.warning("No matching records found to update.")
        return False, df
    df.loc[mask, update_column] = convert_value(new_value, df[update_column].dtype)
    save_to_excel(file_path, df)
    return True, df

# Giao diện Streamlit
st.title("Excel File Uploader & Editor")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    file_path = "/Users/vietannguyen/datasets/" + uploaded_file.name
    # with open(file_path, "wb") as f:
    #     f.write(uploaded_file.getbuffer())
    
    df = pd.read_excel(file_path)
    st.write("### Current Data in Excel:")
    
    with st.form(key="new_data_form"):
        st.write("### Enter new data:")
        new_data = {}
        for col in df.columns:
            if pd.api.types.is_numeric_dtype(df[col]):
                new_data[col] = st.number_input(f"Enter value for {col}", value=0)
            else:
                new_data[col] = st.text_input(f"Enter value for {col}")

        submit_button = st.form_submit_button(label="Save Data")
    
    if submit_button:
        new_row = pd.DataFrame([new_data])
        df = pd.concat([df, new_row], ignore_index=True)
        save_to_excel(file_path, df)
        st.success("Data saved successfully!")
        st.dataframe(df)
    
    with st.form(key="search_data_form"):
        st.write("### Search Data:")
        search_column = st.selectbox("Select column to search in", df.columns)
        search_query = st.text_input("Enter search query")
        search = st.form_submit_button(label="Search")

    if search:
        result_df = search_data(df, search_column, search_query)
        st.write("### Search Results:")
        st.dataframe(result_df)
        
    if st.button("Delete these Data") and search_column and search_query:
        df = delete_searched_data(df, search_column, search_query, file_path)
        st.success("Selected data deleted successfully!")
    # else:
    #     st.error("Could not delete data")
        
    with st.form(key="update_data_form"):
        st.write("### Update Data by Column:")
        update_column = st.selectbox("Select column to update", df.columns)
        update_value = st.text_input("Enter new value")
        update = st.form_submit_button(label="Update")

    if update:
        status, df = update_data_by_column(df, search_column, search_query, update_column, update_value, file_path)
        if status:
            st.success("Data updated successfully!")
        else:
            st.error("Could not update data")
        
    st.dataframe(df)