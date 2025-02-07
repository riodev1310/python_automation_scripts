import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# -------------------- Cấu hình Google Sheets API -------------------- #
# Load credentials từ file JSON (tải về từ Google Cloud)
creds = Credentials.from_service_account_file("google_sheet_key.json", scopes=["https://www.googleapis.com/auth/spreadsheets"])
# Kết nối Google Sheets
client = gspread.authorize(creds)

# ID của Google Sheet (lấy từ URL của Google Sheets)
# SHEET_ID = "1tZWdCq2cWmIZHTmD3_ETiwGr_0dOrjBRRwa5kiaaCts"  # Thay bằng ID thực tế của bạn
# SHEET_NAME = "Employees"


def get_sheet(SHEET_ID, SHEET_NAME):
    """Mở Google Sheet và lấy dữ liệu dưới dạng DataFrame"""
    sheet = client.open_by_key(SHEET_ID).worksheet(SHEET_NAME)
    data = sheet.get_all_records()
    return pd.DataFrame(data)


def save_to_sheet(df, SHEET_ID, SHEET_NAME):
    """Lưu DataFrame vào Google Sheet"""
    sheet = client.open_by_key(SHEET_ID).worksheet(SHEET_NAME)
    sheet.clear()  # Xóa dữ liệu cũ
    sheet.update([df.columns.values.tolist()] + df.values.tolist())  # Cập nhật dữ liệu mới


def search_data(df, column, query):
    """Tìm kiếm thông tin trong DataFrame dựa trên cột và giá trị cần tìm"""
    return df[df[column].astype(str).str.contains(query, case=False, na=False)]


def delete_searched_data(df, column, query, SHEET_ID, SHEET_NAME):
    """Xóa tất cả dữ liệu khớp với kết quả tìm kiếm và lưu lại Google Sheet"""
    df = df[~df[column].astype(str).str.contains(query, case=False, na=False)]
    save_to_sheet(df, SHEET_ID, SHEET_NAME)
    return df

# -------------------- Giao diện Streamlit -------------------- #
st.title("Google Sheets Editor with Streamlit")

SHEET_ID = st.text_input("Enter sheet's ID: ")
SHEET_NAME = st.text_input("Enter sheet's name: ")

# Tải dữ liệu từ Google Sheets
if SHEET_ID and SHEET_NAME:
    df = get_sheet(SHEET_ID, SHEET_NAME)

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
        save_to_sheet(df, SHEET_ID, SHEET_NAME)
        st.success("Data saved successfully!")
        # st.dataframe(df)



    st.write("### Search Data:")
    search_column = st.selectbox("Select column to search in", df.columns)
    search_query = st.text_input("Enter search query")
    if st.button("Search"):
        result_df = search_data(df, search_column, search_query)
        st.write("### Search Results:")
        st.dataframe(result_df)


    if st.button("Delete these Data"):
        df = delete_searched_data(df, search_column, search_query, SHEET_ID, SHEET_NAME)
        st.success("Selected data deleted successfully!")
        # st.dataframe(df)
    
    st.write("### Current Data in Google Sheet:")
    st.dataframe(df)