import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# -------------------- Cấu hình Google Sheets API -------------------- #
# Load credentials từ file JSON (tải về từ Google Cloud)
creds = Credentials.from_service_account_file("google_sheet_key.json", scopes=["https://www.googleapis.com/auth/spreadsheets"])
# Kết nối Google Sheets
client = gspread.authorize(creds)


def get_sheet(SHEET_ID, SHEET_NAME):
    """Mở Google Sheet và lấy dữ liệu dưới dạng DataFrame"""
    sheet = client.open_by_key(SHEET_ID).worksheet(SHEET_NAME)
    data = sheet.get_all_records()
    return pd.DataFrame(data) if data else pd.DataFrame()


def save_to_sheet(df, SHEET_ID, SHEET_NAME):
    """Lưu DataFrame vào Google Sheet"""
    sheet = client.open_by_key(SHEET_ID).worksheet(SHEET_NAME)
    sheet.clear()  # Xóa dữ liệu cũ
    sheet.update([df.columns.values.tolist()] + df.values.tolist())  # Cập nhật dữ liệu mới


def search_data(df, column, query):
    """Tìm kiếm thông tin trong DataFrame dựa trên cột và giá trị cần tìm"""
    return df[df[column].astype(str).str.contains(str(query), case=False, na=False)]


def delete_searched_data(df, column, query, SHEET_ID, SHEET_NAME):
    """Xóa tất cả dữ liệu khớp với kết quả tìm kiếm và lưu lại Google Sheet"""
    df = df[~df[column].astype(str).str.contains(str(query), case=False, na=False)]
    save_to_sheet(df, SHEET_ID, SHEET_NAME)
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


def update_data_by_column(df, search_column, search_value, update_column, new_value, SHEET_ID, SHEET_NAME):
    """Cập nhật dữ liệu trong một cột được chọn dựa trên giá trị tìm kiếm"""
    mask = df[search_column].astype(str) == str(search_value)
    if not mask.any():
        st.warning("No matching records found to update.")
        return False, df
    df.loc[mask, update_column] = convert_value(new_value, df[update_column].dtype)
    save_to_sheet(df, SHEET_ID, SHEET_NAME)
    return True, df

# -------------------- Giao diện Streamlit -------------------- #
st.title("Google Sheets Editor with Streamlit")

SHEET_ID = st.text_input("Enter sheet's ID: ")
SHEET_NAME = st.text_input("Enter sheet's name: ")

# Tải dữ liệu từ Google Sheets
if SHEET_ID and SHEET_NAME:
    df = get_sheet(SHEET_ID, SHEET_NAME)
    if df.empty:
        st.warning("No data found in Google Sheet.")
    else:
        # Form nhập dữ liệu mới
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

        # Tìm kiếm dữ liệu
        st.write("### Search Data:")
        search_column = st.selectbox("Select column to search in", df.columns)
        search_query = st.text_input("Enter search query")

        if st.button("Search"):
            result_df = search_data(df, search_column, search_query)
            st.write("### Search Results:")
            st.dataframe(result_df)

        # Cập nhật dữ liệu theo cột được chọn
        st.write("### Update Data by Column:")
        update_column = st.selectbox("Select column to update", df.columns)
        update_value = st.text_input("Enter new value")

        if st.button("Update Data"):
            status, df = update_data_by_column(df, search_column, search_query, update_column, update_value, SHEET_ID, SHEET_NAME)
            if status:
                st.success("Data updated successfully!")
            else:
                st.error("Could not update data")

        # Xóa dữ liệu với xác nhận trước
        if st.button("Delete these Data"):
            if st.confirm("Are you sure you want to delete matching data?"):
                df = delete_searched_data(df, search_column, search_query, SHEET_ID, SHEET_NAME)
                st.success("Selected data deleted successfully!")

        # Hiển thị dữ liệu hiện tại
        st.write("### Current Data in Google Sheet:")
        st.dataframe(df)