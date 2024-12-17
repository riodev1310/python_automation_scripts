import pandas as pd
import xlwings as xw


folder_path = "/Users/vietannguyen/datasets"
workbook = xw.Book(f"{folder_path}/preprocessing_data.xlsx")

sheet = workbook.sheets[1]

df = pd.read_csv(f"{folder_path}/preprocessing_data.csv")

# Truy cập vào sheet đầu tiên có dữ liệu 
sheet = workbook.sheets["Raw data"]

category_avg = df.groupby("Category")["Purchase Amount (USD)"].transform("mean")
df["Purchase Amount (USD)"].fillna(round(category_avg, 2), inplace=True)

mean_purchase = df["Review Rating"].mean()

df["Review Rating"].fillna(round(mean_purchase, 2), inplace=True)

sheet.range("A1").value = df

workbook.save(f"{folder_path}/clean_dataset.xlsx")
# workbook.close()

print("Data processing complete and saved to Excel.")