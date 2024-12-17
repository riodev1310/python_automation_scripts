import xlwings as xw
import pandas as pd
import seaborn as sns 
import matplotlib.pyplot as plt
from helper_func import *
import os


# Đường dẫn trỏ tới file excel
folder_path = "/Users/vietannguyen/datasets"
chart_path = "./charts"
# Mở file excel
workbook = xw.Book(f"{folder_path}/supermarket_sales.xlsx")

# Truy cập vào sheet đầu tiên có dữ liệu 
sheet = workbook.sheets[0]

# Đọc dữ liệu dùng pandas dataframe
data = pd.read_excel(f"{folder_path}/supermarket_sales.xlsx")
# data["Date"] = data["Date"].astype(str)
# data["Date"] = data["Date"].apply(convert_to_ddmmyyyy)

# Tạo pivot table
transaction_by_branch = create_pivot_table(data, "count","Invoice ID", "Customer type")

pivot_sheet = overwrite_sheet(f"{folder_path}/supermarket_sales.xlsx", "Number of Invoice ID by Branch", transaction_by_branch)


# Vẽ biểu đồ
agg_data = data.groupby('Branch').count()

column_chart = draw_bar_chart(agg_data, agg_data.index, agg_data["Invoice ID"])

pivot_sheet.pictures.add(
    column_chart,
    name="Seaborn1",
    update=True,
    left=pivot_sheet.range("F1").left,
    top=pivot_sheet.range("F1").top,
    # height=200,
    # width=300,
)


gender_distributiions = create_pivot_table(data, "count", "Invoice ID", "Gender")
pivot_sheet = overwrite_sheet(f"{folder_path}/supermarket_sales.xlsx", "Gender Distribution", gender_distributiions)

gender_counts = data["Gender"].value_counts()

pie_chart = draw_pie_chart(gender_counts, gender_counts.index, "Gender Distributions")

pivot_sheet.pictures.add(
    pie_chart,
    name="Seaborn1",
    update=True,
    left=pivot_sheet.range("F1").left,
    top=pivot_sheet.range("F1").top,
    # height=200,
    # width=300,
)

customer_type = create_pivot_table(data, "count", "Invoice ID", "Customer type", "Branch")
pivot_sheet = overwrite_sheet(f"{folder_path}/supermarket_sales.xlsx", "Customer Types", customer_type)
# cluster_column_chart = draw_bar_chart(data, data["Branch"], data["Invoice ID"], data["Customer type"])

plt.figure(figsize=(12, 8))
customer_type_counts = data.groupby(['Branch', 'Customer type']).size().unstack()
customer_type_counts.plot(kind="bar")

plt.xlabel('Branch')
plt.ylabel('Number of Customers')
plt.xticks(rotation=0)
plt.title('Customer Type per branch')
plt.legend(title="Customer Type")
plt.savefig(f"{chart_path}/cluster.png")

# cluster_column_chart = plt.figure(figsize=(12, 8))

pivot_sheet.pictures.add(
    f"{chart_path}/cluster.png",
    name="Seaborn1",
    update=True,
    left=pivot_sheet.range("F1").left,
    top=pivot_sheet.range("F1").top,
    # height=200,
    # width=300,
)

workbook.save(f"{folder_path}/supermarket_sales_report.xlsx")
# workbook.close()

for filename in os.listdir(chart_path):
    if filename.endswith(('.png', '.jpg', '.jpeg')):
        remove_chart(f"{chart_path}/{filename}")
# workbook.save(f"{folder_path}/supermarket_sales_report.xlsx")