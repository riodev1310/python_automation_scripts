from datetime import datetime

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

# Test hàm với các định dạng khác nhau
test_dates = [
    "31/12/99", "12/31/1999", "31-12-1999", "1999-12-31", "12 Dec 1999",
    "03-04-2024", "4/5/23", "04/03/2024", "1/27/2019", "1/5/19"
]

for date in test_dates:
    print(f"{date} -> {convert_to_ddmmyyyy(date)}")
