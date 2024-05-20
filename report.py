from datetime import datetime
from persiantools.jdatetime import JalaliDate
from openpyxl import Workbook

# تابع برای تبدیل تاریخ میلادی به شمسی
def miladi_to_shamsi(miladi_date):
    date_obj = datetime.strptime(miladi_date, '%Y%m%d')
    jalali_date = JalaliDate.to_jalali(date_obj.year, date_obj.month, date_obj.day)
    return jalali_date

# دریافت تاریخ میلادی از کاربر
start_date_miladi = input("Enter start date (YYYYMMDD): ")
end_date_miladi = input("Enter end date (YYYYMMDD): ")

# تبدیل تاریخ‌های میلادی به شمسی
start_date_jalali = miladi_to_shamsi(start_date_miladi)
end_date_jalali = miladi_to_shamsi(end_date_miladi)

# ایجاد نام فایل شمسی
file_name = f"bsi_sms_Magfa_SmsStatistics_{start_date_jalali.year}{start_date_jalali.month:02d}{start_date_jalali.day:02d}-{end_date_jalali.year}{end_date_jalali.month:02d}{end_date_jalali.day:02d}.xlsx"

# ایجاد workbook و worksheet
workbook = Workbook()
worksheet = workbook.active



# ذخیره فایل
workbook.save(filename=file_name)

print(f"Excel file '{file_name}' has been created successfully.")
