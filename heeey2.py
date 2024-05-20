import cx_Oracle
import pandas as pd

# اطلاعات اتصال به دیتابیس
username = "hamed"
password = "hamed"
dsn = "orcl"  # نام دیتابیس مانند 'localhost/XEPDB1' در Oracle
host = "localhost"  # می‌تواند 'localhost' یا هاست دیگری باشد
port = "1521"  # پورت مانند '1521' در Oracle

# اتصال به دیتابیس
connection = cx_Oracle.connect(user=username, password=password, dsn=dsn, encoding="UTF-8")

# ایجاد یک کرسور برای اجرای کوئری‌ها
cursor = connection.cursor()

# درخواست تاریخ شروع از کاربر
start_date = input("لطفا تاریخ شروع (به فرمت YYYY-MM-DD) را وارد کنید: ")

# درخواست تاریخ پایان از کاربر
end_date = input("لطفا تاریخ پایان (به فرمت YYYY-MM-DD) را وارد کنید: ")

# کوئری اجرایی
query = """
SELECT
    from_mobile_number AS "سرخط",
    CASE
        WHEN operator = 'mci' THEN 'همراه اول'
        WHEN operator = 'samantel' THEN 'سامانتل'
        WHEN operator = 'irancell' THEN 'ایرانسل'
        WHEN operator = 'rightel' THEN 'رایتل'
        ELSE operator
    END AS "اپراتور",
    'فارسی' AS "نوع پیام",
    SUM(CASE WHEN status = 'SMSC_MESSAGE_ACCEPTED' THEN farsi ELSE 0 END) AS "Accepted",
    SUM(CASE WHEN status = 'SMSC_MESSAGE_REJECTED' THEN farsi ELSE 0 END) AS "Rejected",
    SUM(CASE WHEN status = 'SMSC_MESSAGE_DELIVERED' THEN farsi ELSE 0 END) AS "Delivered",
    SUM(CASE WHEN status = 'SMSC_MESSAGE_UNDELIVERABLE' THEN farsi ELSE 0 END) AS "Undelivered",
    SUM(farsi) AS "جمع کل"
FROM
    report
WHERE
    creation_date >= :start_date
    AND creation_date < :end_date
    AND operator IS NOT NULL
GROUP BY
    from_mobile_number,
    CASE
        WHEN operator = 'mci' THEN 'همراه اول'
        WHEN operator = 'samantel' THEN 'سامانتل'
        WHEN operator = 'irancell' THEN 'ایرانسل'
        WHEN operator = 'rightel' THEN 'رایتل'
        ELSE operator
    END
UNION
SELECT
    from_mobile_number AS "سرخط",
    CASE
        WHEN operator = 'mci' THEN 'همراه اول'
        WHEN operator = 'samantel' THEN 'سامانتل'
        WHEN operator = 'irancell' THEN 'ایرانسل'
        WHEN operator = 'rightel' THEN 'رایتل'
        ELSE operator
    END AS "اپراتور",
    'غیر فارسی' AS "نوع پیام",
    SUM(CASE WHEN status = 'SMSC_MESSAGE_ACCEPTED' THEN latin ELSE 0 END) AS "Accepted",
    SUM(CASE WHEN status = 'SMSC_MESSAGE_REJECTED' THEN latin ELSE 0 END) AS "Rejected",
    SUM(CASE WHEN status = 'SMSC_MESSAGE_DELIVERED' THEN latin ELSE 0 END) AS "Delivered",
    SUM(CASE WHEN status = 'SMSC_MESSAGE_UNDELIVERABLE' THEN latin ELSE 0 END) AS "Undelivered",
    SUM(latin) AS "جمع کل"
FROM
    report
WHERE
    creation_date >= :start_date
    AND creation_date < :end_date
    AND operator IS NOT NULL
GROUP BY
    from_mobile_number,
    CASE
        WHEN operator = 'mci' THEN 'همراه اول'
        WHEN operator = 'samantel' THEN 'سامانتل'
        WHEN operator = 'irancell' THEN 'ایرانسل'
        WHEN operator = 'rightel' THEN 'رایتل'
        ELSE operator
    END
"""

# پارامترهایی که برای کوئری مورد نیاز است
params = {
    'start_date': start_date,
    'end_date': end_date
}

# اجرای کوئری با استفاده از پارامترها
cursor.execute(query, params)

# گرفتن نتیجه کوئری
result = cursor.fetchall()

# تبدیل نتیجه به DataFrame
df = pd.DataFrame(result, columns=["سرخط", "اپراتور", "نوع پیام", "Accepted", "Rejected", "Delivered", "Undelivered", "جمع کل"])

# ذخیره DataFrame در فایل Excel
file_name = "output.xlsx"
df.to_excel(file_name, index=False)

# بستن کرسور و اتصال
cursor.close()
connection.close()

print("نتیجه با موفقیت در فایل Excel ذخیره شد.")
