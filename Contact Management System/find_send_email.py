import openpyxl
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime, timedelta
import random

birthday_wishes = [
    "祝您生日快乐，岁岁平安！",
    "在这特别的日子里，愿您的笑容如阳光般灿烂！",
    "生日快乐！愿您的每一天都充满幸福和喜悦。",
    "愿这美妙的一天带给您无尽的快乐和美好的回忆！",
    "又长大一岁了，祝您越来越年轻，生日快乐！",
    "愿您的生日充满无尽的快乐，愿您的每天都闪耀着快乐的阳光。",
    "在您的生日之际，向您致以最美好的祝愿！",
    "愿您在这特别的日子里，心想事成，万事如意！"
]

def get_tomorrows_birthdays(filename):
    """
    Retrieves contacts' birthdays that match with tomorrow's date from the given Excel file.

    :param filename: Path to the Excel file
    :return: A list of tuples containing the email and name of contacts whose birthday is tomorrow
    """
    tomorrow = datetime.now().date() + timedelta(days=1)
    wb = openpyxl.load_workbook(filename)
    sheet = wb['Contacts']
    birthdays_tomorrow = []

    for row in sheet.iter_rows(min_row=2, values_only=True):
        name = row[0]
        email = row[3]

        # 只有当生日字段不为空时，才处理生日数据
        if row[1] != None:
            birthday = datetime.strptime(row[1], '%Y-%m-%d').date()

            # Check if the month and day match tomorrow's date
            if birthday.month == tomorrow.month and birthday.day == tomorrow.day:
                birthdays_tomorrow.append((email, name))

    return birthdays_tomorrow

def send_birthday_email(receiver_email, birthday_person, birthday_message):
    # 你的邮箱地址和密码
    sender_email = '2524533592@qq.com'
    sender_password = 'uechhrpzqfqfdhjd'
    smtp_server = 'smtp.qq.com'
    smtp_port = 587

    # 构建邮件内容，使用传入的祝福语
    subject = '生日快乐！'
    body = f'亲爱的 {birthday_person}，{birthday_message}'

    message = MIMEMultipart()
    message['From'] = sender_email
    message['To'] = receiver_email
    message['Subject'] = subject
    message.attach(MIMEText(body, 'plain'))

    # 建立SMTP连接并发送邮件
    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, receiver_email, message.as_string())
        server.quit()
        print(f'生日祝福邮件已发送至 {receiver_email}')
    except Exception as e:
        print(f'邮件发送失败：{str(e)}')

def main():
    filename = "contacts.xlsx"
    birthdays_tomorrow = get_tomorrows_birthdays(filename)

    for email, name in birthdays_tomorrow:
        send_birthday_email(email, name)

if __name__ == "__main__":
    main()
