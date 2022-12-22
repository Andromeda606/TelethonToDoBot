from email.mime.text import MIMEText

from openpyxl import load_workbook
import datetime
import smtplib, ssl

smtp_server = "smtp.gmail.com"
port = 587
sender_email = ""
password = ""
fmt = 'From: {}\r\nTo: {}\r\nSubject: {}\r\n{}'

context = ssl.create_default_context()

email = smtplib.SMTP(smtp_server, port)
email.ehlo()  # Can be omitted
email.starttls(context=context)  # Secure the connection
email.ehlo()  # Can be omitted
email.login(sender_email, password)

yarin = (datetime.datetime.now() + datetime.timedelta(days=1)).day
excel_dosyasi = load_workbook("sample.xlsx")

sayfa1 = excel_dosyasi.active
mailler = []
for mail in sayfa1["F"]:
    if mail.value is None:
        break
    if mail.value == "Mail atılacak kişiler":
        continue
    mailler.append(mail.value)
print(mailler)

for a, b in zip(sayfa1["A"], sayfa1["B"]):
    if a.value == yarin or a.value == yarin + 1 or a.value == yarin + 2:
        print("-------------------------------------Yarın Yapılacaklar-----------------------------------")
        for yapilacak in b.value.split("***"):
            for mail in mailler:
                msg = MIMEText(yapilacak.strip())
                msg['Subject'] = 'Yapılacak'
                msg['From'] = sender_email
                msg['To'] = mail
                email.sendmail(sender_email, mail, msg.as_string())
