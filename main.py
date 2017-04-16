# -*- coding: utf-8 -*-
# !/usr/bin/env python
from __future__ import print_function
from __future__ import unicode_literals
import imaplib
import email
import xlwt


username = '***@gmail.com'
password = '***'
imaplib._MAXLINE = 1000000

gmail = imaplib.IMAP4_SSL('imap.gmail.com', '993')
gmail.login(username, password)
print(gmail.list())
typ, count = gmail.select("INBOX")
print(count)

# !You can replace 'ALL' to '(UNSEEN)'
typ, data = gmail.search(None, "ALL")


style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on', num_format_str='#,##0.00')
wb = xlwt.Workbook()
ws = wb.add_sheet('Received emails')

ws.write(0, 0, "Date", style0)
ws.write(0, 1, "From email", style0)
ws.write(0, 2, "Subject", style0)

j = 1
for i in data[0].split():
    typ, message = gmail.fetch(i, '(RFC822)')
    message = message[0][1]
    print(message)
    print(type(message))

    # message = str(message[0][1])
    mail = email.message_from_bytes(message)
    # date
    try:
        msg_date = mail.get('Date')
        h = email.header.decode_header(msg_date)
        msg_date = h[0][0].decode(h[0][1]) if h[0][1] else h[0][0]
        ws.write(j, 0, str(msg_date))
    except:
        pass

    # sender
    try:
        sender_email = mail.get('From')
        h = email.header.decode_header(sender_email)
        sender_email = h[0][0].decode(h[0][1]) if h[0][1] else h[0][0]
        sender_email = sender_email
        ws.write(j, 1, str(sender_email))
    except:
        pass

    # subject
    try:
        subject = mail.get('Subject')
        h = email.header.decode_header(subject)
        subject = h[0][0].decode(h[0][1]) if h[0][1] else h[0][0]
        subject = subject
        ws.write(j, 2, str(subject))
    except:
        pass
    j += 1

wb.save('gmail.xls')

gmail.close()
gmail.logout()
