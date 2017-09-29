#!/usr/bin/env python

from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate
from email import encoders
from os.path import basename
import openpyxl # untuk kerja dengna excel
import datetime # untuk kerja dengan waktu
import requests
import smtplib # untuk kerja dngan email
import pymssql # untuk kerja dengan mssql server
import json

COLS = 4

def ambil_dan_taro_data():
	conn = pymssql.connect('<server>', '<username>', '<password>', '<db>')
	cursor = conn.cursor()

	date_now = (datetime.datetime.now() - datetime.timedelta(days = 1)).strftime("%d")
	month_now = (datetime.datetime.now() - datetime.timedelta(days = 1)).strftime("%B")
	year_now = (datetime.datetime.now() - datetime.timedelta(days = 1)).strftime("%Y")
	date_7_before = (datetime.datetime.now() - datetime.timedelta(days = 7)).strftime("%d")
	month_7_before = (datetime.datetime.now() - datetime.timedelta(days = 7)).strftime("%B")
	year_7_before = (datetime.datetime.now() - datetime.timedelta(days = 7)).strftime("%Y")

	file_name = ('%s %s %s - %s %s %s.xlsx' % (date_7_before, month_7_before, year_7_before, date_now, month_now, year_now))
	col_name = "a | b | c | d"

	wb = openpyxl.Workbook()
	worksheet = wb.get_sheet_by_name('Sheet')

	x = ord('A')
	for i in col_name.split('|'):
		worksheet[chr(x) + '1'] = i
		x += 1

	query = """SELECT * 
		FROM blah
		WHERE blah = blahh
	"""
	cursor.execute(query)

	x = 2
	row = cursor.fetchone()
	while row:
		column = ord('A')
		for i in range(COLS):
			worksheet[chr(column) + str(x)] = row[i]
			column += 1
		x += 1
		row = cursor.fetchone()

	for col in worksheet.columns:
		max_length = 0
		column = col[0].column # Get the column name
		for cell in col:
			try: # Necessary to avoid error on empty cells
				if len(str(cell.value)) > max_length:
					max_length = len(cell.value)
			except:
				pass
			
			cell.alignment = openpyxl.styles.Alignment(horizontal = 'center')

		adjusted_width = (max_length + 2) * 1.2
		worksheet.column_dimensions[column].width = adjusted_width

	wb.save(file_name)
	return file_name


def kirim_email(file_name):
	smtpobj = smtplib.SMTP('<smtp server>', <port>)
	smtpobj.ehlo()
	smtpobj.starttls()
	smtpobj.login('<user>', '<pass>')

	DEBUG = False

	if DEBUG:
		sender = 'a@a'
		recipients = ['b@b']
	else:
		sender = 'a@a'
		recipients = ['c@c']

	text = ("""
Free text
	"""% file_name.split('.')[0])

	msg = MIMEMultipart()
	msg['From'] = sender
	msg['To'] = ",".join(recipients)
	msg['Date'] = formatdate(localtime = True)
	msg['Subject'] = "subject"

	msg.attach(MIMEText(text))
	part = MIMEApplication(open(file_name).read(), Name = basename(file_name))
	part['Content-Disposition'] = 'attachment; filename="%s"' % basename(file_name)
	msg.attach(part)

	smtpobj.sendmail(sender, recipients, msg.as_string())
	smtpobj.quit()

def send_alert(pesan):
	url = 'https://api.telegram.org/bot<token>/sendMessage?chat_id=<id>&parse_mode=markdown&text=%s' % pesan
	requests.get(url)

if __name__ == '__main__':
	send_alert("*Mulai*")

	try:
		file_name = ambil_dan_taro_data()
	except Exception as e:
		pesan = "*Error saat ambil data di database dan masukkin ke Excel*\n\n"
		pesan += str(e)
		send_alert(pesan)
		exit(1)
	try:
		kirim_email(file_name)
	except Exception as e:
		pesan = "*Error saat kirim email*\n\n"
		pesan += str(e)
		send_alert(pesan)
		exit(1)

	send_alert("*Berhasil dengan sempurna*")