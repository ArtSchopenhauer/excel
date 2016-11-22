from flask import Flask, render_template, request
import json
import requests
import datetime
import pytz
from flask_mail import Mail, Message

app = Flask(__name__)

app.config.update(
	DEBUG = True,
	MAIL_SERVER = 'smtp.gmail.com',
	MAIL_PORT = 465,
	MAIL_USE_SSL = True,
	MAIL_USE_TLS = False,
	MAIL_USERNAME = 'zack.gray@levelsolar.com',
	MAIL_PASSWORD = 'levelsolar'
	)

mail = Mail(app)

recipients = ['zspencergray@gmail.com']

zack = ['zspencergray@gmail.com']

def send_mail():		
	date = str(datetime.datetime.now().month) + "-" + str(datetime.datetime.now().day) + "-" + str(datetime.datetime.now().year)
	date_subj = str(datetime.datetime.now().month) + "/" + str(datetime.datetime.now().day) + "/" + str(datetime.datetime.now().year)
	body = "<p>Hi,</p><p>Please find today's Sales Report attached.</p><p>Thanks,</p><p>Zack</p>"
	msg = Message("Sales Report %s" %date_subj, sender='Zack Gray', recipients=recipients)
	msg.attach("Sales Report %s.xlsx" %date, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", file.read())
	msg.html = body
	mail.send(msg)
	return "Done."

def send_error():
	msg = Message('Sales Report Error', sender='Zack Gray', recipients=zack)
	mail.send(msg)
	return "Done."

if __name__ == "__main__":
	data_file = open("/root/lboard/report_check.json", "r")
	day = json.load(data_file)
	data_file.close()
	if datetime.datetime.now().day == day:
		with app.app_context():
			file = app.open_resource("Sales Report.xlsx")
			send_mail()
			file.close()
	else:
		with app.app_context():
			file = app.open_resource("Sales Report.xlsx")
			send_error()