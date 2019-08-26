import smtplib
import mimetypes
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from time import sleep



class Mail():
	def __init__(self):
		self.passowrd = ''
		self.send_from = ""
		self.send_to = ""
		self.message = ""
		self.subject = "excel"
		self.filename = ""
		self.file_path = ""

	def send(self):
		msg = MIMEMultipart()
		msg["Subject"] = self.subject
		msg["From"] = self.send_from
		msg["To"] = self.send_to
		msg.preamble = "Excel"

		file = self.file_path + "/" + self.filename 
		
		attachment = open(file,'rb')
		fp = open(file,'rb')
		# xls = MIMEBase('application','vnd.ms-excel')
		# xls.set_payload(fp.read())
		# fp.close()
		# encoders.encode_base64(xls)
		# xls.add_header('Content_Disposition','attachment', filename = self.filename)
		# msg.attach(xls)
		print("Sending mail has been initiated.....")
		xlsx = MIMEBase('application','vnd.openxmlformats-officedocument.spreadsheetml.sheet')
		xlsx.set_payload(attachment.read())
		attachment.close()
		encoders.encode_base64(xlsx)
		xlsx.add_header('Content-Disposition', 'attachment', filename = self.filename)
		msg.attach(xlsx)


		# print(msg)
		conn  = smtplib.SMTP('smtp.gmail.com',587)
		print("Mail server has been initiated....!\n")
		conn.ehlo()
		conn.starttls()
		print("TLS connection has been initiated......!\n")
		print("Initating Login..............!\n")
		conn.login(self.send_from, self.passowrd)
		print("LOGIN Successful................!\n")
		conn.sendmail(self.send_from, self.send_to , msg.as_string())
		print("Mail has been sent to {0}".format(self.send_to))
		print("LOGGED OUT of the mail.\n")
		conn.quit()
		sleep(1)
		print("TLS conncetion has been terminated.\n")
		sleep(1)
		print("Mail server has been closed.\n")


email = Mail()
email.send()