import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

import InvoiceAutomation

body1 ='''Dear Hardeep,

Please find attached the invoice for 
'''
body2 =InvoiceAutomation.RentalPeriod

body3 = ''' 

Payment is '''

body33= InvoiceAutomation.DueBy

body4 ='''

For any query, I remain.

Regards,
Kazi Toufiq Wadud
0455455171

'''

mail_content = body1 + body2 + body3 + body33+  body4


#The mail addresses and password
sender_address = 'kazitoufiqw@gmail.com'
sender_pass = 'xxxx'
receiver_address = 'yyyyy@gmail.com'

#Setup the MIME

message = MIMEMultipart()
message['From'] = sender_address
message['To'] = receiver_address
message['Subject'] = 'Van (Reg: 1OC9KC) Rental ' +  InvoiceAutomation.NewInvoiceNo



message.attach(MIMEText(mail_content, 'plain'))

attach_file_name = InvoiceAutomation.new_invoice_pdf_file

attach_file = open(attach_file_name, 'rb') # Open the file as binary mode

payload = MIMEBase('application', 'octate-stream')

payload.set_payload((attach_file).read())

encoders.encode_base64(payload) #encode the attachment

#add payload header with filename  #Content-Decomposition

payload.add_header('Content-Disposition', "attachment; filename= "+attach_file_name)

message.attach(payload)

#Create SMTP session for sending the mail

session = smtplib.SMTP('smtp.gmail.com', 587)
session.starttls() #enable security
session.login(sender_address, sender_pass)
text = message.as_string()
session.sendmail(sender_address, receiver_address, text)
session.quit()

print(InvoiceAutomation.RentalPeriod)

print('Mail Sent')