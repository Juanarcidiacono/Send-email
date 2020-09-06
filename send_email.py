'''
From an excel file I send an email to those they owe money with a personalize message
'''

import xlrd
import time
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# This excel files has information about clients. (Name,email, city, paid, ammount)
path = "clients_status.xlsx"
openFile = xlrd.open_workbook(path)
sheet = openFile.sheet_by_name('datos')

# I put the email, amount and name of those who owes money in three different lists.
mail_list = []
amount = []
name = []
for k in range(sheet.nrows-1):
    client = sheet.cell_value(k+1,0)
    email = sheet.cell_value(k+1,1)
    abon = sheet.cell_value(k+1,3)
    cant_amount = sheet.cell_value(k+1,4)
    if abon == 'No':
        mail_list.append(email) 
        amount.append(cant_amount)
        name.append(client)

# Start sending email process
email = 'myEmail@gmail.com' # sender email
password = 'none' # This is an app password so you dont have to put the original password. It is in the security section in Gmail/Apps passwords
server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login(email, password)

for mail_to in mail_list:
    send_to_email = mail_to
    nro_des = mail_list.index(send_to_email) # get the index so then I can find the name of the person 
    clientName = name[nro_des] 
    subject = f'{clientName} you have a new email'
    message = f'Dear {clientName}: \n' \
              f'We inform you that you owe ${amount[nro_des]} \n'\
              '\n' \
              'Regards' 
              
    msg = MIMEMultipart()
    msg['From '] = send_to_email
    msg['Subject'] = subject
    msg.attach(MIMEText(message, 'plain'))
    text = msg.as_string()
    print(f'Sending email to {clientName}... ') # With this I can know to who is sending the email
    server.sendmail(email, send_to_email, text)
    
server.quit()
print('Process is finished!')
time.sleep(10) # Take to seconds to see the comand window to be sure everything it's ok.
