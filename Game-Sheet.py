import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
from PyPDF2 import PdfFileMerger
import glob
import os


print ("Please enter the following:")
game = input("Game #: ")
team1 = input("Team No.1: ")
team2 = input("Team No.2: ")

title = "Game #" + game + " " + team1 + " vs " + team2
title_with_extension = title + ".pdf"

#collate = input("Do you have a pdf to collate? (Yes or No) ")

#if collate.lower() == 'yes':
list_of_files = glob.glob('D:/Matthew/Desktop/Refereeing/*.pdf') # * means all if need specific format then *.csv
latest_file =sorted(list_of_files, key=os.path.getctime)

pdfs = [latest_file[-1], latest_file[-2]]

merger = PdfFileMerger()

for pdf in pdfs:
    merger.append(pdf)

mergedFile = merger.write("D:/Matthew/Desktop/Refereeing/%s" % title_with_extension)
merger.close()

if os.path.exists(latest_file[-1]): os.remove(latest_file[-1])
if os.path.exists(latest_file[-2]): os.remove(latest_file[-2])

gmail_user = 'matthew17berger@gmail.com'
gmail_password = input("Type your email password and press enter: ")

msg = MIMEMultipart()
msg['Subject'] = title
msg['From'] = gmail_user
msg['To'] = 'bergermaestro@gmail.com'

body = 'Hi There,\n\nAttached are the Game Sheets for ' + title + '.\n\nThanks,\n\nMatthew Berger\nOSA#3955026'
msg.attach(MIMEText(body, 'plain'))

part = MIMEBase('application', "octet-stream")
part.set_payload(open('D:/Matthew/Desktop/Refereeing/'+title_with_extension, "rb").read())
encoders.encode_base64(part)
part.add_header('Content-Disposition', 'attachment; filename= %s' % title_with_extension)
msg.attach(part)


try:
    server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
    server.ehlo()
    server.login(gmail_user, gmail_password)
    server.sendmail(msg['From'], msg['To'], msg.as_string())
    server.close()

    print ('Email sent!')
except:
    print ('Something went wrong...')
