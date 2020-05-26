import pandas as pd

data = pd.read_excel(r'C:\Users\Angu312\Documents\Calibration_Statuses.xlsx', "Tracking")
status = data['Status'].values.tolist()

def send_mail(recipient, subject, message):

    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.base import MIMEBase
    from email.mime.text import MIMEText
    from email import encoders

    username = "USERNAME"
    password = "PASSWORD"

    msg = MIMEMultipart()
    msg['From'] = username
    msg['To'] = recipient
    msg['Subject'] = subject
    msg.attach(MIMEText(message))

    part = MIMEBase('application', "octet-stream")
    part.set_payload(open(r'C:\Users\Angu312\Documents\Calibration_Statuses.xlsx', "rb").read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment; filename="test.xlsx"')
    msg.attach(part)

    try:
        print('sending mail to ' + recipient + ' on ' + subject)

        mailServer = smtplib.SMTP('smtp-mail.outlook.com', 587)
        mailServer.ehlo()
        mailServer.starttls()
        mailServer.ehlo()
        mailServer.login(username, password)
        mailServer.sendmail(username, recipient, msg.as_string())
        mailServer.close()

    except error as e:
        print(str(e))

Email_Recipient = "RECIPIENT"
Subject_Line = 'IMPORTANT! Please review the Measurement Devices Tracker!'
Bad_Message = 'There are instruments listed in the attached document that should be re-calibrated or scheduled to be re-calibrated soon!'

for i in status:
    if i == 'BAD':
        send_mail(Email_Recipient, Subject_Line, Bad_Message)
        break
    else:
        print('Hello')
