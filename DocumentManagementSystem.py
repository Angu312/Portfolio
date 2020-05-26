import os
import shutil
import datetime as dt
import pandas as pd

data = pd.read_excel(r'A:\Users\Angu312\Process Owners List.xlsx', "Main")
keys = data['Process'].values.tolist()
values = data['Email'].values.tolist()
processes = data['Process Folders'].values.tolist()
dictionary = dict(zip(keys, values))

# "Retention Code" function
# This section defines a function to distinguish files by their retention codes. position1 finds the first underscore from the right in the
# string, while position2 finds the period in the string (file name). The function returns a string with solely the file's retention code.
def RetentionCode(string):
    position1 = string.rfind('_')
    position2 = string.rfind(' -')
    return string[position1+1:position2]

def ProcessID(string):
    position1 = string.rfind('US_')
    position2 = string.rfind('_')
    return string[position1--6:position2-4]

def FileName(path):
    return os.path.basename(path)

def send_email(email_recipient, email_subject, email_message):
    import smtplib
    from email.mime.text import MIMEText
    from email.mime.multipart import MIMEMultipart

    email_sender = "EMAIL"
    password = "PASSWORD"

    msg = MIMEMultipart()
    msg['From'] = email_sender
    msg['To'] = email_recipient
    msg['Subject'] = email_subject
    msg.attach(MIMEText(email_message))

    print('sending mail to ' + email_recipient + ' on ' + email_subject)

    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login(email_sender, password)
    server.sendmail(email_sender, email_recipient, msg.as_string())
    print('Email sent!')
    server.quit()

for folders in processes:
    # This section establishes a for loop which calls on the Retention Code function to delete/move files based on their retention codes.
    for files in os.listdir(folders):
        Retention_Code = RetentionCode(files)
        src_file_path = os.path.join(folders, files)
        dest_file_path = os.path.join(folders, files)
        Current_Date = dt.datetime.today().date()
        Creation_Date = dt.datetime.fromtimestamp(os.path.getctime(src_file_path)).date()
        Modification_Date = dt.datetime.fromtimestamp(os.path.getmtime(src_file_path)).date()
        Thirty_Days = dt.timedelta(days=30)
        Sixty_Days = dt.timedelta(days=60)
        One_Year = dt.timedelta(days=365)
        Three_Years = dt.timedelta(days=1096)
        Five_Years = dt.timedelta(days=1826)
        Ten_Years = dt.timedelta(days=3652)
        Fifteen_Years = dt.timedelta(days=5479)
        Twenty_Years = dt.timedelta(days=7305)

        File_Name = FileName(files)
        Subject_Line = 'IMPORTANT! Please review your Process documents under the Quality folder!'
        Thirty_Days_Message = 'The following document under your Process is subject for deletion in 30 days and needs to be reviewed per the Document Management Policy: ' + File_Name
        Sixty_Days_Message = 'The following document under your Process is subject for deletion in 60 days and needs to be reviewed per the Document Management Policy: ' + File_Name
        One_Year_Message = 'This is a message from the company's automated ISO 9001 document management system. The following document under your Process has not been edited in over one year, is subject for deletion, and needs to be reviewed per the Document Management Policy: ' + File_Name

        if Retention_Code == 'RA':
            if Current_Date - Modification_Date == Thirty_Days:
                # Extract the full name of the file and store it in a variable
                File_Name = FileName(files)
                # Extract the Process ID from the file name
                Process_ID = ProcessID(files)
                # Use the Process ID from the file name and match to the Key in Dictionary
                for key in dictionary:
                    if key in Process_ID:
                        Email_Recipient = dictionary[key]
                        # Using the Key from Dictionary, send an email to person in the corresponding Value in Dictionary with the full file
                        send_email(Email_Recipient, Subject_Line, Thirty_Days_Message)
            elif Current_Date - Modification_Date == Sixty_Days:
                File_Name = FileName(files)
                Process_ID = ProcessID(files)
                for key in dictionary:
                    if key in Process_ID:
                        Email_Recipient = dictionary[key]
                        send_email(Email_Recipient, Subject_Line, Sixty_Days_Message)
            elif Current_Date - Modification_Date >= One_Year:
                File_Name = FileName(files)
                Process_ID = ProcessID(files)
                for key in dictionary:
                    if key in Process_ID:
                        Email_Recipient = dictionary[key]
                        send_email(Email_Recipient, Subject_Line, One_Year_Message)
                # if os.path.exists(dest_file_path):
                #     os.remove(dest_file_path)
                #     shutil.move(src_file_path, Destination)
                # else:
                #     shutil.move(src_file_path, Destination)

        elif Retention_Code == 'R1':
            if Current_Date >= (Creation_Date + One_Year) - Thirty_Days:
                File_Name = FileName(files)
                Process_ID = ProcessID(files)
                for key in dictionary:
                    if key in Process_ID:
                        Email_Recipient = dictionary[key]
                        send_email(Email_Recipient, Subject_Line, Thirty_Days_Message)
            elif Current_Date >= (Creation_Date + One_Year) - Sixty_Days:
                File_Name = FileName(files)
                Process_ID = ProcessID(files)
                for key in dictionary:
                    if key in Process_ID:
                        Email_Recipient = dictionary[key]
                        send_email(Email_Recipient, Subject_Line, Sixty_Days_Message)
            elif Modification_Date >= Creation_Date + One_Year:
                File_Name = FileName(files)
                Process_ID = ProcessID(files)
                for key in dictionary:
                    if key in Process_ID:
                        Email_Recipient = dictionary[key]
                        send_email(Email_Recipient, Subject_Line, One_Year_Message)
                # if os.path.exists(dest_file_path):
                #     os.remove(dest_file_path)
                #     shutil.move(src_file_path, Destination)
                # else:
                #     shutil.move(src_file_path, Destination)

        elif Retention_Code == 'R3':
            if Current_Date >= (Creation_Date + Three_Years) - Thirty_Days:
                File_Name = FileName(files)
                Process_ID = ProcessID(files)
                for key in dictionary:
                    if key in Process_ID:
                        Email_Recipient = dictionary[key]
                        send_email(Email_Recipient, Subject_Line, Thirty_Days_Message)
            elif Current_Date >= (Creation_Date + Three_Years) - Sixty_Days:
                File_Name = FileName(files)
                Process_ID = ProcessID(files)
                for key in dictionary:
                    if key in Process_ID:
                        Email_Recipient = dictionary[key]
                        send_email(Email_Recipient, Subject_Line, Sixty_Days_Message)
            elif Modification_Date >= Creation_Date + Three_Years:
                File_Name = FileName(files)
                Process_ID = ProcessID(files)
                for key in dictionary:
                    if key in Process_ID:
                        Email_Recipient = dictionary[key]
                        send_email(Email_Recipient, Subject_Line, One_Year_Message)
                # if os.path.exists(dest_file_path):
                #     os.remove(dest_file_path)
                #     shutil.move(src_file_path, Destination)
                # else:
                #     shutil.move(src_file_path, Destination)
        elif Retention_Code == 'R5':
            if Current_Date >= (Creation_Date + Five_Years) - Thirty_Days:
                File_Name = FileName(files)
                Process_ID = ProcessID(files)
                for key in dictionary:
                    if key in Process_ID:
                        Email_Recipient = dictionary[key]
                        send_email(Email_Recipient, Subject_Line, Thirty_Days_Message)
            elif Current_Date >= (Creation_Date + Five_Years) - Sixty_Days:
                File_Name = FileName(files)
                Process_ID = ProcessID(files)
                for key in dictionary:
                    if key in Process_ID:
                        Email_Recipient = dictionary[key]
                        send_email(Email_Recipient, Subject_Line, Sixty_Days_Message)
            elif Modification_Date >= Creation_Date + Five_Years:
                File_Name = FileName(files)
                Process_ID = ProcessID(files)
                for key in dictionary:
                    if key in Process_ID:
                        Email_Recipient = dictionary[key]
                        send_email(Email_Recipient, Subject_Line, One_Year_Message)
                # if os.path.exists(dest_file_path):
                #     os.remove(dest_file_path)
                #     shutil.move(src_file_path, Destination)
                # else:
                #     shutil.move(src_file_path, Destination)

        elif Retention_Code == 'R10':
            if Current_Date >= (Creation_Date + Ten_Years) - Thirty_Days:
                File_Name = FileName(files)
                Process_ID = ProcessID(files)
                for key in dictionary:
                    if key in Process_ID:
                        Email_Recipient = dictionary[key]
                        send_email(Email_Recipient, Subject_Line, Thirty_Days_Message)
            elif Current_Date >= (Creation_Date + Ten_Years) - Sixty_Days:
                File_Name = FileName(files)
                Process_ID = ProcessID(files)
                for key in dictionary:
                    if key in Process_ID:
                        Email_Recipient = dictionary[key]
                        send_email(Email_Recipient, Subject_Line, Sixty_Days_Message)
            elif Modification_Date >= Creation_Date + Ten_Years:
                File_Name = FileName(files)
                Process_ID = ProcessID(files)
                for key in dictionary:
                    if key in Process_ID:
                        Email_Recipient = dictionary[key]
                        send_email(Email_Recipient, Subject_Line, One_Year_Message)
                # if os.path.exists(dest_file_path):
                #     os.remove(dest_file_path)
                #     shutil.move(src_file_path, Destination)
                # else:
                #     shutil.move(src_file_path, Destination)

        elif Retention_Code == 'R15':
            if Current_Date >= (Creation_Date + Fifteen_Years) - Thirty_Days:
                File_Name = FileName(files)
                Process_ID = ProcessID(files)
                for key in dictionary:
                    if key in Process_ID:
                        Email_Recipient = dictionary[key]
                        send_email(Email_Recipient, Subject_Line, Thirty_Days_Message)
            elif Current_Date >= (Creation_Date + Fifteen_Years) - Sixty_Days:
                File_Name = FileName(files)
                Process_ID = ProcessID(files)
                for key in dictionary:
                    if key in Process_ID:
                        Email_Recipient = dictionary[key]
                        send_email(Email_Recipient, Subject_Line, Sixty_Days_Message)
            elif Modification_Date >= Creation_Date + Fifteen_Years:
                File_Name = FileName(files)
                Process_ID = ProcessID(files)
                for key in dictionary:
                    if key in Process_ID:
                        Email_Recipient = dictionary[key]
                        send_email(Email_Recipient, Subject_Line, One_Year_Message)
                # if os.path.exists(dest_file_path):
                #     os.remove(dest_file_path)
                #     shutil.move(src_file_path, Destination)
                # else:
                #     shutil.move(src_file_path, Destination)
        elif Retention_Code == 'R20':
            if Current_Date >= (Creation_Date + Twenty_Years) - Thirty_Days:
                File_Name = FileName(files)
                Process_ID = ProcessID(files)
                for key in dictionary:
                    if key in Process_ID:
                        Email_Recipient = dictionary[key]
                        send_email(Email_Recipient, Subject_Line, Thirty_Days_Message)
            elif Current_Date >= (Creation_Date + Twenty_Years) - Sixty_Days:
                File_Name = FileName(files)
                Process_ID = ProcessID(files)
                for key in dictionary:
                    if key in Process_ID:
                        Email_Recipient = dictionary[key]
                        send_email(Email_Recipient, Subject_Line, Sixty_Days_Message)
            elif Modification_Date >= Creation_Date + Twenty_Years:
                File_Name = FileName(files)
                Process_ID = ProcessID(files)
                for key in dictionary:
                    if key in Process_ID:
                        Email_Recipient = dictionary[key]
                        send_email(Email_Recipient, Subject_Line, One_Year_Message)
                # if os.path.exists(dest_file_path):
                #     os.remove(dest_file_path)
                #     shutil.move(src_file_path, Destination)
                # else:
                #     shutil.move(src_file_path, Destination)
