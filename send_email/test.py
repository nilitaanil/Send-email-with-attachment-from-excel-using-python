import pandas as pd
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

your_name = "nilita anil"
your_email = "nilita.anil@gmail.com"
your_password = "**************"

server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
server.ehlo()
server.login(your_email, your_password)

email_list = pd.read_excel("/home/pachi/test_sheet.ods", engine="odf")

all_names = email_list['Name']
all_emails = email_list['Email']
all_subjects = email_list['Subject']
all_messages = email_list['Message']
all_image = email_list['Image']


# To check is file exits and is accessible in the location given
def is_accessible(path, mode='r'):
    try:
        f = open(path, mode)
        f.close()
    except IOError:
        return False
    return True

#
# for image in all_image:
#     file_location = "/home/pachi/" + image
#     if not is_accessible(file_location):
#         print(image, "not found in the given directory.")



# Loop through the emails
for idx in range(len(all_emails)):

    # Get each records name, email, subject and message
    name = all_names[idx]
    email = all_emails[idx]
    subject = all_subjects[idx]
    msg = all_messages[idx]
    image = all_image[idx]

    message = MIMEMultipart()
    message["From"] = "(NITK GradBook)<your_email>"
    message["To"] = email
    message["Subject"] = subject
    message.attach(MIMEText(msg, "plain"))

    path_folder = "/home/pachi/"  # add ypur path
    filename = "photo.pdf"
    attach_photo = path_folder + image  # instead of sample.pdf add the excel image column
    attachment = open(attach_photo, "rb")

    # instance of MIMEBase and named as p
    p = MIMEBase('application', 'octet-stream')

    # To change the payload into encoded form
    p.set_payload(attachment.read())

    # encode into base64
    encoders.encode_base64(p)
    p.add_header('Content-Disposition', "attachment; filename= %s" % filename)

    # attach the instance 'p' to instance 'msg'

    message.attach(p)
    text = message.as_string()

    try:
        # server.sendmail(your_email, [email], full_email)
        server.sendmail(your_email, [email], text)
        print('Email to {} successfully sent!\n\n'.format(email))
    except Exception as e:
        print('Email to {} could not be sent :( because {}\n\n'.format(email, str(e)))

# Close the smtp server
server.close()
