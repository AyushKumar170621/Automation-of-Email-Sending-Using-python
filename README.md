# Automation-of-Email-Sending-Using-python
Automation of  Email sending using python API.
Step 1: Setting up a Gmail account for Development: -
First, we must Create a Gmail account.
Then we must adjust Gmail account’s security settings to allow access from Python code.
Turn Allow less secure app on.

Step 2: Import the following modules
import smtplib
import time
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
import pandas as pd
import os
import datetime as dt

Step 3: Let’s set up a connection to our email server.
We will provide the server address and port no. to initiate SMTP connection.
now we will use smtp.ehlo to send an extended hello command.
And then, we will use smtp.starttls to enable the  transport layer security (TLS) encryption.

smtp=smtplib.SMTP('smtp.gmail.com',587)
smtp.ehlo()
smtp.starttls()

Step 4: Login to my gmail Account.
Take input from user their email and password and then login.
useremail=input("Enter your Email :")
userpassword=input("Enter password :")
 
smtp.login(useremail, userpassword)

Step 5: Inputting email address of receivers.
Ask user whether user want to enter address of image manually or want to give excel sheet.
If user give excel file, then read it using pandas.
ch=int(input("Enter 1 to give email address manually or  2 to give address of excel sheet :"))
if ch==1:
    n=int(input(("Enter number of TO recipients:")))
    for i in range(n):
        reci=input("Enter mail id:")
        to.append(reci)
    n=int(input(("Enter number of CC recipients :")))
    for i in range(n):
        reci=input("Enter mail id:")
        cc.append(reci)
    n=int(input(("Enter number of BCC recipients :")))
    for i in range(n):
        reci=input("Enter mail id:")
        bcc.append(reci)
else :
    ads=input("Enter address of Excel sheet :") 
    # Reading using pandas
    e=pd.read_excel(ads)
    to=e['TO'].values
    cc=e['CC'].values
    bcc=e['BCC'].values

Step 6: Now we will take input of subject for email from user.
sub=input("Enter subject of email : ")

Step 7: Now we will take input of text message for email from user.
txt=input("Enter message :")

Step 8: In this step we will Import images to attach.
Ask user to give address of folder or enter manually.
If address of folder is given by user then list all directories in folder and make list of all file ending with jpg, jpeg or png.
ch=int(input("Enter 1 to give address of image folder or 2 to give address of image manually one by one:"))
if ch==1:
    fol_path = input("Enter folder path :")
    fol_path=fol_path+'/'
    files = os.listdir(fol_path)
    # making list of images
    for file in files:
        # make sure file is an image
        if file.endswith(('.jpg', '.png', 'jpeg')):
            picture_path = fol_path + file
            pics.append(picture_path)
else:
    ran=int(input("Enter number of images :"))
    for i in range(ran):
        x=input("Enter address :")
        pics.append(x)

Step 9: lets Import other attachment.
Ask user to give total number of attachment and ,then input address of attachment one by one.
attach=[]
ath_no=int(input("Enter number of attachment :"))
for i in range(ath_no):
    at_path=input("Enter address :")
    attach.append(at_path)

Step 10: Now, we built the message content.
First, we assign the MIMEMultipart object to the msg variable after initializing it.
Call msg=message(sub,to,cc,bcc,txt,pics,attach)
def message(subject='Python  Email',to=None,cc=None,bcc=None,text="",img=None,attachment=None):
    msg=MIMEMultipart()
    msg['To']=','.join(map(str, to))
    msg['Cc']=','.join(map(str, cc))
    msg['Bcc']=','.join(map(str, bcc))
    msg['Subject']=subject
    msg.attach(MIMEText(text))

Step 11: Attach pictures to the msg in message fuction.
First, we read the image as binary data.
Now, attach the image data to MIMEMultipart using MIMEImage, we add the given filename use os.basename
	for each_img in img:
            imgdata = open(each_img, 'rb').read()  
            msg.attach(MIMEImage(imgdata,name=os.path.basename(each_img)))
Step 12: Attach other attachment to the msg in message function.
Read in the attachment using MIMEApplication.
Finally, we add the attachment to our message.
for every_attachment in attachment:
            with open(every_attachment, 'rb') as f:
                file = MIMEApplication(
                    f.read(),name=os.path.basename(every_attachment) )
            file['Content-Disposition'] = f'attachment;\
            filename="{os.path.basename(every_attachment)}"'
            msg.attach(file)

now return from message function
return msg

Step 13: Ask whether user want to message Instantly or want to Schedule the message.
send_ch=int(input("Enter 1 to Sent Email Instantly or 2 to Schedule :"))
if send_ch==1:
    print("Sending....")
else:
    year,month,day=[int(x) for x in input("Enter year month and date :").split()]
    hour,minute,sec=[int(x) for x in input("Enter hour minute and second :").split()]
    schedule_time = dt.datetime(year,month,day,hour,minute,sec) 
    time.sleep(schedule_time.timestamp() - time.time())
    print("Sending....")

Step 14: Send message and quit the server.
First, we make a list of all the emails we want to send.
toadd=[]
toadd.extend(to)                              
toadd.extend(cc)                              
toadd.extend(bcc)

Then, by using the sendmail function, pass parameters such as from address, to address, and the message content. Now, send the message and quit server.
smtp.sendmail(from_addr=useremail,to_addrs=toadd, msg=msg.as_string())
smtp.quit()
print("Email Sent")
