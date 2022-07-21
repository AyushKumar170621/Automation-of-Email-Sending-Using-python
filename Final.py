import smtplib
import time
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
import pandas as pd
import os
import datetime as dt

#set up a connection to our email server
smtp=smtplib.SMTP('smtp.gmail.com',587)
smtp.ehlo()
smtp.starttls()


#login to gmail Account
useremail=input("Enter your Email :")
userpassword=input("Enter password :")

smtp.login(useremail, userpassword)

def message(subject='Python  Email',to=None,cc=None,bcc=None,text="",img=None,attachment=None):
    msg=MIMEMultipart()
    msg['To']=','.join(map(str, to))
    msg['Cc']=','.join(map(str, cc))
    msg['Bcc']=','.join(map(str, bcc))
    msg['Subject']=subject
    msg.attach(MIMEText(text))
    #check if img is none
    if img is not None:
        # Check whether img is list oo not
        if type(img) is not list:  
            img = [img] 
        # Now iterate through our list
        for each_img in img:
            #read image in binary form
            imgdata = open(each_img, 'rb').read()  
            # Attach the image data to MIMEMultipart using MIMEImage & os.path.basename
            msg.attach(MIMEImage(imgdata,name=os.path.basename(each_img)))
  
    if attachment is not None:
        #check whether the we have list of attachment
        if type(attachment) is not list:
             attachment = [attachment]  
        #iterating
        for every_attachment in attachment:
            # Read the attachment using MIMEApplication
            with open(every_attachment, 'rb') as f:
                
                file = MIMEApplication(
                    f.read(),name=os.path.basename(every_attachment)
                )
            file['Content-Disposition'] = f'attachment;\
            filename="{os.path.basename(every_attachment)}"'
            #attach to message
            msg.attach(file)
    return msg

#inputing Email address of Receiver
to=[] 
toadd=[]
bcc=[]
cc=[]
ch=int(input("Enter 1 to give email address manually or  2 to give address of execel sheet :"))
if ch==1:
    n=int(input(("Enter number of recipents ('TO'):")))
    for i in range(n):
        r=input("Enter mail id:")
        to.append(r)
    n=int(input(("Enter number of recipent ('CC'):")))
    for i in range(n):
        r=input("Enter mail id:")
        cc.append(r)
    n=int(input(("Enter number of recipent ('BCC'):")))
    for i in range(n):
        r=input("Enter mail id:")
        bcc.append(r)
else :
    ads=input("Enter address of Excel sheet :")
    # reading using pandas
    e=pd.read_excel(ads)
    to=e['TO'].values
    cc=e['CC'].values
    bcc=e['BCC'].values

# attaching to cc and bcc to toadd
toadd.extend(to)                              
toadd.extend(cc)                              
toadd.extend(bcc)

#input Subject of Email
sub=input("Enter subject of email :")

#input Message
txt=input("Enter message :")

#imporing image attachment
pics=[]
ch=int(input("Enter 1 to give address of image folder or 2 to give address of image manually one by one:"))
if ch==1:
    folder_path = input("Enter folder path :")
    folder_path=folder_path+'/'
    files = os.listdir(folder_path)
    # making list of images
    for file in files:
        # make sure file is an image
        if file.endswith(('.jpg', '.png', '.jfif')):
            pic_path = folder_path + file
            pics.append(pic_path)
else:
    ran=int(input("Enter number of images :"))
    for i in range(ran):
        x=input("Enter address :")
        pics.append(x)

#importing other attachment
attach=[]
ath_no=int(input("Enter number of attachment :"))
for i in range(ath_no):
    at_path=input("Enter address :")
    attach.append(at_path)

#building the message content to send
msg=message(sub,to,cc,bcc,txt,pics,attach)

#asking to schedule or send instantly
send_ch=int(input("Enter 1 to Send Email Instantly or Enter  2 to Schedule :"))
if send_ch==1:
    print("Sending....")
else:
    year,month,day=[int(x) for x in input("Enter year month and date of schedule time :").split()]
    hour,minu,sec=[int(x) for x in input("Enter hour minute and second of schedule time :").split()]
    schedule_time = dt.datetime(year,month,day,hour,minu,sec) 
    time.sleep(schedule_time.timestamp() - time.time())
    print("Sending....")

#sending Message
smtp.sendmail(from_addr=useremail,to_addrs=toadd, msg=msg.as_string())
smtp.quit()
print("Email Sent Sucessfully")