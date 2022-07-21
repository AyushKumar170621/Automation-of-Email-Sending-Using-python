import smtplib
import pandas as pd
import os
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart

smtp=smtplib.SMTP('smtp.gmail.com',587)
smtp.ehlo()
smtp.starttls()

user_email=input("Enter your Email :")
user_password=input("Enter password :")

smtp.login(user_email, user_password)

def message(subject='Python  Email',text="",img=None,attachment=None):
    msg=MIMEMultipart()
    msg['Subject']=subject
    msg.attach(MIMEText(text))
    if img is not None:
        if type(img) is not list:  
            img = [img] 
        for one_img in img:
            img_data = open(one_img, 'rb').read()  
            msg.attach(MIMEImage(img_data,
                                 name=os.path.basename(one_img)))
  
    if attachment is not None:
        if type(attachment) is not list:
            
            attachment = [attachment]  
  
        for one_attachment in attachment:
  
            with open(one_attachment, 'rb') as f:
                
                file = MIMEApplication(
                    f.read(),
                    name=os.path.basename(one_attachment)
                )
            file['Content-Disposition'] = f'attachment;\
            filename="{os.path.basename(one_attachment)}"'
              
            msg.attach(file)
    return msg
'''import glob
images = glob.glob('images/*.jpg')'''
'''from pathlib import Path
images = Path("/var/www/html/myfolder/images").glob("*.jpg")'''
'''import os
path = "path/to/img/folder/"
jpgs = [os.path.join(path, file)
        for file in os.listdir(path)
        if file.endswith('.jpg')]'''
txt=input("Enter message :")
pics=["C:/Users/kvbhk/Desktop/PROJECT/4-43169_m.jpg",
"C:/Users/kvbhk/Desktop/PROJECT/35386.jpg",
"C:/Users/kvbhk/Desktop/PROJECT/HD-wallpaper-spiderman-homecoming-marvel-spider.jpg"]
msg=message('Automated Email',txt,pics,r"C:\PROGRAMING\DSA C\C revison\3rdnonrepeat.c")

#to=["ayushk1701@gmail.com","ayushpanwar691@gmail.com","kvbhkd000043@rediffmail.com"]
to=[]
ch=int(input("Enter 1 to give email address manually 2 for execel sheet "))
if ch==1:
    n=int(input(("Enter number of recipent")))
    for i in range(n):
        r=input("Enter mail id:")
        to.append(r)
else :
    ads=input("Enter address of Excel sheet :")
    e=pd.read_excel(ads)
    to=e['Email'].values
smtp.sendmail(from_addr=user_email,to_addrs=to, msg=msg.as_string())
smtp.quit()