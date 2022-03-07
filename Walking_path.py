# -*- coding: utf-8 -*-
"""
Created on Sun Feb  6 22:11:33 2022

@author: TiaaUser
"""
import webbrowser
import re
import os
from datetime import datetime
import time
import openpyxl
#import schedule
import time
import smtplib
import ssl
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import requests


print("....Starting.....")


def folder_create(currentYear, currentMonth, currentDay, data):
    path = "C:/Users/TiaaUser/Desktop/P_folder/"
    if not os.path.isdir(path):
        os.mkdir(path)
    
    # creaing folder by Year  logic here
    if not os.path.isdir(path+'/'+currentYear):
        os.mkdir(path+'/'+currentYear)
        data.append("Year folder is created")
   
    # creaing folder by Month  logic here
    if not os.path.isdir(path+'/'+currentYear+'/'+currentMonth):
        os.mkdir(path+'/'+currentYear+'/'+currentMonth)
        os.mkdir(path+'/'+currentYear+'/'+"Log")
        data.append("Month and log folder is created")
        
   # creaing folder by day     
    if not os.path.isdir(path+'/'+currentYear+'/'+currentMonth+'/'+currentDay):
        weekno = datetime.today().weekday()
        if weekno < 5 :
            os.mkdir(path+'/'+currentYear+'/'+currentMonth+'/'+ str(currentDay))
            data.append(f"Day folder is created : {currentDay}")
            C_Day= path+'/'+currentYear+'/'+currentMonth+'/'+currentDay
            time.sleep(2)
            create_excel(C_Day, data)
        else:  # 5 Sat, 6 Sun
            data.append("It's a weekend")
    data.append("Year, Month, Day folder's are already exists")            
    return data  

# creaing folder by Excel sheet by day  logic here
def create_excel(C_Day, data):
    now = datetime.now().strftime("%d")
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.title = now
    wb.save(C_Day+'/'+now+'.xlsx')
    data.append("Excel file is succefully created")

def is_connected(data):
    #return True
    url = "https://www.google.com"
    try:
        request = requests.get(url, timeout = 1)
        data.append("Connected to the Internet")
        return True
    except (requests.ConnectionError, requests.Timeout) as err:
        data.append("No internet connection.")
        return False

def Find(string):
    url= re.findall('http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\(\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+', string)
    return url
    
def WebLauncher(path, data):
    with open (path) as fp:
        for line in fp:
            print(line)
            url = Find(line)
            for str in url:
                webbrowser.open(str, new = 2)
                data.append(f"Tab opened {url}")
    return data 

# sending email logic here

def send_mail(log_path):
    subject = "Daily Activity report"                        #str(input("Enter the subject of email :"))
    body = "Please find the attachment of daily Python script run details"                                    #str(input("Enter the body of email :"))
    sender_email = "ingle38877@gmail.com"
    password = "Ram@8793340769"
    receiver_email = "ingle3889@gmail.com"               #str(input("Enter Destination email id :"))
    

    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Subject"] = subject
    message["Bcc"] = receiver_email  # Recommended for mass emails

    message.attach(MIMEText(body, "plain"))
    
    
    filename = log_path
    with open(filename, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header(
        "Content-Disposition",
        f"attachment; filename= {filename}",)

    message.attach(part)
    text = message.as_string()
    
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(sender_email, password)
        server.sendmail(sender_email, receiver_email, text)
        print("Email succesfully sent")
          
# Main function calling  
def main(): 
    path = "C:/Users/TiaaUser/Desktop/Daily/29.txt"
    path_s = "C:/Users/TiaaUser/Desktop/P_folder/"
    data= []
    currentDay = str(datetime.now().strftime("%d"))
    currentYear = str(datetime.now().year)
    currentMonth = datetime.today().strftime("%b")
    #schedule.every().day.at("01:50").do(folder_create(currentYear, currentMonth, currentDay, data))
    folder_create(currentYear, currentMonth, currentDay, data)
    #schedule.every().day.at("01:51").do(WebLauncher(path, data))
    Connected = is_connected(data)
    if Connected:
        WebLauncher(path, data)
    else:
        data.append("Unable to connect to internet")    
    now = datetime.now().strftime("%Y_%b_%d_%H_%M_%S")
    
    with open(path_s+'/'+currentYear+'/'+"Log"+'/'+now+".txt",'w',encoding = 'utf-8') as f:
        f.write("Programmatacally created log file day wise :"+now+"\n")
        for i in range(len(data)):           
            f.write("\n"+str(i)+':---%s\n' % data[i])               
        f.write("\n"+"All data succefully saved and closed")     
    log_path =  path_s+'/'+currentYear+'/'+"Log"+'/'+now+".txt"
    #schedule.every().day.at("01:52").do(send_mail(log_path, password))
    
    send_mail(log_path)
    
if __name__ == "__main__":
    start_time = time.time()
    main()
    end_time = time.time()
    total = end_time-start_time
    print("Program Execution time is :",total)
    exit()
    
    