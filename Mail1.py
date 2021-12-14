import time
import pandas as pd
import os
import datetime
today = datetime.date.today()
first = today.replace(day=1)
lastMonth = first - datetime.timedelta(days=1)
month=lastMonth.strftime("%B")
x = datetime.datetime.now()
year=x.strftime("%Y")
import win32com.client as win32
parent_dir=os.getcwd()
ak=pd.read_excel(parent_dir+"\\Plant_Name.xlsx")
ap=parent_dir+'\\BR\\DR_PUR\\'
for ind in ak.index:
    a=ak['PLANT NO'][ind]
    jh=ak['Plant Name'][ind]
    name=ak['Name'][ind]
    a1=ak['Email'][ind]
    a=str(a)
    b=ap+a
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.Display()
    mail.To =a1
    mail.Subject = "Closing Report For Month of "+month+" "+year+" "+a+" "+jh
    mail.SentOnBehalfOfName='billinghelpdesk@bunge.com'
    mail.Body = 'Dear '+name+','+'\n'+'\n'+'Please find attached closing reports for the month of '+month+" "+year+"."+'\n'+'\n'+'\n'+'\n'+'\n'+'\n'+"Thanks & Regards"+'\n'+'Amit Mehra'+'\n'
    arr = os.listdir(ap+a)
    for i in arr:
        mail.Attachments.Add(ap+a+"/"+i)
    mail.Send()
    time.sleep(10)
