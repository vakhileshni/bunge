import time
import pandas as pd
import os
import win32com.client as win32
ak=pd.read_excel(r"C:\Users\UAKHPAL\OneDrive - BUNGE\Desktop\DR\Plant_Name.xlsx")
ap='C:/Users/UAKHPAL/OneDrive - BUNGE/Desktop/DR/BR/DR_PUR/'
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
    mail.Cc= 'billinghelpdesk@bunge.com'
    mail.Subject = "Closing Report For Month of Dec '2020'  "+a+' '+jh
    mail.Body = 'Dear '+name+','+'\n'+'\n'+'Please find attached closing reports for the month of December 2020.'+'\n'+'\n'+'\n'+'\n'+'\n'+'\n'+"Thanks & Regards"+'\n'+'Akhilesh Pal'+'\n'
    arr = os.listdir(ap+a)
    for i in arr:
        mail.Attachments.Add(ap+a+"/"+i)
    mail.Send()
    time.sleep(5)
