import os
import time
from datetime import date
import win32com.client as client
today=date.today()
today1=str(today)
import datetime
x=datetime.datetime.now()
from PIL import ImageGrab
excel=client.Dispatch('Excel.Application')
wb=excel.Workbooks.Open(r"C:\Users\UAKHPAL\OneDrive - BUNGE\Desktop\Party\B2C Model parties progress report.xlsx")
sheet=wb.Sheets.Item(4)
import pandas as pd
party=pd.read_excel('party List.xlsx')
for ind in party.index:
    Su=party['Subject'][ind]
    Na=party['Name'][ind]
    Em=party['Email'][ind]
    MA=party['RBM'][ind]
    sheet.Cells(3,2).Value = Su
    copyrange=sheet.Range('F2:R30')
    copyrange.CopyPicture(Appearance=1,Format=2)
    ImageGrab.grabclipboard().save('RBM.jpg')
    RBM=r"C:\Users\UAKHPAL\OneDrive - BUNGE\Desktop\Party\RBM.jpg"
    html_body=r"""Dear """ +Na+""",
    <p>Please find below your automation score card.</p>
    {Image1}
    <br></br>
    <br></br>
    <br></br>
    <br></br>
    <H4> Thanks & Regards,</H4>
    <H5> Akhilesh Pal </H5>
    """
    outlook=client.Dispatch('Outlook.Application')
    mail=outlook.CreateItem(0)
    mail.HTMLBody=html_body
    inspector=mail.GetInspector
    inspector.Display()
    mail.To =Em
    mail.Cc=MA
    mail.Subject=Na+" | DMS Performance | "+today1
    doc=inspector.WordEditor
    selection=doc.Content
    selection.Find.Text=r"{Image1}"
    selection.Find.Execute()
    selection.Text=""
    selection.Text
    img=selection.InlineShapes.AddPicture(RBM,0,1)
    per = 25
    img.Height = int(per*20.00)
    img.Width  = int(per*30.400)
    #mail.Send()
    time.sleep(5)
excel.Quit()