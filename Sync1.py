import os
import win32com.client as client
from PIL import ImageGrab
excel=client.Dispatch('Excel.Application')
wb=excel.Workbooks.Open(r"C:\Users\UAKHPAL\OneDrive - BUNGE\Desktop\DBMS mapping darta\Daily Rollout Working files\Sync Report\Distributor Sync Status.xlsb")
sheet=wb.Sheets.Item(1)
copyrange=sheet.Range('A2:j23')
copyrange.CopyPicture(Appearance=1,Format=2)
ImageGrab.grabclipboard().save('RBM.jpg')
copyrange=sheet.Range('A28:H75')
copyrange.CopyPicture(Appearance=1,Format=2)
ImageGrab.grabclipboard().save('ASM.jpg')
sheet2=wb.Sheets.Item(2)
copyrange=sheet2.Range('A4:H15')
copyrange.CopyPicture(Appearance=1,Format=2)
ImageGrab.grabclipboard().save('New.jpg')
sheet3=wb.Sheets.Item(3)
copyrange=sheet3.Range('A4:H17')
copyrange.CopyPicture(Appearance=1,Format=2)
ImageGrab.grabclipboard().save('Old.jpg')
excel.quit()
RBM=r"C:\Users\UAKHPAL\OneDrive - BUNGE\Desktop\DBMS mapping darta\Daily Rollout Working files\Sync Report\RBM.jpg"
New=r"C:\Users\UAKHPAL\OneDrive - BUNGE\Desktop\DBMS mapping darta\Daily Rollout Working files\Sync Report\New.jpg"
Old=r"C:\Users\UAKHPAL\OneDrive - BUNGE\Desktop\DBMS mapping darta\Daily Rollout Working files\Sync Report\Old.jpg"
ASM=r"C:\Users\UAKHPAL\OneDrive - BUNGE\Desktop\DBMS mapping darta\Daily Rollout Working files\Sync Report\ASM.jpg"
html_body=r""" Hi All
<p>Please find the status of DMS Sync report.</p>
<H3><u>RBM</u></H3>
{Image1}
<H3><u>New Tally Users</u></H3>
{Image2}
<H3><u>Old System Users</u></H3>
{Image3}
<H3><u>SM</u></H3>
{Image4}
"""

outlook=client.Dispatch('Outlook.Application')
mail=outlook.CreateItem(0)
mail.HTMLBody=html_body
inspector=mail.GetInspector
inspector.Display()
mail.To ='arun.neogi@bunge.com;Sandeep.Kaul@bunge.com;subrata.das@bunge.com;mahesh.kumar@bunge.com;rs.murthy@bunge.com;sanjeev.kumar@bunge.com;allen.rodrigues@bunge.com;Rajesh.Mehta@bunge.com;raj.nandi@bunge.com;'
mail.Cc= 'digpal.singh@bunge.com;Sanjeev.Giri@bunge.com;Nisha.Chawla@bunge.com;Shamsundar.Sg@bunge.com;ankur.srivastava@bunge.com;Akanksha.Garg@bunge.com;Nand.Kishor@bunge.com;amit.mehra@bunge.com;akhilesh.pal@bunge.com'
mail.Bcc='Parveen.tarika@bunge.com;vivek.bansal@bunge.com;'
mail.Subject='Distributor Sync Status -DMS  till 17 JAN 2021'
doc=inspector.WordEditor
selection=doc.Content
selection.Find.Text=r"{Image1}"
selection.Find.Execute()
selection.Text=""
selection.Text
img=selection.InlineShapes.AddPicture(RBM,0,1)
selection.Find.Text=r"{Image2}"
selection.Find.Execute()
selection.Text=""
selection.Text
img=selection.InlineShapes.AddPicture(New,0,1)
selection.Find.Text=r"{Image3}"
selection.Find.Execute()
selection.Text=""
selection.Text
img=selection.InlineShapes.AddPicture(Old,0,1)
selection.Find.Text=r"{Image4}"
selection.Find.Execute()
selection.Text=""
selection.Text
img=selection.InlineShapes.AddPicture(ASM,0,1)
mail.Attachments.Add("C:/Users/UAKHPAL/OneDrive - BUNGE/Desktop/DBMS mapping darta/Daily Rollout Working files/Sync Report/Distributor Sync Status.xlsb")
