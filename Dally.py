import pandas as pd
import numpy as np
import os
import win32com.client as client
import dataframe_image as dfi
df=pd.read_excel(r"C:\Users\UAKHPAL\OneDrive - BUNGE\Desktop\DBMS mapping darta\Daily Rollout Working files\Rollout_Status.xlsx",sheet_name='count')
df1=pd.read_excel(r"C:\Users\UAKHPAL\OneDrive - BUNGE\Desktop\DBMS mapping darta\Daily Rollout Working files\Rollout_Status.xlsx",sheet_name='Table')
dfi.export(df,"count.jpg")
dfi.export(df1,"Table.jpg")
count=r"C:\Users\UAKHPAL\OneDrive - BUNGE\Desktop\DBMS mapping darta\Daily Rollout Working files\count.jpg"
Table=r"C:\Users\UAKHPAL\OneDrive - BUNGE\Desktop\DBMS mapping darta\Daily Rollout Working files\Table.jpg"
html_body=r""" Hi Garima
<p>Sharing the list of missing items in DMS Mapped Item List relate to those distributors which were marked rollout “Done”
<br></br>
<br></br>
Please review following tables and tell us the reason for missing.
</p>
{Image1}
<br></br>
<br></br>
{Image2}
"""
outlook=client.Dispatch('Outlook.Application')
mail=outlook.CreateItem(0)
mail.HTMLBody=html_body
inspector=mail.GetInspector
inspector.Display()
doc=inspector.WordEditor
selection=doc.Content
selection.Find.Text=r"{Image1}"
selection.Find.Execute()
selection.Text=""
selection.Text
img=selection.InlineShapes.AddPicture(count,0,1)
selection.Find.Text=r"{Image2}"
selection.Find.Execute()
selection.Text=""
selection.Text
img=selection.InlineShapes.AddPicture(Table,0,1)
shadow=img.Shadow
