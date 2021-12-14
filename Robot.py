import win32com.client
import datetime
import os
import email
 
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
folder = outlook.Folders.Item("BAS BillingHelpdesk")
inbox = folder.Folders.Item("Inbox")
#inbox = outlook.GetDefaultFolder(6).Folders.Item("billinghelpdesk@bunge.com")
message = inbox.items
for message in inbox.Items:
    if message.Unread == True:
        #for attachment in message.Attachments:
                        #print(attachment.FileName)
                        print(message.Subject)




import win32com.client
import datetime
import os
import email
from time import time
OneMinutes = time() + 30

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
folder = outlook.Folders.Item("BAS BillingHelpdesk")
inbox = folder.Folders.Item("Inbox").Folders['ROBOT']
message = inbox.items
for message in inbox.Items:
    if message.Unread == True:
        print(message.Subject)
        if time() > OneMinutes : break