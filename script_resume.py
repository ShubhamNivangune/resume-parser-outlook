import win32com.client
import os
import shutil
from tkinter import * 
from tkinter import messagebox
from datetime import datetime, timedelta
#----------------------------------Import lib-----------------------------------

outlook = win32com.client.Dispatch('outlook.application')
mapi = outlook.GetNamespace("MAPI")
for account in mapi.Accounts:
	print(account.DeliveryStore.DisplayName)
inbox = mapi.GetDefaultFolder(6)
messages = inbox.Items
#Let's assume we want to save the email attachment to the below directory
outputDir = r"G:\portfolio\attachement"
try:
    for message in list(messages):
        try:
            s = message.sender
            for attachment in message.Attachments:
                attachment.SaveASFile(os.path.join(outputDir, attachment.FileName))
                print(f"attachment {attachment.FileName} from {s} saved")
                
        except Exception as e:
            print("error when saving the attachment:" + str(e))
except Exception as e:
		print("error when processing emails messages:" + str(e))
  
#-----------------------Doc ,  pdf sorting--------------------------------

arr = os.listdir('G:\portfolio\Attachement')
for i in range(len(arr)):
    if(('.docx') in arr[i] or ('.doc') in arr[i] or ('.pdf') in arr[i]  ):
        original = r'G:\\portfolio\\Attachement\\'+arr[i]
        target = r'G:\portfolio\Sorted\\'+arr[i]
        print(shutil.copyfile(original, target))
    else:
        print("Not Found")
       
#------------------------- specific resume sorting (With condition)---------------------------------------

arr = os.listdir('G:\portfolio\Sorted')
data=[]
for i in range(len(arr)):
    word = win32com.client.Dispatch("Word.Application")
    word.visible = False
    wb = word.Documents.Open("G:\\portfolio\\Sorted\\"+arr[i])
    doc = word.ActiveDocument
    data=doc.Range().Text
    if ('summary' or 'skill' or 'resume' or 'education' or 'experience' in data):
        print(arr[i])
        original = r'G:\\portfolio\\Sorted\\'+arr[i]
        target = r'G:\portfolio\Resumes\\'+arr[i]
        print(shutil.copyfile(original, target))    
word.Application.Quit()

#---------------------------@solicitous business solution pvt ltd--------------------------------------------------