import win32com.client as win
import os
import re
from datetime import datetime

#  connect to Outlook
outlook = win.Dispatch("Outlook.Application")
mapi = outlook.GetNamespace("MAPI")
now = datetime.now()
month = str(int(now.strftime("%m"))-1).zfill(2)

def autoRetrive(month):
    parentDir = r"\\172.16.10."
    path = os.path.join(parentDir, month)
    
    if not os.path.exists(path):
        os.mkdir(path)
      
    # outlook test測試資料夾路徑如下: messages = mapi.Folders(1).Folders(2).Folders(12).Items
    # 找到KRI的index位於5
    messages = mapi.Folders(1).Folders(2).Folders(1).Folders(3).Folders(5).Items
    for msg in list(messages):
        if msg.Unread and '主要風險指標(KRI)彙總通報表' in msg.Subject:
            for attachment in msg.Attachments:
                reg = re.match('.*xlsx',attachment.FileName)
                msg.Unread = False
                if reg :
                    attachment.SaveAsFile(os.path.join(parentDir, month, str(attachment.FileName)))
                    # if '主要風險指標(KRI)彙總通報表' in msg.Subject:
                    #     break

autoRetrive(month)











