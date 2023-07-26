import win32com.client
#other libraries to be used in this script
import os
from datetime import datetime, timedelta


if __name__ == "__main__":
    outlook = win32com.client.Dispatch('outlook.application').GetNamespace("MAPI")
    
    selectedAccount = ""
    for account in outlook.Accounts:	    
        if "gmail" in account.DeliveryStore.DisplayName:
            selectedAccount = account
            break
    
    #print(selectedAccount)
    inbox = outlook.Folders('email@email.com').Folders('Inbox')
    #for f in outlook.Folders:
    #    print(f)
    #inbox = outlook.GetDefaultFolder(6) #6- Inbox
    messages = inbox.Items
    
    received_dt = datetime.now() - timedelta(days=500)
    received_dt = received_dt.strftime('%m/%d/%Y %H:%M %p')
    messages = messages.Restrict("[ReceivedTime] >= '" + received_dt + "'")
    
    messages = messages.Restrict("[SenderEmailAddress] = 'noreply@kaggle.com'")
    #messages = messages.Restrict("[Subject] = 'Sample Report'")
    
    for message in messages:
        print(message.Subject)
        #message.delete()