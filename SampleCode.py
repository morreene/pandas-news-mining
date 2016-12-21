
import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case,
                                    # the inbox. You can change that number to reference
                                    # any other folder
messages = inbox.Items
print(messages.count)
#message = messages.GetLast()
#body_content = message.body
#print(body_content)

for message in messages:
    print(message.subject)


# Print all email boxes
import win32com
outlook=win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

for i in range(50):
    try:
        box = outlook.Folders(i)
        name = box.Name
        print(i, name)
    except:
        pass


# Print all default folders
import win32com
outlook=win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
for i in range(50):
    try:
        box = outlook.GetDefaultFolder(i)
        name = box.Name
        print(i, name)
    except:
        pass



###################################################æ

import win32com.client
#import active_directory
session = win32com.client.gencache.EnsureDispatch("MAPI.session")
win32com.client.gencache.EnsureDispatch("Outlook.Application")
outlook = win32com.client.Dispatch("Outlook.Application")
mapi = outlook.GetNamespace('MAPI')
inbox =  mapi.GetDefaultFolder(win32com.client.constants.olFolderInbox)

fldr_iterator = inbox.Folders   
desired_folder = None
while 1:
    f = fldr_iterator.GetNext()
    if not f: break
    print(f.Name)    
    
    if f.Name == 'test':
        print('found "test" dir')
        desired_folder = f
        break

print(desired_folder.Name)





