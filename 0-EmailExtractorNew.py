import os
os.getcwd()
os.chdir(r"H:\@TMP")
os.chdir(r"H:\pandas-news-mining-master")



###################################################################################################

import datefinder
import win32com.client
import codecs
import os
from bs4 import BeautifulSoup

#import itertools

'''###################################################################################################
   Export Outlook emails to TXT files
   Emails should be under Outlook folder "@ News" >> "To Export"
   This part can only run on office computer with outlook configration
   It needs only run once for a set of emails (say for emails in 2015)
###################################################################################################'''

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
folder = outlook.Folders("Dayong.Yu@wto.org").Folders("@ News").Folders("@ News 2016").Folders("tmp")

messages = folder.Items
print(messages.count)

for message in messages:
    print(message.subject)
    str_subject = message.subject.replace('News, Monday', '')
    str_subject = str_subject.replace('news, Monday', '')
    str_subject = str_subject.replace('news. Monday', '')
    str_subject = str_subject.replace('news Monday', '')
    
    # Extract date from subject to use as file name
    matches = list(datefinder.find_dates(str_subject))
    if len(matches) > 0:    
        filedate = matches[0].date()
    else:
        filedate = message.subject
    # Save file as ASCII, but have to ignore errors!!
    text_file = codecs.open('Files2016ASCII/' + str(filedate) +'.txt', 'w', encoding="ascii", errors="ignore")
#    text_file = open('Files2016/' + str(filedate) +'.txt', 'w')    
    soup = BeautifulSoup(message.body, "lxml")
    text_file.write(soup.get_text())
    text_file.close()




