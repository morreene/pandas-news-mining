import pandas as pd
import datefinder
import numpy as np
#import itertools
import win32com.client
import codecs
import os
import re

###################################################################################################
# Export Outlook emails to TXT files
# Emails should be under Outlook folder "@ News" >> "To Export"
###################################################################################################

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
folder = outlook.Folders("Dayong.Yu@wto.org").Folders("@ Other").Folders("News").Folders("@ News 2015").Folders("tmp")

messages = folder.Items
#print(messages.count)

for message in messages:
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

    text_file = codecs.open('Files2015A/' + str(filedate) +'.txt', 'w', encoding='utf-8')
    text_file.write(message.body)
    text_file.close()
#    print(message.subject)

###################################################################################################
# Extract articles from a TXT file and convert to DF
###################################################################################################



def extractor(filename): 
#    filename = 'Files2015A/2015-06-06.txt'
    file = codecs.open(filename, 'r', encoding='utf-8', errors='ignore')
    lines = [line.strip() for line in file if line.strip()]
    file.close()
    
    df_raw = pd.DataFrame(lines, columns=['Texts']).reset_index().rename(columns={'index':'ParaID'})
    
    # Clean texts
    # Remove bulletes
    df_raw['Texts'] = df_raw['Texts'].str.replace(r'•\t', '')
    df_raw['Texts'] = df_raw['Texts'].str.replace(r'•', '')
    df_raw['Texts'] = df_raw['Texts'].str.replace(r'·', '')
    df_raw['Texts'] = df_raw['Texts'].str.replace(r'·', '')
    df_raw['Texts'] = df_raw['Texts'].str.replace(r'¡P', '')
    df_raw['Texts'] = df_raw['Texts'].str.replace(r'-', '')
    df_raw['Texts'] = df_raw['Texts'].str.replace(r'·', '')
    df_raw['Texts'] = df_raw['Texts'].str.replace(r'·', '')    
    df_raw['Texts'] = df_raw['Texts'].str.replace(r' +', ' ')
#    df_raw['Texts'] = df_raw['Texts'].str.replace(r'\n', '')
#    df_raw['Texts'] = df_raw['Texts'].str.replace(r'\r', '')
    
    # Remove any email address
    for i in range(0, len(df_raw)):
        match = re.findall(r'[\w\.-]+@[\w\.-]+', df_raw.loc[i, 'Texts'])
        if len(match)>0: df_raw.loc[i, 'Texts'] = df_raw.loc[i, 'Texts'].replace(match[0], '')
        
    # Remove hyperlink
    df_raw['Texts'] = df_raw['Texts'].str.replace(r'HYPERLINK \"javascript\:void\(0\)\;\"', '')
    df_raw['Texts'] = df_raw['Texts'].str.replace(r'HYPERLINK \\l \"', '')
    df_raw['Texts'] = df_raw['Texts'].str.replace(r'\\l \"', '')
    df_raw['Texts'] = df_raw['Texts'].str.replace(r'\\l', '')

    df_raw['Texts'] = df_raw['Texts'].str.strip()
    
#    df_raw.loc[df_raw['Texts'].str.contains('END'), 'ToDelete'] = True
    df_raw.loc[df_raw['Texts'].str.contains('Information and External Relations Division'), 'ToDelete'] = True
    df_raw.loc[df_raw['Texts'].str.contains('Josep Bosch'), 'ToDelete'] = True
    df_raw.loc[df_raw['Texts'].str.contains('Head of Media Section'), 'ToDelete'] = True
    df_raw.loc[df_raw['Texts'].str.contains('41227395681'), 'ToDelete'] = True
    df_raw.loc[df_raw['Texts'].str.contains('Please consider the environment before printing this email or its attachment'), 'ToDelete'] = True
    df_raw.loc[df_raw['Texts'].str.contains(r'\d{1,4} words'), 'ToDelete'] = True
    
    df_raw = df_raw[df_raw['ToDelete'].isnull()]
    
    # Find records the same as titles in headline
    df_raw['TitleFlag'] = df_raw['Texts'].duplicated()
    
    # First duplication above are titles, sometimes dates or agency are in single line and may be duplicated
    first_title = df_raw[df_raw['TitleFlag'] == True].iloc[0]['ParaID']
    
    # Put titles and articles in different DF
    df_title = df_raw.iloc[0:first_title][['ParaID', 'Texts']].rename(columns={"Texts": "Title"})
    df_article = df_raw.iloc[first_title:][['ParaID', 'Texts']].rename(columns={"Texts": "Content"})
    
    # Matche titles with ful article
    df = pd.merge(df_article, df_title, how='left', left_on='Content', right_on='Title')
    df = df.fillna(method='ffill')
    df = df.rename(columns={"ParaID_y": "ArticleCode"})
    
    # Concatenate contents of article
    df = df[df['Title']!=df['Content']].groupby(['ArticleCode','Title'])['Content'].apply(lambda x: "%s" % ' '.join(x)).reset_index()
    df['FileName'] = filename
    df_title['FileName'] = filename

    return df, df_title

df = pd.DataFrame()
df_title = pd.DataFrame()

#df_tmp = extractor('Files2015/file-61.txt')
# Run extractor on all files under the directory
# Problematic file will be identified for further investgation: likely to be irregular format
indir = 'Files2015A/'
for root, dirs, filenames in os.walk(indir):
    for f in filenames:
        try:
            df_tmp, df_tmp_title = extractor(indir + f)
            if len(df_tmp.index) < 3 : print(f + ' - ', len(df_tmp.index))
            df = df.append(df_tmp)
            df_title = df_title.append(df_tmp_title)
        except Exception as e: 
            print(f + ' - ' + str(e))

##################################################################
#   Checking with entire DF
##################################################################

#Match title table with all content to identify error docs
# Correction will be made in the raw txt file. then run this part of the code

df_tmp = pd.merge(df_title[~df_title['Title'].isin(['Headlines:','Details:',
'Headlines','HEADLINES:','TITL​ES','FULL ARTICLES','﻿HEADLINES:','Details','﻿TRADE NEWS'])], df, how='left', on=['Title','FileName'])
df_tmp1 = df_tmp[df_tmp['Content'].isnull()]
df_tmp_count = df_tmp.groupby('FileName').size().reset_index(name='Count')
df_tmp1 = pd.merge(df_tmp1, df_tmp_count, on='FileName')




##################################################################
#   Add columns
##################################################################



## This is module to identified the dates from string
## Not used, use file name directly
#df3 = df.copy().reset_index()
## Extract date from article content
#for i in range(0, len(df3)):
#    matches = list(datefinder.find_dates(df3.loc[i, 'Content'][0:100]))
#    if len(matches) > 0:
#        # date returned will be a datetime.datetime object. here we are only using the first match.
#        try:                
##                print(matches[0])
#            df3.loc[i, 'Date'] = matches[0]
#        except:
#            df3.loc[i, 'Date'] = np.NaN
#    else:
#        df3.loc[i, 'Date'] = np.NaN
   
df['Date'] = df['FileName'].str[11:21]

df=df.reset_index()

# Extract agency from article content
#agencies = ['Interfax','The Hindu','The Western Mail','POLITICO','Financial Times','Taipei Times','Agence Europe',
#            'Business Line (The Hindu)','Business Standard','MintAsia','Bloomberg Newsweek','Politico',
#            'The Washington Post','Nikkei Report','Kyodo News','Deutsche Presse-Agentur','Xinhua News Agency',
#            'New York Times', 'MintAsia',  'The Washington Post', 'Nikkei Report','Kyodo News',
#            'Deutsche Presse-Agentur', 'Bulletin Quotidien Europe', 'Forbes.com','Reuters News',
#            'Bloomberg News','South China Morning Post','Investopedia','Sputnik News','BelTA',
#            'Mint','The Hans India','Agence France Presse','Unian','Taipei Times',
#            'Inside U.S. Trade','The Baltic Course','ITAR-TASS World Service','All Africa',
#            'Implications for LDCs','The Financial','All Africa',
#            'Business Standard','Bulletin Quotidien Europe','Times Of Oman','Business Times Singapore',
#            'The Hindu','Domain-B','LaPresseAffaires.com','The Times of India','Business Line (The Hindu)',
#            'India Blooms News Service ','NDTV ','Millennium Post','Sputnik News ','Yale Global Online',
#            'Mondaq Business Briefing','China.org.cn (China)','Peoples Daily','News International',
#            ]
#
#for agency in agencies:
#    df3.loc[df3['Content'].str.contains(agency), 'Author'] = agency

# Find languages
import langdetect 
for i in range(0, len(df)):
    df.loc[i, 'Language'] = langdetect.detect(df.loc[i, 'Content'][0:200])

df4 = df[~df['Language'].isin(['en', 'es','fr'])]
































#df.to_csv('ArticleTable.txt', sep='§')

##################################################################
#   Other Codes
##################################################################

