import os
os.getcwd()
os.chdir(r"F:\Projects\pandas-news-mining")
os.chdir(r"C:\Users\yu\Projects\pandas-news-mining-master")


###################################################################################################

import pandas as pd
import numpy as np
#import datefinder
#import win32com.client
import codecs
import re
#import sqlite3 as db
import pyodbc
#import itertools

'''######################################################################################
    Extract articles from a TXT file and convert to DF
    Extractor read and parse a TXT file.
    Titles of news are listed on the top of email as an index. Extractor matches 
        headings in index with each title to identify individual article.
    There will be some articles having problems to match. Open the TXT file and correct irregular texts. Then run this process again.
    Progress: data done: 2015,
######################################################################################'''


#################################################
# Extractor - get data from individual txt
#################################################

def extractor(filename): 
#    filename = 'Files2015A/2015-06-06.txt'
    file = codecs.open(filename, 'r', encoding='ascii', errors='ignore')
    lines = [line.strip() for line in file if line.strip()]
    file.close()
    
    df_raw = pd.DataFrame(lines, columns=['Texts']).reset_index().rename(columns={'index':'ParaID'})
    
    # Clean texts
    # Remove bulletes
#    df_raw['Texts'] = df_raw['Texts'].str.replace(r'•\t', '')
#    df_raw['Texts'] = df_raw['Texts'].str.replace(r'•', '')
#    df_raw['Texts'] = df_raw['Texts'].str.replace(r'·', '')
#    df_raw['Texts'] = df_raw['Texts'].str.replace(r'·', '')
#    df_raw['Texts'] = df_raw['Texts'].str.replace(r'·', '')
#    df_raw['Texts'] = df_raw['Texts'].str.replace(r'¡P', '')
#    df_raw['Texts'] = df_raw['Texts'].str.replace(r'-', '')
#    df_raw['Texts'] = df_raw['Texts'].str.replace(r'·', '')
#    df_raw['Texts'] = df_raw['Texts'].str.replace(r'·', '')    
#    df_raw['Texts'] = df_raw['Texts'].str.replace(r' +', ' ')
#    df_raw['Texts'] = df_raw['Texts'].str.replace(r'’s', '')
    
#    df_raw['Texts'] = df_raw['Texts'].str.replace(r'\n', '')
#    df_raw['Texts'] = df_raw['Texts'].str.replace(r'\r', '')

    # Replace double spaces with single
    df_raw['Texts'] = df_raw['Texts'].str.replace(r'  ', ' ')
    
    # Remove any email address
    for i in range(0, len(df_raw)):
        match = re.findall(r'[\w\.-]+@[\w\.-]+', df_raw.loc[i, 'Texts'])
        if len(match) > 0: df_raw.loc[i, 'Texts'] = df_raw.loc[i, 'Texts'].replace(match[0], '')
        
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

# Initialize new dataframes: df and df_title
df = pd.DataFrame()
df_title = pd.DataFrame()

# Run extractor on all files under the directory
# Problematic file will be identified for further investgation: likely to be irregular format
indir = 'Files2016ASCII/'
for root, dirs, filenames in os.walk(indir):
    for f in filenames:
        try:
            df_tmp, df_tmp_title = extractor(indir + f)
            # Print file with less than 2 items, possible errors
            if len(df_tmp.index) < 3 : print(f + ' - ', len(df_tmp.index))
            df = df.append(df_tmp)
            df_title = df_title.append(df_tmp_title)
        except Exception as e: 
            print(f + ' - ' + str(e))

#################################################
#   Checking the entire DF
#################################################

'''
After running the code above to convert TXT files to DF, there will be a lot of files cannot be coverted properly.
In 2016 the style for the exmail changed. Additional text added next to the title on top the emails, like below.
These TXT files need to be edited manually. The code in this section is to facilitate the manual work.

For specific example below: additional text added after ":" need to be removed!! then run code above and below over and over till 
all problems resovled.

EXAMPLE:
Angry voters were made on factory floors: In 1989, a few months before the Berlin Wall fell, ...

'''





# Match title table with all content to identify error docs
# Correction will be made in the raw txt file. then run this part of the code
# In the end, df_tocheck_problem should have 0 rows

df_tocheck = pd.merge(df_title[~df_title['Title'].isin(['Headlines:','Details:','Headlines','HEADLINES:',
                           'TITLES','FULL ARTICLES','HEADLINES:','Details','TRADE NEWS'])], 
                            df, how='left', on=['Title','FileName'])
df_tocheck_problem = df_tocheck[df_tocheck['Content'].isnull()]
df_tocheck_count = df_tocheck.groupby('FileName').size().reset_index(name='Count')
df_tocheck_problem = pd.merge(df_tocheck_problem, df_tocheck_count, on='FileName')

df_tocheck_files = df_tocheck_problem.groupby('FileName').Count.size().reset_index(name='Count')
df_tocheck_problem_ncount = pd.merge(df_tocheck_problem,df_tocheck_files[df_tocheck_files['Count']==1],on='FileName')
#################################################
#   Add columns: date, agencies and language
#################################################



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

df = df.reset_index()
df['Date'] = df['FileName'].str[11:21]
             
# Detect languages
import langdetect 
for i in range(0, len(df)):
    df.loc[i, 'Language'] = langdetect.detect(df.loc[i, 'Content'][0:200])
    
# Check items with languages other than EN, FR, ES
#df4 = df[~df['Language'].isin(['en', 'es','fr'])]


#################################################
# Prepare the dataframe for analysis
# Normalize terms, remove useless words
#################################################         

# Save and read documents to/from pickle
#df.to_pickle('df_save.pck')
df = pd.read_pickle('df_save.pck')

# Only work on English documents
df4 = df[df['Language'].isin(['en'])].copy()

# Combine title and contents
df4['Text'] = df['Title'] + ' ' + df['Content']

# Normalize words
df4['Text'] = df4['Text'].str.replace(r'\'s', '')
df4['Text'] = df4['Text'].str.replace('Indian', 'India')
df4['Text'] = df4['Text'].str.replace('nextgeneration', 'next generation')
df4['Text'] = df4['Text'].str.replace('//iconnect\.wto\.org/', '')
df4['Text'] = df4['Text'].str.replace('-', ' ')
df4['Text'] = df4['Text'].str.replace('U.S.', 'United States')
df4['Text'] = df4['Text'].str.replace('US', 'United States')

df4['Text'] = df4['Text'].str.replace('S.Korea', 'South Korea')
df4['Text'] = df4['Text'].str.replace('S. Korea', 'South Korea')
df4['Text'] = df4['Text'].str.replace('WTO', 'world trade organization')
df4['Text'] = df4['Text'].str.replace('‘', '')
df4['Text'] = df4['Text'].str.replace('imports', 'import')
df4['Text'] = df4['Text'].str.replace('Imports', 'import')
df4['Text'] = df4['Text'].str.replace('exports', 'export')
df4['Text'] = df4['Text'].str.replace('Exports', 'export')
df4['Text'] = df4['Text'].str.replace('NZ ', 'New Zealand ')
df4['Text'] = df4['Text'].str.replace('\"', '')
df4['Text'] = df4['Text'].str.replace('\'', '')

#df4['Text'] = df4['Text'].str.replace('U.S.', 'United States')

'''######################################################################################

    Clustering analysis referenced from "Document Clustering with Python"
    * use clusters to identify categories

######################################################################################'''

import nltk
from bs4 import BeautifulSoup
from sklearn import feature_extraction
#import mpld3

# Prepare lists
lst_texts = df4['Text'].tolist()
lst_titles = df4['Title'].tolist()
lst_dates = df4['Date'].tolist()
lst_articlecodes = df4['ArticleCode'].astype(int).tolist()

    
# load nltk's English stopwords as variable called 'stopwords'
my_stop_words = nltk.corpus.stopwords.words('english')

# Stop words normal
my_stop_words = my_stop_words + ['world_trade_organization','years','year','said','important',
                                 'new','would','','','','']

# Stop words excluding country names  NOT WORKING
#my_stop_words = my_stop_words + ['world_trade_organization','years','year','said','important','new','would','united_states',
#                                   'japan','india','obama','canada','mexico','russia','eu','european','china','chinese','would']

# MWETokenizer can attach words together
from nltk.tokenize import MWETokenizer
tokenizer = MWETokenizer([('world', 'bank'), ('world', 'trade', 'organization'), ('doha', 'round'),
                          ('united', 'states'), ('european', 'union'), ('new', 'zealand'),
                          ('per', 'cent'),('south', 'korea'),
                          ])
# Test the tokenizer
#tokenizer.tokenize('In a little or a little bit world trade organization'.split())
# Test the function
#tokenize_and_stem('In World Bank or a_little. bit  _ World Trade Organization. United States')

## load nltk's SnowballStemmer as variabled 'stemmer'
#from nltk.stem.snowball import SnowballStemmer
#stemmer = SnowballStemmer("english")

# Use WordNetLemmatizer instead of stemmer
#from nltk.stem import WordNetLemmatizer
#lemmatizer = WordNetLemmatizer()
#df['Lemmatized'] = df['StopRemoved'].apply(lambda x: [lemmatizer.lemmatize(y) for y in x])

#import string

# Define a tokenizer and stemmer which returns the set of stems in the text that it is passed
def tokenize_and_stem(text):
    # Remove punctuation
#    text = text.translate(str.maketrans('','',string.punctuation))

    # first tokenize by sentence, then by word to ensure that punctuation is caught as it's own token
    tokens = [word for sent in nltk.sent_tokenize(text) for word in nltk.word_tokenize(sent)]
#    tokens = [word for sent in nltk.sent_tokenize(text) for word in tokenizer.tokenize(sent.lower().split())]
    
    # MWETokenizer: manually link words, when disabled, use n-gram range in TF-IDF
#    tokens = tokenizer.tokenize(text.lower().split())
    tokens = tokenizer.tokenize(tokens)
    

    filtered_tokens = []
    # filter out any tokens not containing letters (e.g., numeric tokens, raw punctuation)
    for token in tokens:
        if (re.search('[a-zA-Z]', token)): 
            filtered_tokens.append(token)
    stems = filtered_tokens
#    stems = [stemmer.stem(t) for t in filtered_tokens]
    # WordNetLemmatizer
#    stems = [lemmatizer.lemmatize(t) for t in filtered_tokens]
    return stems


#def tokenize_only(text):
#    # Remove punctuation
##    text = text.translate(str.maketrans('','',string.punctuation))
#
#    # first tokenize by sentence, then by word to ensure that punctuation is caught as it's own token
##    tokens = [word.lower() for sent in nltk.sent_tokenize(text) for word in nltk.word_tokenize(sent)]
##    tokens = [word for sent in nltk.sent_tokenize(text) for word in nltk.word_tokenize(sent)]
##    tokens = [word for sent in nltk.sent_tokenize(text) for word in tokenizer.tokenize(sent.lower().split())]
#
#    # MWETokenizer: manually link words, when disabled, use n-gram range in TF-IDF
#    tokens = tokenizer.tokenize(text.lower().split())
#
#
#    filtered_tokens = []
#    # filter out any tokens not containing letters (e.g., numeric tokens, raw punctuation)
#    for token in tokens:
#        if re.search('[a-zA-Z]', token):
#            filtered_tokens.append(token)
#    return filtered_tokens
#
#totalvocab_stemmed = []
#totalvocab_tokenized = []
#for i in texts:
#    allwords_stemmed = tokenize_and_stem(i)
#    totalvocab_stemmed.extend(allwords_stemmed)
#    
#    allwords_tokenized = tokenize_only(i)
#    totalvocab_tokenized.extend(allwords_tokenized)

df4['Wordsss'] = df4['Text'].apply(tokenize_and_stem)

dddddd=df4[df4['Text'].str.contains(' ha ')]
    
########################################################################
# Test  
#df4['Words'] = df4['Text'].apply(tokenize_and_stem)
#df_word = df['Words'].apply(pd.Series)
#df_word = df_word.stack().to_frame()
#df_word.columns = ['POSTagged']
    
    
    
vocab_frame = pd.DataFrame({'words': totalvocab_tokenized}, index = totalvocab_stemmed)


# Check and identify wrong words
vocab_frame_group = vocab_frame.groupby('words').size()


#################################################
##    Tf-idf and document similarity
#################################################

from sklearn.feature_extraction.text import TfidfVectorizer

tfidf_vectorizer = TfidfVectorizer(max_df=0.9, max_features=200000,
                                   min_df=0.1, stop_words=my_stop_words, 
                                   use_idf=True, tokenizer=tokenize_and_stem, ngram_range=(1,3))

%time tfidf_matrix = tfidf_vectorizer.fit_transform(lst_texts)

print(tfidf_matrix.shape)

terms = tfidf_vectorizer.get_feature_names()

from sklearn.metrics.pairwise import cosine_similarity
dist = 1 - cosine_similarity(tfidf_matrix)

#################################################
##    K-means clustering
#################################################

from sklearn.cluster import KMeans
num_clusters = 15
km = KMeans(n_clusters=num_clusters)
%time km.fit(tfidf_matrix)
clusters = km.labels_.tolist()

df_tfidf_matrix = pd.DataFrame(tfidf_matrix.toarray())

#################################################
#from sklearn.externals import joblib
#
##joblib.dump(km,  'doc_cluster.pkl')
#km = joblib.load('doc_cluster.pkl')
#clusters = km.labels_.tolist()
#################################################

#################################################
# Clustering results to DF
#################################################

news = {'date': lst_dates,'articlecode': lst_articlecodes,
        'title': lst_titles,'text': lst_texts,'cluster': clusters}
frame = pd.DataFrame(news, index = [clusters], columns = ['date','articlecode','title','text','cluster'])

frame['cluster'].value_counts()

# Export to Access to analyze clusters
frame.to_csv('newsclusters.txt', sep='^')


# print top words of each cluster
print("Top terms per cluster:")

order_centroids = km.cluster_centers_.argsort()[:, ::-1]
for i in range(num_clusters):
    print()
    pterms=''
#    print("%d\t " % i, end='')
    for ind in order_centroids[i, :6]:
#        print(' %s' % vocab_frame.ix[terms[ind].split(' ')].values.tolist()[0][0], end=',')
#        print(' %s' % terms[ind], end=',')        
        pterms = pterms + terms[ind]+', '
    print(pterms)
    
#    print("Cluster %d titles:" % i, end='')
#    for title in frame.ix[i]['title'].values.tolist():
#        print(' %s ¦¦ ' % title, end='')
#    print()
#    print()

    

###########################################################
# Hierarchical document clustering
###########################################################

import matplotlib.pyplot as plt
import matplotlib as mpl

from scipy.cluster.hierarchy import ward, dendrogram

linkage_matrix = ward(dist) #define the linkage_matrix using ward clustering pre-computed distances

fig, ax = plt.subplots(figsize=(15, 20)) # set size
ax = dendrogram(linkage_matrix, orientation="right", labels=titles);

plt.tick_params(\
    axis= 'x',          # changes apply to the x-axis
    which='both',      # both major and minor ticks are affected
    bottom='off',      # ticks along the bottom edge are off
    top='off',         # ticks along the top edge are off
    labelbottom='off')

plt.tight_layout() #show plot with tight layout

#uncomment below to save figure
plt.savefig('ward_clusters.png', dpi=200) #save figure as ward_clusters



###########################################################
# Latent Dirichlet Allocation
###########################################################

#strip any proper names from a text...unfortunately right now this is yanking the first word from a sentence too.
import string
def strip_proppers(text):
    # first tokenize by sentence, then by word to ensure that punctuation is caught as it's own token
    tokens = [word for sent in nltk.sent_tokenize(text) for word in nltk.word_tokenize(sent) if word.islower()]
    return "".join([" "+i if not i.startswith("'") and i not in string.punctuation else i for i in tokens]).strip()


from gensim import corpora, models, similarities 

#remove proper names
%time preprocess = [strip_proppers(doc) for doc in lst_texts]

#tokenize
%time tokenized_text = [tokenize_and_stem(text) for text in preprocess]

#remove stop words
%time texts = [[word for word in text if word not in my_stop_words] for text in tokenized_text]


#create a Gensim dictionary from the texts
dictionary = corpora.Dictionary(texts)

#remove extremes (similar to the min/max df step used when creating the tf-idf matrix)
dictionary.filter_extremes(no_below=1, no_above=0.8)

#convert the dictionary to a bag of words corpus for reference
corpus = [dictionary.doc2bow(text) for text in texts]

# took 43 min to finish
%time lda = models.LdaModel(corpus, num_topics=20, id2word=dictionary, update_every=5, chunksize=10000, passes=100)


################################################
from sklearn.externals import joblib
#
joblib.dump(lda,  'lda.pkl')
##km = joblib.load('doc_cluster.pkl')
##clusters = km.labels_.tolist()
################################################


lda.show_topics()

topics_matrix = lda.show_topics(formatted=False, num_words=20)
topics_matrix1 = np.array(topics_matrix, dtype=float)

topic_words = topics_matrix[:,:,1]
for i in topic_words:
    print([str(word) for word in i])
    print()




































# CVwctorize

from sklearn.feature_extraction.text import CountVectorizer
docs = ['this this this book',
        'this cat good',
        'cat good shit']
count_model = CountVectorizer(ngram_range=(1,1)) # default unigram model
X = count_model.fit_transform(df4['Text'])

print(X)

Xc = (X.T * X) # this is co-occurrence matrix in sparse csr format
Xc.setdiag(0) # sometimes you want to fill same word cooccurence to 0
print(Xc.todense()) # print out matrix in dense format


















