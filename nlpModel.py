from locale import normalize
import numpy as np 
import re
import nltk
from nltk.corpus import stopwords
from nltk.stem.porter import PorterStemmer
from wordcloud import WordCloud, STOPWORDS
import matplotlib.pyplot as plt
import pandas as pd
import collections
import matplotlib.pyplot as plt

df = pd.read_excel('consolidated.xlsx')
allBrands = df['Brand'].unique()
corpus = []
worddict = {}
reviews = df['All Reviews']

for i in range(0, len(df['All Reviews'])):
    review = re.sub('[^a-zA-Z]', ' ', reviews[i])
    review = review.lower()
    review = review.split()
    ps = PorterStemmer()
    review = [ps.stem(word) for word in review
        if not word in set(stopwords.words('english'))]
    review = ' '.join(review)
    corpus.append(review)
comment_words = ""
stopwords = set(STOPWORDS)
for word in corpus:
    # typecaste each val to string
    word = str(word).lower()
    word = re.sub('mask', '', word)
    comment_words += word+' '
    print(comment_words)

wordcloud = WordCloud(width = 800, height = 800,
                background_color ='white',
                stopwords = stopwords,
                min_font_size = 10).generate(comment_words)
# plot the WordCloud image     
     
plt.figure(figsize = (8, 8), facecolor = None)
plt.imshow(wordcloud)
plt.axis("off")
plt.tight_layout(pad = 0)
plt.savefig('wordcloud_all.png')



# Read input file, note the encoding is specified here 
# It may be different in your text file

a= ' '.join(corpus)
# Stopwords
stopwords = set(line.strip() for line in corpus)
# Instantiate a dictionary, and for every word in the file, 
# Add to the dictionary if it doesn't exist. If it does, increase the count.
wordcount = {}
# To eliminate duplicates, remember to split by punctuation, and use case demiliters.
for word in a.lower().split():
    word = word.replace(".","")
    word = word.replace(",","")
    word = word.replace(":","")
    word = word.replace("\"","")
    word = word.replace("!","")
    word = word.replace("â€œ","")
    word = word.replace("â€˜","")
    word = word.replace("*","")
    if word not in stopwords:
        if word not in wordcount:
            wordcount[word] = 1
        else:
            wordcount[word] += 1
# Print most common word
n_print = int(input("How many most common words to print: "))
print("\nOK. The {} most common words are as follows\n".format(n_print))
word_counter = collections.Counter(wordcount)
for word, count in word_counter.most_common(n_print):
    print(word, ": ", count)
# Close the file

# Create a data frame of the most common words 
# Draw a bar chart
lst = word_counter.most_common(n_print)
df = pd.DataFrame(lst, columns = ['Word', 'Count'])
df.to_excel('wordcounts.xlsx')
