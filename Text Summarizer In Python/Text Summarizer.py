import docx2txt
import re
import heapq
import nltk
import docx



text= docx2txt.process("Text Input.docx")

#Removing thw Square Brackets and Extra Spaces
article_text = re.sub(r'\[[0-9]*\]', ' ', text)
article_text = re.sub(r'\s+', ' ',article_text )

#Sentence tokenization
sentence_list = nltk.sent_tokenize(article_text)

#Now we need to find the frequency of each word before it was tokenized into sentences
#Word frequency
#Also we are removing the stopwords from the text we are using

stopwords = nltk.corpus.stopwords.words('english')

#Storing the word frequencies in a dictionary

word_frequencies = {}
for word in nltk.word_tokenize(article_text):
    if word not in stopwords:
        if word not in word_frequencies.keys():
            word_frequencies[word] = 1
        else:
            word_frequencies[word] += 1
            
#Now we need to find the weighted frequency
#We shall divide the number of occurances of all the words by the frequency of the most occurring word            
            
maximum_frequncy = max(word_frequencies.values())

for word in word_frequencies.keys():
    word_frequencies[word] = (word_frequencies[word]/maximum_frequncy)
    
#Using relative word frequency, not absolute word fequency so as to distribute the values from 0-1
    
#Calculating Sentence Scores
#Scores for each sentence obtained by adding weighted frequencies of the words that occur in that particular sentence.

sentence_scores = {}
for sent in sentence_list:
    for word in nltk.word_tokenize(sent.lower()):
        if word in word_frequencies.keys():
            if len(sent.split(' ')) < 25:
                if sent not in sentence_scores.keys():
                    sentence_scores[sent] = word_frequencies[word]
                else:
                    sentence_scores[sent] += word_frequencies[word]
                    
#To summarize the article we are going to take the sentences with 10 highest scores 
#This parameter can be changed according to the length of the text
                    
                    

summary_sent = heapq.nlargest(10, sentence_scores, key=sentence_scores.get)

summary = ' '.join(summary_sent)
print(summary)


mydoc = docx.Document()
mydoc.add_paragraph(summary)
mydoc.save("Text Output.docx")






