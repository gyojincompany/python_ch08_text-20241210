# nltk, pandas, matplotlib, wordcloud, xlrd 5개 패키지 설치 필요
import nltk
import glob

import pandas as pd
import re  # 정규표현식
from nltk.tokenize import word_tokenize  # 단어 토큰화
from nltk.corpus import stopwords  # 불용어 묶음
from nltk.stem import WordNetLemmatizer  # 표제어 추출
from functools import reduce  # 1차원 리스트로 변환
from collections import Counter  # 빈도수 추출
from wordcloud import STOPWORDS, WordCloud  # 워드 클라우드 모듈

import matplotlib.pyplot as plt

# nltk.download()

all_files_Name = glob.glob("data/exportExcelData_*.xls")
# data 폴더 내 10개 엑셀 파일의 이름들을 모두 불러와서 list로 저장
print(all_files_Name)

all_files_data = []

for fileName in all_files_Name:
    data_frame = pd.read_excel(fileName)  # 저장된 실제 excel 파일불러오기
    all_files_data.append(data_frame)

# all_files_data.append(pd.read_excel("exportExcelData_20241210184725.xls"))
# pd.read_excel("exportExcelData_20241210184725.xls")
# pd.read_excel("exportExcelData_20241210184725.xls")
# pd.read_excel("exportExcelData_20241210184725.xls")


print(all_files_data[0])
print(all_files_data[1])

all_files_data_concat = pd.concat(all_files_data, axis=0, ignore_index=True)
# 10개의 엑셀 파일을 모두 병합하여 DataFrame으로 변환
print(all_files_data_concat)

all_title = all_files_data_concat["제목"]
print(all_title)

stopWords = set(stopwords.words("english"))  # 영어 불용어 묶음
lemma = WordNetLemmatizer()  # 표제어 추출을 위한 객체

words = []

for title in all_title:
    EnWords = re.sub(r"[^a-zA-z]+"," ",str(title))  # 영문자를 제외한 모든 문자를 제거
    EnWordsToken = word_tokenize(EnWords.lower())  # 영문자를 모두 소문자로 변경 후 토큰화
    EnWordsTokenStop = [word for word in EnWordsToken if word not in stopWords]  # 불용어를 제거한 단어 리스트
    # print(EnWordsTokenStop)
    EnWordsTokenStopLemma = [lemma.lemmatize(word) for word in EnWordsTokenStop]  # 표제어만 추출한 단어 리스트
    # print(EnWordsTokenStopLemma)
    words.append(EnWordsTokenStopLemma)

# print(words)
# print("---------------------------------------------------------------")

# words가 2차원 리스트 이므로, 1차원 리스트로 변환
words2 = list(reduce(lambda x, y: x+y, words))
# print(words2)

count = Counter(words2)
print(count)

word_count = dict()

for tag, counts in count.most_common(50):
    if(len(tag)>1):  # 단어의 길이가 1보다 큰 단어만 추출
        word_count[tag] = counts

# print(word_count) -> 단어의 길이가 2자 이상이고, 빈도수가 상위 50위 이상인 단어들이 dic 자료

# 검색어인 big과 data 단어를 제거
del word_count["big"]
del word_count["data"]


sorted_Keys = sorted(word_count, reverse=True, key=word_count.get)  # 빈도수의 내림차순으로 정렬된 단어들
# print(sorted_Keys)
sorted_Values = sorted(word_count.values(), reverse=True)  # 내림차순으로 정렬한 빈도수 값들

plt.bar(range(len(word_count)), sorted_Values, align="center")
plt.xticks(range(len(word_count)), list(sorted_Keys), rotation=85)
plt.show()

# 워드 클라우드 만들기
stopWords2 = set(STOPWORDS)  # 워드클라우드에서 제공하는 불용어 모음
wc = WordCloud(stopwords=stopWords2, width=800, height=600, background_color="ivory")
cloud = wc.generate_from_frequencies(word_count)
plt.imshow(cloud)
plt.axis("off")  # 워드클라우드 이미지에 출력되는 축 제거
plt.show()

cloud.to_file("data/eng_big_data_cloud.jpg")  # 워드 클라우드 이미지를 jpg파일로 저장


