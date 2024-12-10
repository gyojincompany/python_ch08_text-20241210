# nltk, pandas, matplotlib, wordcloud, xlrd 5개 패키지 설치 필요
import nltk
import glob

import pandas as pd

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




