# -*- coding:utf-8 -*-
import xlrd
import xlsxwriter
from collections import Counter
from itertools import chain
from janome.tokenizer import Tokenizer

#https://qiita.com/d-cabj/items/d934eb87e3012a02e23a
#↑ここ参考にした。

# VSCodeでデバッグ。ステップアウトなどなど

# Excel読み込み
#book=xlrd.open_workbook(r"Excelのパス")
book=xlrd.open_workbook(r"yahoo.xlsx")

# Excelワークシートの1枚目（0番目）を変数に格納
sheet_1=book.sheet_by_index(0)
col=0
data=[]

each_data=[]

print("------------ Debug Start ------------------\n")

t=Tokenizer()

for row in range(sheet_1.nrows):
    val=sheet_1.cell(row,col).value
    tokens=t.tokenize(val)
    for token in tokens:
        partOfSpeech=token.part_of_speech.split(',')[0]
        #print(partOfSpeech)
        # 今回抽出するのは名詞だけ
        if partOfSpeech ==u'名詞':
            each_data.append(token.surface)
        if partOfSpeech ==u'動詞':
            each_data.append(token.surface)
    data.append(each_data)
    each_data=[]


# 文章を形態素毎に分割したデータをいれるエクセルファイル作成
#data_book=xlsxwriter.Workbook(r"書き込むエクセルファイルのパス")
data_book=xlsxwriter.Workbook(r"yahoo.w.xlsx")
data_sheet=data_book.add_worksheet('data')
for row in range(len(data)):
    for i in range(len(data[row])):
        data_sheet.write(row,i,data[row][i])
# エクセルを保存

data_book.close()

# すべての語彙を同じ配列に格納

chain_data=list(chain.from_iterable(data))
c=Counter(chain_data)

result_ranking=c.most_common(100)  # 出現頻度100までを格納

# 出現頻度ランキングをエクセルファイルに保存
#result_book=xlsxwriter.Workbook(r"エクセルファイルのパス（ランキング用））")
result_book=xlsxwriter.Workbook(r"yahoo.r.xlsx")
# commit
# resultという名前のワークシートをつくる
result_sheet=result_book.add_worksheet('result')
for row in range(len(result_ranking)):
    for i in range(len(result_ranking[row])):
        result_sheet.write(row,i,result_ranking[row][i])

# エクセルを保存
result_book.close()