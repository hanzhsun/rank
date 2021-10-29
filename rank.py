import os
import xlrd
import jieba
jieba.set_dictionary("./dict.txt")
jieba.initialize()
from textrank4zh import TextRank4Keyword



for f in os.listdir("./"):
	if os.path.splitext(f)[1] == '.xlsx':
		wb = xlrd.open_workbook(f)
		sheet = wb.sheet_by_index(0)
		cols = sheet.col_values(sheet.ncols-7)
		tr4w = TextRank4Keyword()
		tr4w.analyze(text=" ".join(cols))
		with open('./result.txt','w+') as f0: 
			for item in tr4w.get_keywords(20, word_min_len=2):
				print(item,file=f0)