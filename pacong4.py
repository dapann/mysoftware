# -*- coding:utf-8 -*-
#
# 读取本地html，提取目标位置及页数
#


import urllib
import re
from bs4 import BeautifulSoup


path="E:/002.html"
d={}
# html=urllib.urlopen(url).read()											# 将网页源代码赋予menuCode

html=open(path).read()

soup=BeautifulSoup(html,'lxml')

contents=soup.find_all("th") #"span",class_="tps"

for content in contents:
	if content.a!=None:
		key=content.a['href'].split('-')[1]
		if content.span!=None:
			if str(content.span.get_text()).strip()[-1]!=None and content.span.get_text()!="已关闭":
				ment=int(str(content.span.get_text()).strip()[-1])
				d[key]=ment

		else:
			d[key]=1

	# print content.get_text()
for key in sorted(d):
	print("'"+key+"':"+str(d[key])+',')