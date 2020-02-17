# -*- coding:utf-8 -*- 
# 读取本地html，提取目标位置及页数
# 
import xlwings as xw
from urllib import request
import os,random,re,time
from bs4 import BeautifulSoup
import socket
import telnetlib

User_Agent = 'Mozilla/5.0 (Windows NT 6.3; WOW64; rv:43.0) Gecko/20100101 Firefox/43.0'
header = {}
header['User-Agent'] = User_Agent
    
def getProxyIp(): 
    proxy = []
    
    try:
        url = "http://www.xicidaili.com/nn/1"
        req = request.Request(url,headers=header)  
        res = request.urlopen(req).read() 
        soup = BeautifulSoup(res,'lxml')  
        ips = soup.find_all('tr')  
        for x in range(1,len(ips)):  
            ip = ips[x]  
            tds = ip.find_all("td")  
            ip_temp = tds[1].contents[0]+":"+tds[2].contents[0]  
            proxy.append(ip_temp)  
    except:
        print('failed')
    return proxy
     
''''' 
验证获得的代理IP地址是否可用 
'''  
def validateIp(proxy):
    url = "http://ip.chinaz.com/getip.aspx"
    validip=[]  
    # f = open("E:\ip.txt","w")  
    socket.setdefaulttimeout(3)
    for i in range(0,len(proxy)):  
        try:  
            ip = proxy[i].strip().split("\t")  
            telnetlib.Telnet(ip[0], port=ip[1], timeout=20)
        except:
            # print(ip[0]+':connect failed')
            continue 
        else:
            validip.append(ip[0]+':'+ip[1])
    return validip
            # f.write(validip+'\n')          
    # f.close()
def use_proxy(url,proxy_addr):
	timeout = 30
	socket.setdefaulttimeout(timeout)
	try:
	    req=request.Request(url)
	    req.add_header("User-Agent","Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/49.0.2623.221 Safari/537.36 SE 2.X MetaSr 1.0")
	    proxy=request.ProxyHandler({'http':proxy_addr})
	    opener=request.build_opener(proxy,request.HTTPHandler)
	    request.install_opener(opener)
	    html=request.urlopen(req).read()
	    # print(htm.status)
	    data=html.decode('GBK','ignore')
	except:
		data='fail'
	return data 

my_headers = [  
		"Mozilla/5.0 (Windows NT6.1;WOW64;rv:27.0)Gecko/20100101 Firefox/27.0",  
		"Mozilla/5.0 (Windows;U;Windows NT 6.1;en-US;rv:1.9.1.6)Gecko/20091201 Firefox/3.5.6",  
		"Mozilla/5.0 (Windows;U;Windows NT 6.1)AppleWebKit/537.36(KHTML,like Gecko) Chrome/34.0.1838.2 Safari/",  
] 

app=xw.App(visible=True,add_book=False)					# 导入xlwings模块，打开Excel程序，默认设置：程序可见，只打开不新建工作薄，屏幕更新关闭
filepath=r'E:\pacong\tag\tag001.xlsx'						# 文件位置：filepath，打开test文档，然后保存，关闭，结束程序

wb=app.books.open(filepath)
shta=xw.books['tag001'].sheets['Sheet2']

i=593

while i <=594:						#594
	proxy = getProxyIp()
	print(proxy)
	
	if proxy!=[]:
		x=0
		while x<len(proxy):
			try:
				sCode=str(shta.range(i,1).value)
			 
				url="http://www.ngotcmszh.com/thread-"+sCode+"-1-1.html" 			#访问网址
				print("%i%s%i%s"%(x,'/',len(proxy),' ')+proxy[x]+' waiting')						
				html=use_proxy(url,proxy[x])

				if html!='fail':
					
					print(proxy[x]+' success')
					soup=BeautifulSoup(html,'lxml')  # 使用html解析器进行解析
					keywords=soup.find(attrs={"name":"keywords"})['content']
					
					if len(keywords)<=50:
						shta.range(i,3).value=keywords
						i+=1
						if i>594:
							break
					else:
						i-=1
						x+=1
				else:
					x+=1
			except:
				continue

wb.save()
wb.close()
app.quit()


