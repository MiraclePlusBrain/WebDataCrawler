from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
import pandas as pd
import numpy as np
import json
import re
import xlsxwriter
import time
#创建储存数据的框
linkin1=pd.DataFrame()
#导入数据
base1=pd.read_excel('source1.xlsx',index_col=False)
i=0
list(base1[['公司名称']].iloc[i])[0]

browser = webdriver.Firefox()

#这一段在第一次采集的时候需要创建一个名为linkincookies的txt空文件,然后运行
#第二次运行就不用了
'''
url='https://www.linkedin.com/home'
browser.get(url)
#输入手机号密码
cookies=browser.get_cookies()
fw=open('linkincookies.txt', 'w')
json.dump(cookies, fw)
fw.close()
'''

url='https://www.linkedin.com/company/snackpass/about'
#导入cookies
browser.get(url)
fr=open('linkincookies.txt', 'r')
cookieslist=json.load(fr)
fr.close()
for cookie in cookieslist:
    browser.add_cookie(cookie)
#刷新网页
browser.refresh()
browser.implicitly_wait(6)

#开始采集(循环内内容)
for i in range(132,int(base1.shape[0])):
    projectname=list(base1[['公司名称']].iloc[i])[0]
    url='https://www.linkedin.com/search/results/companies/?keywords='+projectname+'&origin=SWITCH_SEARCH_VERTICAL&sid=9Gv'
    browser.get(url)
    #点击第一个搜索结果
    try:
        suburl1=browser.find_elements('class name','app-aware-link')[0].get_attribute('href')
        projectname2=browser.find_elements('class name','entity-result__title-text')[0].text
    except IndexError:
        continue
    else:pass
    browser.get(suburl1)
    #采集赛道和地点
    try:
        topbar=browser.find_element('class name','org-top-card-summary-info-list').text
        topbar2=str.split(topbar,' ')
        #行业赛道和城市
        career=topbar2[0];city=topbar2[1]
    except (NoSuchElementException,IndexError):
        career='';city=''
    else:pass
    #采集基本的联系方式和融资信息
    try:
        stage=browser.find_element('class name','org-funding-info__content').find_elements('class name','t-14')[0].text
    except NoSuchElementException:
        stage=''
    else:pass
    try:
        fimoney=browser.find_element('class name','org-funding-info__content').find_elements('class name','t-14')[1].text
    except (NoSuchElementException,IndexError):
        fimoney=''
    else:pass
    #采集其它信息
    suburl2=suburl1+'about/'
    browser.get(suburl2)
    try:
        onesentence=browser.find_element('class name','org-top-card-summary__tagline').text
    except NoSuchElementException:
        onesentence=''
    else:pass
    try:
        content=browser.find_elements('class name','artdeco-card')[2].find_element('class name','t-black--light').text
    except (NoSuchElementException,IndexError):
        continue
    else:pass
    #others包含很混杂的信息，需要清洗
    others=browser.find_elements('class name','artdeco-card')[2].find_element('class name','overflow-hidden').text
    otherslist=str.split(others,'\n')
    #print(otherslist)
    #others信息提取
    website='';comsize='';starttime='';tel=''
    for k in range(int(len(otherslist))):
        #公司网站
        website+=otherslist[k+1] if otherslist[k]=='公司网站' else ''
        #公司规模
        comsize+=otherslist[k+1] if otherslist[k]=='公司规模' else ''
        #创立时间
        starttime+=otherslist[k+1] if otherslist[k]=='创立时间' else ''
        #手机号
        tel+=str(otherslist[k+1]) if otherslist[k]=='联系电话' else ''

    #采集页面融资信息
    #暂时没有那么需要
    #缺失信息,需要检测元素是否存在
    suburl3=suburl1+'people/?keywords=CEO'
    browser.get(suburl3)
    #锁定所有能采集的链接
    js_bottom = "var q=document.documentElement.scrollTop=2000"
    browser.execute_script(js_bottom)
    time.sleep(3)
    try:
        element1=browser.find_elements('class name','link-without-visited-state')[2:-1]
    except NoSuchElementException:
        continue
    else:pass
    ceoname='';ceourl='';ceoidentity='';otherurls=''
    element1=browser.find_elements('class name','link-without-visited-state')[2:]
    #循环收集所有链接
    for j in range(len(element1)):
        #名字
        name=element1[j].text
        #领英链接
        nameurl=element1[j].get_attribute('href')
        #主要联系人链接
        try:
            otherurls+=nameurl+';'
        except TypeError:
            continue
        else:pass
        idname1=re.sub('\d','',element1[j].get_attribute('id'))+str(int(re.sub('\D','',element1[j].get_attribute('id')))+2)
        idname2=re.sub('\d','',element1[j].get_attribute('id'))+str(int(re.sub('\D','',element1[j].get_attribute('id')))+4)
        #人名title
        nametitle1=browser.find_element('id',idname1).text
        nametitle2=browser.find_element('id',idname2).text
        #判断哪个是我们要的
        if nametitle1.find('CEO')!=-1:
            ceoname=name
            ceourl=nameurl
            ceoidentity=nametitle1
        if nametitle2.find('CEO')!=-1:
            ceoname=name
            ceourl=nameurl
            ceoidentity=nametitle2
            
    content=np.array([ceoname,ceoidentity,projectname2,onesentence,starttime,career,content,stage,fimoney,website,tel,ceourl,otherurls,comsize,city]).reshape(1,15)
    tempo=pd.DataFrame(content,columns=['姓名','第一联系人身份','项目名称','一句话介绍','成立时间','行业赛道','资料','当前轮次','融资金额','官网','手机号','第一联系人链接','主要联系人链接','公司规模','城市'])
    print(tempo,linkin1.shape)
    linkin1=pd.concat([linkin1,tempo],axis=0)

linkin1.to_excel('MPB_linkin爬取_'+str(i)+'.xlsx',engine='xlsxwriter')