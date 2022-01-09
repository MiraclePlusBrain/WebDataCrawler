from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
import pandas as pd
import numpy as np
import json
import xlsxwriter

#导入数据
tianyan1=pd.DataFrame()

base1=pd.read_excel('MPB_output.xlsx',index_col=False)
base1.columns
i=0
list(base1[['公司注册名称']].iloc[i])[0]
#设置链接
#开启网页模拟器
#这里咱们先手动登陆一下，
browser = webdriver.Firefox()
'''
browser.get(url)
#输入手机号密码
cookies=browser.get_cookies()
fw=open('tianyancookies.txt', 'w')
json.dump(cookies, fw)
fw.close()
'''
i=1
url='https://www.tianyancha.com/search?key='+list(base1[['公司注册名称']].iloc[i])[0]
#导入cookies
browser.get(url)
fr=open('tianyancookies.txt', 'r')
cookieslist=json.load(fr)
fr.close()
for cookie in cookieslist:
    browser.add_cookie(cookie)
#刷新网页
browser.refresh()
browser.implicitly_wait(6)
#其实不用10秒，可以再短一点

#下面开始正式采集
for i in range(3119,base1.shape[0]):
    url='https://www.tianyancha.com/search?key='+list(base1[['公司注册名称']].iloc[i])[0]
    browser.get(url)
    #锁定项目名元素
    try:  
        browser.find_element('id','user_mobile')
    except NoSuchElementException:
        judge=False
    else:
        judge=True
    #识别到验证码的时候退出循环
    if judge:
        tianyan1.to_excel('MPB_融资数据补充_'+str(i)+'.xlsx')
        print('需要手动验证，此时的断点是'+str(i))
        break
    #真正开始提取元素了
    try: 
        projectname=browser.find_element('class name','sv-search-company-brand').find_element('class name','title')
    except NoSuchElementException:
        print('该公司无融资信息')
        continue
    else:pass
    projectname2=projectname.text
    projectname.click()
    handle=browser.window_handles
    #切换focus
    browser.switch_to.window(handle[1])

    #公司注册名称
    try:
        businame=browser.find_elements('id','project_web_top')[0].find_element('class name','link-click').text
    except (NoSuchElementException,IndexError):
        print('该公司无融资信息')
        browser.close()
        browser.switch_to.window(handle[0])
        continue
    else:pass
    #行业赛道
    try:
        sector=browser.find_elements('id','project_web_top')[0].find_element('class name','tag-common').text
    except (NoSuchElementException,IndexError):
        sector='无'
    else:pass
    #简介
    try:
        description=browser.find_elements('id','project_web_top')[0].find_element('class name','detail-content').text
    except (NoSuchElementException,IndexError):
        description='无'
    else:pass
    #最近一次融资日期
    try:
        finandate=browser.find_elements('class name','block-data')[0].find_elements('tag name','span')[0].text
    except (NoSuchElementException,IndexError):
        finandate='无'
    else:pass
    #融资金额
    try:
        finannum=browser.find_elements('class name','block-data')[0].find_elements('tag name','span')[1].text
    except (NoSuchElementException,IndexError):
        finannum='无'
    else:pass
    #当前轮次
    try:
        finanround=browser.find_elements('class name','block-data')[0].find_elements('tag name','span')[2].text
    except (NoSuchElementException,IndexError):
        finanround='不明'
    else:pass
    #历史投资方
    try:
        financer=browser.find_elements('class name','block-data')[0].find_elements('class name','touzi')
    except (NoSuchElementException,IndexError):
        financer='无'
    else:pass
    financer2=','.join([(i.text) for i in financer])

    content=np.array([businame,projectname2,sector,description,finandate,finannum,finanround,financer2]).reshape(1,8)
    tempo=pd.DataFrame(content,columns=['公司注册名称','项目名称','行业赛道','简介','上一轮融资至今时长','融资金额','当前轮次','历史投资方'])
    print(tempo)
    tianyan1=pd.concat([tianyan1,tempo],axis=0)
    browser.close()
    browser.switch_to.window(handle[0])

tianyan1.to_excel('MPB_融资数据补充_'+str(i)+'.xlsx',engine='xlsxwriter')


