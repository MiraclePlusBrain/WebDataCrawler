from email import header
from selenium import webdriver
from tkinter.filedialog import askopenfilename
import pandas as pd
import numpy as np
import time
from selenium.webdriver.common.proxy import *
from selenium.webdriver.firefox.firefox_profile import FirefoxProfile
from selenium.webdriver.firefox.options import Options
from selenium.common.exceptions import NoSuchElementException,WebDriverException,InvalidSessionIdException,TimeoutException
import requests
import csv
import os

# 当前运行的代码文件所在文件夹路径
cur_path = os.path.dirname(__file__)

def check_input_file(file_path):
    '''
    检查输入的源文件是否有效，标准: 表格形式文件（目前仅支持xlsx文件）

    :param file_path: str
    :return: DataFrame
    '''
    # 支持的文件类型tuple
    support_file_types = ('xlsx',)

    # 判断文件格式
    if not file_path.endswith(support_file_types):  # 判断文件格式是否为xlsx
        print('输入的文件必须为xlsx文件!')
        return None  # 如果文件格式不为xlsx文件，返回None

    # 获取excel文件的内容
    df_raw_data = pd.read_excel(file_path, header=None)
    return df_raw_data

# 获取源文件的路径
file_path = askopenfilename(title='输入招聘关键词的源数据')
#导入外部数据
base1=pd.read_excel(file_path,header=None)
#设置存储结果的frame
employ1=pd.DataFrame()

#http://http.tiqu.letecs.com/getip3?num=1&type=1&pro=0&city=0&yys=0&port=1&time=1&ts=0&ys=0&cs=0&lb=4&sb=0&pb=4&mr=1&regions=
#k表示着使用ip的序号,开始从第一个用起.每爬取一页就换一个

def buildbrowser(url):
    fp=webdriver.FirefoxProfile()
    #从ip代理商的api中获取ip
    ip_api=requests.get('http://http.tiqu.alibabaapi.com/getip3?num=1&type=1&neek=566573&port=1&lb=4&pb=45&gm=4&regions=')
    ip_api=ip_api.text.split('\n')
    ip_pool=ip_api[0:-1]
    #暂时使用第一个ip
    proxy = ip_pool[0]
    print('更换ip:'+ip_pool[0])
    ip, port = proxy.split(":")
    port = int(port)
    # 设置代理配置
    fp.set_preference('network.proxy.type', 1)
    fp.set_preference('network.proxy.http', ip)
    fp.set_preference('network.proxy.http_port', port)
    fp.set_preference('network.proxy.ssl', ip)
    fp.set_preference('network.proxy.ssl_port', port)
    fp.update_preferences()
    options=Options()
    #设置动态ip
    browser=webdriver.Firefox(firefox_profile=fp,options=options)
    #设置隐式等待
    browser.implicitly_wait(5)
    #页面最大加载时间35秒
    browser.set_page_load_timeout(35)
    try:
        browser.get(url)
        time.sleep(2)
        webelem=browser.find_elements('class name','job-title')
    except (WebDriverException,InvalidSessionIdException,TimeoutException):
        browser.quit()
        webelem,browser=buildbrowser(url)
    else:pass
    if webelem==[]:
        browser.quit()
        webelem,browser=buildbrowser(url)
    else:pass
    return webelem,browser

for p in range(base1.shape[0]):
    #下面开始正式爬取,构造url
    url="https://www.zhipin.com/c101010100/s_301/?query="+base1.iloc[p,0]+"&ka=sel-scale-301"
    index=True
    i=0
    ##这里要加入一个选择器来选择“融资阶段”
    webElem2,browser=buildbrowser(url)
    while index:
        #设置代理ip
        #开始爬取
        #time.sleep(5)
        #爬取一页的部分
        try:
            webElem2=browser.find_elements('class name','job-title')
            while i < len(webElem2):
                time.sleep(1)
                webElem2_2=webElem2[i]
                browser.execute_script("arguments[0].click();", webElem2_2)
                # time.sleep(8)
                handle=browser.window_handles
                time.sleep(1.5)
                #搜集数据
                browser.switch_to.window(handle[1])
                #开始提取数据
                try:
                    webElem3_1=browser.find_element('xpath','/html/body/div[1]/div[2]/div[3]/div/div[1]/div[2]/p[2]').text
                except NoSuchElementException:    
                    webElem3_1='提取不到元素'
                else:pass
                try:
                    webElem3_2=browser.find_element('xpath','/html/body/div[1]/div[2]/div[3]/div/div[1]/div[2]/p[3]').text
                except NoSuchElementException:    
                    webElem3_2='提取不到元素'
                else:pass
                try:
                    webElem3_3=browser.find_element('xpath','/html/body/div[1]/div[2]/div[3]/div/div[1]/div[2]/p[4]').text
                except NoSuchElementException:    
                    webElem3_3='提取不到元素'
                else:pass
                try:
                    webElem3_4=browser.find_elements('class name','job-sec')[2].find_element(by='tag name',value='div').text
                except NoSuchElementException:    
                    webElem3_4='提取不到元素'
                else:pass
                try:
                    webElem3_5=browser.find_elements('class name','job-sec')[3].find_element(by='tag name',value='div').text
                except (NoSuchElementException,IndexError):    
                    webElem3_5='提取不到元素'
                else:pass
                try:
                    webElem3_6=browser.find_elements('class name','job-sec')[4].find_element(by='tag name',value='div').text
                except (NoSuchElementException,IndexError):    
                    webElem3_6='提取不到元素'
                else:pass
                #将上面采集的数据整合到一起                                                                                   
                content=np.array([webElem3_1,webElem3_2,webElem3_3,webElem3_4,webElem3_5,webElem3_6]).reshape(1,6)
                tempo=pd.DataFrame(content,columns=['当前阶段','职工人数','行业赛道','工商注册名称1','工商注册名称2','工商注册名称3'])
                print(tempo,employ1.shape)
                employ1=pd.concat([employ1,tempo],axis=0)
                browser.close()
                #切换浏览器窗口
                browser.switch_to.window(handle[0])
                i+=1
        except (WebDriverException,IndexError,InvalidSessionIdException):
            browser.quit()
            webElem2,browser=buildbrowser(url)
            i+=1
            continue
        else:
            try:
            ## 翻页按钮
                url_pre=browser.current_url
                webElem1_next=browser.find_element('class name','next').click()
                time.sleep(10)
                url=browser.current_url
                ##设置停止的点
                if url_pre==url:
                    index=False
                    browser.quit()
                i=0
            except NoSuchElementException:
                browser.quit()
                index=False
            else:pass
       
    url="https://www.zhipin.com/c101010100/s_302/?query="+base1.iloc[p,0]+"&ka=sel-scale-302"
    index=True
    i=0
    ##这里要加入一个选择器来选择“融资阶段”
    webElem2,browser=buildbrowser(url)
    while index:
        #设置代理ip
        #开始爬取
        #time.sleep(5)
        #爬取一页的部分
        try:
            webElem2=browser.find_elements('class name','job-title')
            while i < len(webElem2):
                time.sleep(1)
                webElem2_2=webElem2[i]
                browser.execute_script("arguments[0].click();", webElem2_2)
                # time.sleep(8)
                handle=browser.window_handles
                time.sleep(1.5)
                #搜集数据
                browser.switch_to.window(handle[1])
                #如果后面几个都报错的话是不是后面的就看不到了_是的!
                try:
                    webElem3_1=browser.find_element('xpath','/html/body/div[1]/div[2]/div[3]/div/div[1]/div[2]/p[2]').text
                except NoSuchElementException:    
                    webElem3_1='提取不到元素'
                else:pass
                try:
                    webElem3_2=browser.find_element('xpath','/html/body/div[1]/div[2]/div[3]/div/div[1]/div[2]/p[3]').text
                except NoSuchElementException:    
                    webElem3_2='提取不到元素'
                else:pass
                try:
                    webElem3_3=browser.find_element('xpath','/html/body/div[1]/div[2]/div[3]/div/div[1]/div[2]/p[4]').text
                except NoSuchElementException:    
                    webElem3_3='提取不到元素'
                else:pass
                try:
                    webElem3_4=browser.find_elements('class name','job-sec')[2].find_element(by='tag name',value='div').text
                except NoSuchElementException:    
                    webElem3_4='提取不到元素'
                else:pass
                try:
                    webElem3_5=browser.find_elements('class name','job-sec')[3].find_element(by='tag name',value='div').text
                except (NoSuchElementException,IndexError):    
                    webElem3_5='提取不到元素'
                else:pass
                try:
                    webElem3_6=browser.find_elements('class name','job-sec')[4].find_element(by='tag name',value='div').text
                except (NoSuchElementException,IndexError):    
                    webElem3_6='提取不到元素'
                else:pass
                                                                                                    
                content=np.array([webElem3_1,webElem3_2,webElem3_3,webElem3_4,webElem3_5,webElem3_6]).reshape(1,6)
                tempo=pd.DataFrame(content,columns=['当前阶段','职工人数','行业赛道','工商注册名称1','工商注册名称2','工商注册名称3'])
                print(tempo,employ1.shape)
                employ1=pd.concat([employ1,tempo],axis=0)
                browser.close()
                browser.switch_to.window(handle[0])
                i+=1
        except (WebDriverException,IndexError,InvalidSessionIdException):
            browser.quit()
            webElem2,browser=buildbrowser(url)
            i+=1
            continue
        else:
            try:
            ## 翻页按钮
                url_pre=browser.current_url
                webElem1_next=browser.find_element('class name','next').click()
                time.sleep(10)
                url=browser.current_url
                ##设置停止的点
                if url_pre==url:
                    index=False
                i=0
            except NoSuchElementException:
                browser.quit()
                index=False
            else:pass
    
    employ1.to_excel("123_"+str(p)+"_.xlsx")



