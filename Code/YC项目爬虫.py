from selenium import webdriver
import pandas as pd
import numpy as np
url1='https://www.ycombinator.com/companies/?batch=S17&batch=W17&batch=IK12&batch=W16&batch=S16&batch=S15&batch=W15&batch=W18&batch=S18&team_size=%5B%221%22%2C%2250%22%5D'
url2='https://www.ycombinator.com/companies/?batch=W19&batch=S19&batch=W20&batch=S20&team_size=%5B%221%22%2C%2250%22%5D'
url3='https://www.ycombinator.com/companies/?batch=W21&batch=S21&batch=W22&team_size=%5B%221%22%2C%2250%22%5D'
#储存数据
YC2=pd.DataFrame()

browser=webdriver.Firefox()
browser.get(url3)
js_bottom = "var q=document.documentElement.scrollTop=200000"

#设置翻页,直到翻到底部
while browser.find_element('class name','styles-module__rightCol___2NKRr').find_elements('tag name','div')[-1].text=='Loading more...':
    browser.execute_script(js_bottom)

element1s=browser.find_elements('class name','styles-module__coName___3zz21')
element2s=browser.find_elements('class name','styles-module__coLocation___yhKam')
for i in range(len(element1s)):
    content=np.array([element1s[i].text,element2s[i].text]).reshape(1,2)
    tempo=pd.DataFrame(content,columns=['项目名称','城市'])
    YC2=pd.concat([YC2,tempo],axis=0)
    
YC2.to_excel("MPB_YC2_.xlsx")