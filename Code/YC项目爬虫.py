from selenium import webdriver
import pandas as pd
import numpy as np
import os

# 当前运行的代码文件所在文件夹路径
cur_path = os.path.dirname(__file__)

def init_browser():
    '''启动浏览器'''
    browser = webdriver.Firefox()  # 最好驱动firefox，而不要使用chrome（新版chrome已经不支持selenium）

    # 设置参数
    browser.implicitly_wait(3)  # 页面加载等待时长：3s

    return browser

url1 = 'https://www.ycombinator.com/companies/?batch=S17&batch=W17&batch=IK12&batch=W16&batch=S16&batch=S15&batch=W15&batch=W18&batch=S18&team_size=%5B%221%22%2C%2250%22%5D'
url2 = 'https://www.ycombinator.com/companies/?batch=W19&batch=S19&batch=W20&batch=S20&team_size=%5B%221%22%2C%2250%22%5D'
url3 = 'https://www.ycombinator.com/companies/?batch=W21&batch=S21&batch=W22&team_size=%5B%221%22%2C%2250%22%5D'

# 启动浏览器
browser = init_browser()

# 储存数据
df_output = pd.DataFrame()
browser.get(url3)
js_bottom = "var q=document.documentElement.scrollTop=200000"

# 设置翻页,直到翻到底部
while browser.find_element('class name', 'styles-module__rightCol___2NKRr').find_elements('tag name', 'div')[-1].text == 'Loading more...':
    browser.execute_script(js_bottom)

element1s = browser.find_elements(
    'class name', 'styles-module__coName___3zz21')
element2s = browser.find_elements(
    'class name', 'styles-module__coLocation___yhKam')
for i in range(len(element1s)):
    content = np.array([element1s[i].text, element2s[i].text]).reshape(1, 2)
    tempo = pd.DataFrame(content, columns=['项目名称', '城市'])
    df_output = pd.concat([df_output, tempo], axis=0)

# 创建输出文件夹
if not os.path.exists(os.path.join(cur_path, '..', '爬取结果')):
    os.mkdir(os.path.join(cur_path, '..', '爬取结果'))

# 编写文件输出路径
export_path = os.path.join(cur_path, '..', '爬取结果',
                           'MPB_YC2_%i.xlsx' % len(element1s))  # i记录了爬取条数
# 导出结果文件
df_output.to_excel(export_path, index=None)
print('信息爬取完成，输出文件: %s' % export_path)
