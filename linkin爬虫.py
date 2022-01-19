from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from tkinter.filedialog import askopenfilename
import pandas as pd
import numpy as np
import json
import re
import time
import os
import sys

# 当前运行的代码文件所在文件夹路径
cur_path = os.path.dirname(__file__)


def check_input_file(file_path):
    '''
    检查输入的源文件是否有效，有两个标准: 表格形式文件（目前仅支持xlsx文件）且有一列列头为“公司名称”

    :param file_path: str
    :return: DataFrame
    '''
    # 支持的文件类型tuple
    support_file_types = ('xlsx',)
    # 项目支持的搜索主键列表
    search_keys = ('公司名称', )

    # 判断文件格式
    if not file_path.endswith(support_file_types):  # 判断文件格式是否为xlsx
        print('输入的文件必须为xlsx文件!')
        return None  # 如果文件格式不为xlsx文件，返回None

    # 获取excel文件的内容
    df_raw_data = pd.read_excel(file_path, index_col=False)

    # 判断是否存在列头“公司名称”
    if any([(key in df_raw_data.columns) for key in search_keys]):  # 只需要任何一个支持的键列在表头中即可
        return df_raw_data
    else:
        print('文件必须有一列头为“公司名称”。')
        return None  # 内容不包含“公司名称”列，返回None


def init_browser():
    '''启动浏览器'''
    browser = webdriver.Firefox()  # 最好驱动firefox，而不要使用chrome（新版chrome已经不支持selenium）

    # 设置参数
    browser.implicitly_wait(3)  # 页面加载等待时长：3s

    return browser


def login(browser):
    '''
    登录天眼查账户
    :param browser: class
    :return: class
    '''
    # 登入天眼查
    browser.get('https://www.linkedin.com/company/snackpass/about')
    # 先卡住主线程
    input('正在登录中……（登录完成后Enter即可恢复工作）')

    def login_with_cookies():
        '''利用cookies登录'''
        nonlocal browser
        # 确认是否有cookies

        url = 'https://www.linkedin.com/company/snackpass/about'
        # 导入cookies
        browser.get(url)
        fr = open('linkincookies.txt', 'r')
        cookieslist = json.load(fr)
        fr.close()
        for cookie in cookieslist:
            browser.add_cookie(cookie)

    # 预期方案是用cookies登录，但暂时一直出错，先使用手动登录
    # login_with_cookies()

    return browser


# 获取源文件的路径
file_path = askopenfilename(title='输入源文件')
# 获取有效的源数据DataFrame
base1 = check_input_file(file_path=file_path)
# 如果源数据不合规，则退出程序
if base1 is None:
    print('源数据不合规。正在退出程序……')
    sys.exit(1)

# 开启浏览器
browser = init_browser()
# 登录linkedin
login(browser)

# 准备查询的公司名称列表
companies = base1['公司名称'].to_list()
# 声明一个储存结果DataFrame
linkin1 = pd.DataFrame()
# 声明一个有效数据数量的计数器
cnt_valid_data = 0

# #这一段在第一次采集的时候需要创建一个名为linkincookies的txt空文件,然后运行
# #第二次运行就不用了
# url='https://www.linkedin.com/home'
# browser.get(url)
# #输入手机号密码
# cookies=browser.get_cookies()
# fw=open('linkincookies.txt', 'w')
# json.dump(cookies, fw)
# fw.close()


# 开始采集(循环内内容)
i = 0
for i in range(len(companies)):
    projectname = companies[i]

    # 公司名称为空的情况，跳过这一行。e.g. 有其他列的信息，但没有公司名称
    if pd.isna(projectname):
        continue

    url = 'https://www.linkedin.com/search/results/companies/?keywords=' + \
        projectname+'&origin=SWITCH_SEARCH_VERTICAL&sid=9Gv'
    browser.get(url)
    # 点击第一个搜索结果
    try:
        suburl1 = browser.find_elements(
            'class name', 'app-aware-link')[0].get_attribute('href')
        projectname2 = browser.find_elements(
            'class name', 'entity-result__title-text')[0].text
    except IndexError:
        continue
    else:
        pass
    browser.get(suburl1)
    # 采集赛道和地点
    try:
        topbar = browser.find_element(
            'class name', 'org-top-card-summary-info-list').text
        topbar2 = str.split(topbar, ' ')
        # 行业赛道和城市
        career = topbar2[0]
        city = topbar2[1]
    except (NoSuchElementException, IndexError):
        career = ''
        city = ''
    else:
        pass
    # 采集基本的联系方式和融资信息
    try:
        stage = browser.find_element(
            'class name', 'org-funding-info__content').find_elements('class name', 't-14')[0].text
    except NoSuchElementException:
        stage = ''
    else:
        pass
    try:
        fimoney = browser.find_element(
            'class name', 'org-funding-info__content').find_elements('class name', 't-14')[1].text
    except (NoSuchElementException, IndexError):
        fimoney = ''
    else:
        pass
    # 采集其它信息
    suburl2 = suburl1+'about/'
    browser.get(suburl2)
    try:
        onesentence = browser.find_element(
            'class name', 'org-top-card-summary__tagline').text
    except NoSuchElementException:
        onesentence = ''
    else:
        pass
    try:
        content = browser.find_elements(
            'class name', 'artdeco-card')[2].find_element('class name', 't-black--light').text
    except (NoSuchElementException, IndexError):
        continue
    else:
        pass
    # others包含很混杂的信息，需要清洗
    others = browser.find_elements(
        'class name', 'artdeco-card')[2].find_element('class name', 'overflow-hidden').text
    otherslist = str.split(others, '\n')
    # print(otherslist)
    # others信息提取
    website = ''
    comsize = ''
    starttime = ''
    tel = ''
    for k in range(int(len(otherslist))):
        # 公司网站
        website += otherslist[k+1] if otherslist[k] == '公司网站' else ''
        # 公司规模
        comsize += otherslist[k+1] if otherslist[k] == '公司规模' else ''
        # 创立时间
        starttime += otherslist[k+1] if otherslist[k] == '创立时间' else ''
        # 手机号
        tel += str(otherslist[k+1]) if otherslist[k] == '联系电话' else ''

    # 采集页面融资信息
    # 暂时没有那么需要
    # 缺失信息,需要检测元素是否存在
    suburl3 = suburl1+'people/?keywords=CEO'
    browser.get(suburl3)
    # 锁定所有能采集的链接
    js_bottom = "var q=document.documentElement.scrollTop=2000"
    browser.execute_script(js_bottom)
    time.sleep(3)
    try:
        element1 = browser.find_elements(
            'class name', 'link-without-visited-state')[2:-1]
    except NoSuchElementException:
        continue
    else:
        pass
    ceoname = ''
    ceourl = ''
    ceoidentity = ''
    otherurls = ''
    element1 = browser.find_elements(
        'class name', 'link-without-visited-state')[2:]
    # 循环收集所有链接
    for j in range(len(element1)):
        # 名字
        name = element1[j].text
        # 领英链接
        nameurl = element1[j].get_attribute('href')
        # 主要联系人链接
        try:
            otherurls += nameurl+';'
        except TypeError:
            continue
        else:
            pass
        idname1 = re.sub('\d', '', element1[j].get_attribute(
            'id'))+str(int(re.sub('\D', '', element1[j].get_attribute('id')))+2)
        idname2 = re.sub('\d', '', element1[j].get_attribute(
            'id'))+str(int(re.sub('\D', '', element1[j].get_attribute('id')))+4)
        # 人名title
        nametitle1 = browser.find_element('id', idname1).text
        nametitle2 = browser.find_element('id', idname2).text
        # 判断哪个是我们要的
        if nametitle1.find('CEO') != -1:
            ceoname = name
            ceourl = nameurl
            ceoidentity = nametitle1
        if nametitle2.find('CEO') != -1:
            ceoname = name
            ceourl = nameurl
            ceoidentity = nametitle2

    content = np.array([ceoname, ceoidentity, projectname2, onesentence, starttime, career,
                       content, stage, fimoney, website, tel, ceourl, otherurls, comsize, city]).reshape(1, 15)
    tempo = pd.DataFrame(content, columns=['姓名', '第一联系人身份', '项目名称', '一句话介绍', '成立时间',
                         '行业赛道', '资料', '当前轮次', '融资金额', '官网', '手机号', '第一联系人链接', '主要联系人链接', '公司规模', '城市'])
    # 合并当前公司的融资信息至总表
    linkin1 = pd.concat([linkin1, tempo], axis=0)
    # 有效数据+1
    cnt_valid_data += 1

# 创建输出文件夹
if not os.path.exists(os.path.join(cur_path, '..', '爬取结果')):
    os.mkdir(os.path.join(cur_path, '..', '爬取结果'))

# 编写文件输出路径
export_path = os.path.join(cur_path, '..', '爬取结果',
                           'MPB_linkedin爬取_%i.xlsx' % cnt_valid_data)  # i记录了爬取条数
# 导出结果文件
linkin1.to_excel(export_path, index=None)
print('信息爬取完成，输出文件: %s' % export_path)
