from operator import index
from tkinter.filedialog import askopenfilename
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
import pandas as pd
import numpy as np
import json
import os
import sys

# 当前运行的代码文件所在文件夹路径
cur_path = os.path.dirname(__file__)


def check_input_file(file_path):
    '''
    检查输入的源文件是否有效，有两个标准: 表格形式文件（目前仅支持xlsx文件）且有一列列头为“公司注册名称”

    :param file_path: str
    :return: DataFrame
    '''
    # 支持的文件类型tuple
    support_file_types = ('xlsx',)
    # 项目支持的搜索主键列表
    search_keys = ('公司注册名称', )

    # 判断文件格式
    if not file_path.endswith(support_file_types):  # 判断文件格式是否为xlsx
        print('输入的文件必须为xlsx文件!')
        return None  # 如果文件格式不为xlsx文件，返回None

    # 获取excel文件的内容
    df_raw_data = pd.read_excel(file_path, index_col=False)

    # 判断是否存在列头“公司注册名称”
    if any([(key in df_raw_data.columns) for key in search_keys]):  # 只需要任何一个支持的键列在表头中即可
        return df_raw_data
    else:
        print('文件必须有一列头为“公司注册名称”。')
        return None  # 内容不包含“公司注册名称”列，返回None


def init_browser():
    '''启动浏览器'''
    browser = webdriver.Firefox()  # 最好驱动firefox，而不要使用chrome（新版chrome已经不支持selenium）

    # 设置参数
    browser.implicitly_wait(3)  # 页面加载等待时长：3s

    return browser


def save_cookies(browser):
    '''
    保存cookies
    '''
    browser.get(url)
    # 输入手机号密码
    cookies = browser.get_cookies()
    fw = open('tianyancookies.txt', 'w')
    json.dump(cookies, fw)
    fw.close()


def login(browser):
    '''
    登录天眼查账户
    :param browser: class
    :return: class
    '''
    # 登入天眼查
    browser.get('https://www.tianyancha.com/')
    # 先卡住主线程
    input('正在登录中……（登录完成后Enter即可恢复工作）')

    def login_with_cookies():
        '''利用cookies登录'''
        nonlocal browser
        # 确认是否有cookies

        # 登录天眼查页面
        url = 'https://www.tianyancha.com/'
        browser.get(url)

        # 载入cookies登录
        fr = open('tianyancookies.txt', 'r+')
        cookieslist = json.load(fr)
        fr.close()
        for cookie in cookieslist:
            browser.add_cookie(cookie)

        # 刷新网页
        browser.refresh()

    # 预期方案是用cookies登录，但暂时一直出错，先使用手动登录
    # login_with_cookies()

    return browser


# 获取源文件的路径
file_path = askopenfilename(title='输入源文件')
# 获取有效的源数据DataFrame
df_raw_data = check_input_file(file_path=file_path)
# 如果源数据不合规，则退出程序
if df_raw_data is None:
    print('源数据不合规。正在退出程序……')
    sys.exit(1)

# 开启浏览器
browser = init_browser()
# 登录天眼查账户
login(browser)

# 准备查询的公司名称列表
companies = df_raw_data['公司注册名称'].to_list()
# 声明一个储存结果DataFrame
df_output = pd.DataFrame()
# 声明一个有效数据数量的计数器
cnt_valid_data = 0

# 下面开始正式采集
for i in range(len(companies)):
    # 公司名称
    company_name = companies[i]

    # 公司名称为空的情况，跳过这一行。e.g. 有其他列的信息，但没有公司名称
    if pd.isna(company_name):
        continue

    # 开始进行爬取
    url = 'https://www.tianyancha.com/search?key='+company_name
    browser.get(url)
    # 锁定项目名元素
    try:
        browser.find_element('id', 'user_mobile')
    except NoSuchElementException:
        judge = False
    else:
        judge = True
    # 识别到验证码的时候退出循环
    if judge:
        df_output.to_excel('MPB_融资数据补充_'+str(i)+'.xlsx')
        print('需要手动验证，此时的断点是'+str(i))
        break
    # 真正开始提取元素了
    try:
        projectname = browser.find_element(
            'class name', 'sv-search-company-brand').find_element('class name', 'title')
    except NoSuchElementException:
        print('该公司无融资信息')
        continue
    else:
        pass
    projectname2 = projectname.text
    projectname.click()
    handle = browser.window_handles
    # 切换focus
    browser.switch_to.window(handle[1])

    # 公司注册名称
    try:
        businame = browser.find_elements('id', 'project_web_top')[
            0].find_element('class name', 'link-click').text
    except (NoSuchElementException, IndexError):
        print('%i、【%s】无融资信息' % (i, company_name))
        browser.close()
        browser.switch_to.window(handle[0])
        continue
    else:
        pass
    # 行业赛道
    try:
        sector = browser.find_elements('id', 'project_web_top')[
            0].find_element('class name', 'tag-common').text
    except (NoSuchElementException, IndexError):
        sector = '无'
    else:
        pass
    # 简介
    try:
        description = browser.find_elements('id', 'project_web_top')[
            0].find_element('class name', 'detail-content').text
    except (NoSuchElementException, IndexError):
        description = '无'
    else:
        pass
    # 最近一次融资日期
    try:
        finandate = browser.find_elements(
            'class name', 'block-data')[0].find_elements('tag name', 'span')[0].text
    except (NoSuchElementException, IndexError):
        finandate = '无'
    else:
        pass
    # 融资金额
    try:
        finannum = browser.find_elements(
            'class name', 'block-data')[0].find_elements('tag name', 'span')[1].text
    except (NoSuchElementException, IndexError):
        finannum = '无'
    else:
        pass
    # 当前轮次
    try:
        finanround = browser.find_elements(
            'class name', 'block-data')[0].find_elements('tag name', 'span')[2].text
    except (NoSuchElementException, IndexError):
        finanround = '不明'
    else:
        pass
    # 历史投资方
    try:
        financer = browser.find_elements(
            'class name', 'block-data')[0].find_elements('class name', 'touzi')
    except (NoSuchElementException, IndexError):
        financer = '无'
    else:
        pass
    financer2 = ','.join([(i.text) for i in financer])

    content = np.array([businame, projectname2, sector, description,
                       finandate, finannum, finanround, financer2]).reshape(1, 8)
    tempo = pd.DataFrame(content, columns=[
                         '公司注册名称', '项目名称', '行业赛道', '简介', '上一轮融资至今时长', '融资金额', '当前轮次', '历史投资方'])
    # 合并当前公司的融资信息至总表
    df_output = pd.concat([df_output, tempo], axis=0)
    # 有效数据+1
    cnt_valid_data += 1
    # 控制台进度提示
    print('%i、已完成【%s】的融资信息爬取。' % (i, company_name))

    # 切换到下一个页面
    browser.close()
    browser.switch_to.window(handle[0])

# 创建输出文件夹
if not os.path.exists(os.path.join(cur_path, '..', '爬取结果')):
    os.mkdir(os.path.join(cur_path, '..', '爬取结果'))

# 编写文件输出路径
export_path = os.path.join(cur_path, '..', '爬取结果',
                           'MPB_融资数据补充_%i.xlsx' % cnt_valid_data)  # i记录了爬取条数
# 导出结果文件
df_output.to_excel(export_path, index=None)
print('信息爬取完成，输出文件: %s' % export_path)

# 输出可优化的空间，按照大牌的资本排行格式化输出：红杉(A：千万级美金)、蓝驰(天使)……