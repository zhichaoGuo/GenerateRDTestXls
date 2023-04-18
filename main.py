import json
import logging

import selenium

from selenium.common.exceptions import *
from RD_utils import *


def run_RD_test_xls(start_time=None, end_time=None, log_level='DEBUG'):
    # LOG格式设置，如无法识别log的等级，则默认为warning
    LOG_FORMAT = "%(asctime)s - %(levelname)s - %(message)s"
    if log_level == 'DEBUG':
        log_level = logging.DEBUG
    elif log_level == 'INFO':
        log_level = logging.INFO
    elif log_level == 'ERROR':
        log_level = logging.ERROR
    else:
        log_level = logging.WARNING

    logging.basicConfig(filename='./log/' + str(generate_time()) + '.log', level=log_level, format=LOG_FORMAT)
    logging.info('today is ' + str(generate_time()))
    if (start_time is None) & (end_time is None):
        # 不输入起始时间也不输入终止时间，按照今日研发测试进行搜索
        start_time = generate_select_time(generate_time())
        select_end_time = generate_time()
    elif (start_time is not None) & (end_time is not None):
        select_end_time = end_time
    else:
        start_time = generate_time()
        select_end_time = generate_time()
        print('不可只输入start time or end time')

    # 目前暂时这样设计，后面增加开始截至时间时再优化

    logging.info("today will select " + generate_select_time(generate_time()) + 'to ' + str(select_end_time))
    # 取cookies
    # get_cookies()
    # 执行主要操作
    w_excel(gen_fix_dict(start_time, select_end_time))



def upgrade_driver(src_dir):
    print('--------------------------------------------------')
    import requests  # 请求并下载相应的msedgedriver版本
    import zipfile  # 用于解压下载文件
    import os  # 检查目录下的文件
    import re  # 对信息进行正则匹配
    import xml.dom.minidom  # 处理包含浏览器版本信息的xml文件
    from selenium import webdriver
    from selenium.webdriver.edge.options import Options

    def unzip(file_dir, out_dir):
        zf = zipfile.ZipFile(file_dir)  # 实例化压缩文件
        try:
            zf.extract('msedgedriver.exe', path=out_dir)  # 解压文件
        except RuntimeError as e:
            print(e)
        zf.close()

    if os.path.isfile(src_dir + 'msedgedriver.exe'):  # 查看是否有msedgedriver.exe文件
        pass
    else:
        dom = xml.dom.minidom.parse(r'C:\Program Files (x86)\Microsoft\Edge\Application\msedge'
                                    r'.VisualElementsManifest.xml')  # 读取edge文件夹下面的xml文件(包含版本信息)
        ve_text = dom.documentElement.getElementsByTagName('VisualElements')[0].toxml()  # 包含版本号的字符串文本
        rematch = re.match(r'(.*)\"(.*)\\VisualElements\\Logo.png', ve_text)
        edge_version = rematch.group(2)  # 匹配得到版本号
        url = 'https://msedgedriver.azureedge.net/' + edge_version + '/edgedriver_win64.zip'
        response = requests.get(url=url)  # 请求edgedriver下载链接
        file_dir = src_dir + edge_version + 'edgedriver_win64.zip'
        open(file_dir, 'wb').write(response.content)  # 下载edgedriver压缩包
        unzip(file_dir, src_dir)  # 在下载目录下解压edgedriver压缩包
        if os.path.isfile(file_dir):
            os.remove(file_dir)
        else:
            pass
    try:
        options = Options()
        options.add_argument('headless')
        webdriver.Edge(options=options)
    # except SessionNotCreatedException as msg:
    except selenium.common.exceptions.WebDriverException as msg:
        reg = re.search("(.*)headless MicrosoftEdge=(.*)\)", str(msg))  # 识别并匹配Exception信息中出现的版本号
        edge_version = reg.group(2)  # 获得版本号
        url = 'https://msedgedriver.azureedge.net/' + edge_version + '/edgedriver_win64.zip'
        print(url)
        response = requests.get(url=url)
        file_dir = src_dir + edge_version + '.zip'
        open(file_dir, 'wb').write(response.content)  # 下载压缩文件
        unzip(file_dir, src_dir)  # 解压文件
        if os.path.isfile(file_dir):
            os.remove(file_dir)  # 将多余的压缩文件删除
        else:
            pass
    pass

if __name__ == '__main__':
    # # LOG格式设置
    # LOG_FORMAT = "%(asctime)s - %(levelname)s - %(message)s"
    # logging.basicConfig(filename='./log/' + str(generate_time()) + '.log', level=logging.DEBUG, format=LOG_FORMAT)
    # logging.info('today is ' + str(generate_time()))
    # # 目前暂时这样设计，后面增加开始截至事件时再优化
    # select_end_time = generate_time()
    # logging.info("today will select " + generate_select_time(generate_time()) + 'to ' + str(select_end_time))
    # 取cookies

    # webdriver = 'C:\\Users\\admin\\AppData\\Local\\Programs\\Python\\Python37\\'
    # try:
    #     get_cookies()
    # except selenium.common.exceptions.WebDriverException:
    #     upgrade_driver(webdriver)
    #     get_cookies()

    # download ：https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/
    # # 执行主要操作
    # w_excel(gen_fix_dict(generate_select_time(generate_time())))
    run_RD_test_xls(log_level='DEBUG')
    # run_RD_test_xls('2023-4-4','2023-4-6',log_level='DEBUG')
    # a = b'\xd5\xd2\xb2\xbb\xb5\xbd\xd6\xb8\xb6\xa8\xb5\xc4\xc4\xa3\xbf\xe9\xa1\xa3'
    # print(a.decode('gbk'))


