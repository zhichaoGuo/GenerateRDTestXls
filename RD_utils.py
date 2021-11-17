import json
import logging
from time import sleep

import requests
import yaml
from selenium import webdriver
from selenium.webdriver.common.by import By
import datetime
from datetime import date, timedelta
import openpyxl
from openpyxl.styles import Alignment, PatternFill


def get_cookies():
    """
    通过selenium模拟登陆的方式获取cookie
    :return:
    """
    with open("cfg.yaml", "r+") as f:
        root_url = yaml.safe_load(f)['root_url']
        f.close()
    # 打开edge浏览器
    driver = webdriver.Edge()
    logging.info("OPEN Edge")
    # 访问repo登陆界面
    driver.get(root_url + 'login/')
    logging.info("open" + root_url + "login/")
    # 隐式等待
    driver.implicitly_wait(10)
    # 取用户名密码
    logging.info("open password.yaml")
    with open("password.yaml", encoding="UTF-8") as f:
        logging.info("load data")
        user_data = yaml.safe_load(f)
        logging.info("load data over")
    f.close()
    logging.info("close password.yaml")
    # 输入用户名密码
    driver.find_element(By.ID, 'f_user').send_keys(user_data["username"])
    logging.info("input username " + user_data["username"])
    driver.find_element(By.ID, 'f_pass').send_keys(user_data["password"])
    logging.info("input password len is " + str(len(user_data["password"])))
    driver.find_element(By.ID, 'f_remember').click()
    logging.info("click remember me")
    # 点击登录
    driver.find_element(By.ID, 'b_signin').click()
    logging.info("click sign in")
    driver.get(root_url + 'changes/')
    logging.info("open" + root_url + "changes/")
    # 取cookies存yaml
    cookie = driver.get_cookies()
    logging.info("get cookies")
    with open("cookies.yaml", "w", encoding="UTF-8") as f:
        logging.info("open cookies.yaml")
        yaml.dump(cookie, f)
        logging.info("save cookies")
    f.close()
    logging.info("close cookies.yaml")
    sleep(1)
    # 关闭网页
    driver.quit()
    logging.info("quit driver")


def remove_same_key_value_list(list: list, key: str):
    """
    删除列表中字典里指定key中value相同的字典
    :param list:
    :param key:
    :return:
    """
    logging.debug('=============================================================================')
    logging.debug('remove same key value for list')
    temp = []
    result = []
    for i in range(len(list)):
        if list[i][key] not in temp:
            temp.append(list[i][key])
            logging.debug("保存记录 " + key + ':' + list[i][key])
            result.append(list[i])
        else:
            logging.info("移除重复 " + key + ':' + list[i][key])
    logging.debug('=============================================================================')
    return result


def gen_fix_dict(select_time, select_end_time='2023-11-11'):
    """
    整理修建从json中获取的dict
    :param select_end_time:
    :param select_time:
    :return:
    """
    # 打开cookies文件读取并编制好cookies
    with open("cookies.yaml", encoding="UTF-8") as f:
        yaml_data = yaml.safe_load(f)
        GerritAccount = yaml_data[1]['value']
        XSRF_TOKEN = yaml_data[0]['value']
        cookie = 'jenkins-timestamper-offset=-28800000; GerritAccount=' + GerritAccount + '; XSRF_TOKEN=' + XSRF_TOKEN
        logging.info("use cookies :" + cookie)
        f.close()
    cookies = {'Cookie': cookie}
    # 读取仓库根目录
    with open("cfg.yaml", "r+") as f:
        root_url = yaml.safe_load(f)['root_url']
    # 读取预制好的请求头
    with open("cfg.yaml", "r+") as f:
        headers = yaml.safe_load(f)['header']
        f.close()
    # 将cookies拼接到请求头上
    headers.update(cookies)
    # 拼接请求地址
    get_url = root_url + 'changes/?O=81&S=0&n=50&q=status%3Amerged%20after%3A' + select_time
    # 添加结束时间的url 暂时还没使用
    new_url = root_url + 'changes/?O=81&S=0&n=25&q=status%3Amerged%20after%3A' + select_time + '%20before%3A' + select_end_time
    logging.info('get ' + select_time + '数据')
    # 发送请求
    r = requests.get(get_url, headers=headers)
    logging.info('get request status code :' + str(r.status_code))
    # 整理响应数据删除头部符号
    json_text = r.text
    json_text = json_text[4:-1]
    data_list = json.loads(json_text)
    # 在log中记录data_list
    logging.debug('=============================================================================')
    logging.debug('get data is')
    logging.debug(data_list)
    logging.debug('=============================================================================')
    # 定义需要删除的key值并删除
    del_list = ['_number', 'branch', 'created', 'deletions', 'has_review_started', 'hashtags', 'id', 'insertions',
                'labels', 'owner', 'requirements', 'status', 'submitted', 'submitter', 'total_comment_count',
                'unresolved_comment_count', 'updated']
    logging.debug('=============================================================================')
    logging.debug('removed key list is')
    logging.debug(del_list)
    logging.debug('=============================================================================')
    logging.debug('start delete')
    for i in range(len(data_list)):
        for j in del_list:
            logging.debug('delete ' + str(data_list[i][j]))
            del data_list[i][j]
    logging.debug('=============================================================================')
    # 删除change_id重复的项
    data_list = remove_same_key_value_list(data_list, 'change_id')
    logging.info('=============================================================================')
    logging.info('整理后数据为')
    logging.info(data_list)
    logging.info('=============================================================================')
    # # 存整理后的数据 （因添加至log中故放弃存为过程数据）
    # with open("fix_info_dict.yaml", "w", encoding="UTF-8") as f:
    #     yaml.dump(data_list, f, allow_unicode=True)
    #     f.close()
    logging.info('data_list is ' + str(data_list))
    return data_list


def generate_time():
    """
    获取当日日期 ->str  2021-11-01
    :return:
    """
    today = date.today()
    return today


def generate_select_time(today):
    """
    根据传入日期推算验单日期（如周一则瑞算至周五）->str 2021-10-29
    :param today:str
    :return select_time:str
    """
    if today.isoweekday() == 1:
        delete_day = datetime.timedelta(days=3)
    else:
        delete_day = datetime.timedelta(days=1)
    select_time = today - delete_day
    return str(select_time)


def w_excel(data_list: list):
    """
    通过修建后的list生成工作表
    :param data_list:
    """
    # 新建工作簿与工作表
    wb = openpyxl.Workbook()
    wb.create_sheet(index=0, title='sheet1')
    sheet1 = wb.worksheets[0]
    # 设置列宽
    sheet1.column_dimensions['A'].width = 43
    sheet1.column_dimensions['B'].width = 5
    sheet1.column_dimensions['C'].width = 55
    sheet1.column_dimensions['D'].width = 20
    # 设置表头格式
    sheet1.merge_cells('A1:D1')
    sheet1.cell(1, 1).fill = PatternFill("solid", fgColor="6BA9E6")
    title_name = generate_select_time(generate_time()) + ' PATCH 信息'
    sheet1.cell(1, 1, title_name)
    alignment_center = Alignment(horizontal='center', vertical='center')
    sheet1.cell(1, 1).alignment = alignment_center
    fill = PatternFill("solid", fgColor="80A2DA")
    fill1 = PatternFill("solid", fgColor="EAF6E1")
    # 设置表头样式
    sheet1.cell(2, 1, 'change_id').fill = fill
    sheet1.cell(2, 2, 'type').fill = fill
    sheet1.cell(2, 3, 'JIRA ID and DESC').fill = fill
    sheet1.cell(2, 4, 'Comment').fill = fill
    # 循环写入内容
    for i in range(len(data_list)):
        sheet1.cell(i + 3, 1, data_list[i]['change_id']).fill = fill1
        sheet1.cell(i + 3, 2).fill = fill1
        sheet1.cell(i + 3, 3, data_list[i]['subject']).alignment = Alignment(wrapText=True)
        sheet1.cell(i + 3, 3).fill = fill1
        sheet1.cell(i + 3, 4).fill = fill1
        # sheet1.cell(i+3, 4, data_list[i]['project'])
    # 写入表尾内容并设置格式
    sheet1.cell(len(data_list) + 4, 1, '总结').fill = fill
    sheet1.merge_cells(start_row=len(data_list) + 4, start_column=2, end_row=len(data_list) + 4, end_column=4)
    sheet1.cell(len(data_list) + 4, 2, '本次总计测试   个patch 测试出了   个问题，今日合入质量: ').fill = fill
    # 保存工作表
    wb.save('./ExcelFile/' + str(generate_time()) + '研发测试.xlsx')
