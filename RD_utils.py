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
    # 打开edge浏览器
    driver = webdriver.Edge()
    logging.info("OPEN Edge")
    # 访问repo登陆界面
    driver.get('http://repo.htek.com:8081/login/')
    logging.info("open http://repo.htek.com:8081/login/")
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
    logging.info("input username "+user_data["username"])
    driver.find_element(By.ID, 'f_pass').send_keys(user_data["password"])
    logging.info("input password len is " + str(len(user_data["password"])))
    driver.find_element(By.ID, 'f_remember').click()
    logging.info("click remember me")
    # 点击登录
    driver.find_element(By.ID, 'b_signin').click()
    logging.info("click sign in")
    driver.get('http://repo.htek.com:8081/changes/')
    logging.info("open http://repo.htek.com:8081/changes/")
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
    temp = []
    result = []
    for i in range(len(list)):
        if list[i][key] not in temp:
            temp.append(list[i][key])
            logging.debug("保存记录 " + key + ':' + list[i][key])
            result.append(list[i])
        else:
            logging.info("移除重复 " + key + ':' + list[i][key])
    return result


def gen_fix_dict(select_time):
    """
    整理修建从json中获取的dict
    :param select_time:
    :return:
    """
    with open("cookies.yaml", encoding="UTF-8") as f:
        yaml_data = yaml.safe_load(f)
        GerritAccount = yaml_data[1]['value']
        XSRF_TOKEN = yaml_data[0]['value']
        cookie = 'jenkins-timestamper-offset=-28800000; GerritAccount=' + GerritAccount + '; XSRF_TOKEN=' + XSRF_TOKEN
        logging.info("use cookies :" + cookie)
        f.close()

    headers = {
        'Host': 'repo.htek.com:8081',
        'Connection': 'keep-alive',
        'Upgrade-Insecure-Requests': '1',

        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) '
                      'Chrome/95.0.4638.54 Safari/537.36 Edg/95.0.1020.40',
        'Accept': '*/*',
        'Accept-Encoding': 'gzip,deflate',
        'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6',
        'Cookie': cookie
    }
    get_url = 'http://repo.htek.com:8081/changes/?O=81&S=0&n=50&q=status%3Amerged%20after%3A' + select_time

    logging.info('get ' + select_time + '数据')
    r = requests.get(get_url, headers=headers)
    logging.info('get request status code :' + str(r.status_code))
    json_text = r.text
    json_text = json_text[4:-1]
    data = json.loads(json_text)

    with open("info_dict.yaml", "w", encoding="UTF-8") as f:
        logging.debug('open info_dict.yaml')
        yaml.dump(data, f, allow_unicode=True)
        f.close()
        logging.debug('close info_dict.yaml')
    with open('info_dict.yaml', 'r+', encoding="UTF-8") as f:
        data_list = yaml.safe_load(f)
        data_json = json.dumps(data_list, ensure_ascii=False)
        f.close()
    for i in range(len(data_list)):

        logging.debug('delete ' + str(data_list[i]['_number']))
        del data_list[i]['_number']

        logging.debug('delete ' + str(data_list[i]['branch']))
        del data_list[i]['branch']

        logging.debug('delete ' + str(data_list[i]['created']))
        del data_list[i]['created']

        logging.debug('delete ' + str(data_list[i]['deletions']))
        del data_list[i]['deletions']

        logging.debug('delete ' + str(data_list[i]['has_review_started']))
        del data_list[i]['has_review_started']

        logging.debug('delete ' + str(data_list[i]['hashtags']))
        del data_list[i]['hashtags']

        logging.debug('delete ' + str(data_list[i]['id']))
        del data_list[i]['id']

        logging.debug('delete ' + str(data_list[i]['insertions']))
        del data_list[i]['insertions']

        logging.debug('delete ' + str(data_list[i]['labels']))
        del data_list[i]['labels']

        logging.debug('delete ' + str(data_list[i]['owner']))
        del data_list[i]['owner']

        logging.debug('delete ' + str(data_list[i]['requirements']))
        del data_list[i]['requirements']

        logging.debug('delete ' + str(data_list[i]['status']))
        del data_list[i]['status']

        logging.debug('delete ' + str(data_list[i]['submitted']))
        del data_list[i]['submitted']

        logging.debug('delete ' + str(data_list[i]['submitter']))
        del data_list[i]['submitter']

        logging.debug('delete ' + str(data_list[i]['total_comment_count']))
        del data_list[i]['total_comment_count']

        logging.debug('delete ' + str(data_list[i]['unresolved_comment_count']))
        del data_list[i]['unresolved_comment_count']

        logging.debug('delete ' + str(data_list[i]['updated']))
        del data_list[i]['updated']
    data_list = remove_same_key_value_list(data_list, 'change_id')
    with open("fix_info_dict.yaml", "w", encoding="UTF-8") as f:
        yaml.dump(data_list, f, allow_unicode=True)
        f.close()
    print(data_list)
    logging.info('data_list is ' + str(data_list))
    return data_list


def generate_time():
    """
    获取当日日期  2021-11-01
    :return:
    """
    today = date.today()
    return today


def generate_select_time(today):
    """
    根据传入日期推算验单日期（如周一则瑞算至周五） 2021-10-29
    :param today:
    :return:
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
    :return:
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
    wb.save(str(generate_time()) + '研发测试.xlsx')
