import json
import logging

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


if __name__ == '__main__':
    # # LOG格式设置
    # LOG_FORMAT = "%(asctime)s - %(levelname)s - %(message)s"
    # logging.basicConfig(filename='./log/' + str(generate_time()) + '.log', level=logging.DEBUG, format=LOG_FORMAT)
    # logging.info('today is ' + str(generate_time()))
    # # 目前暂时这样设计，后面增加开始截至事件时再优化
    # select_end_time = generate_time()
    # logging.info("today will select " + generate_select_time(generate_time()) + 'to ' + str(select_end_time))
    # # 取cookies
    get_cookies()
    # # 执行主要操作
    # w_excel(gen_fix_dict(generate_select_time(generate_time())))
    run_RD_test_xls(log_level='DEBUG')
