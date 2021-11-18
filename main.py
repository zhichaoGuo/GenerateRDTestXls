import json
import logging

from RD_utils import *

if __name__ == '__main__':
    # LOG格式设置
    LOG_FORMAT = "%(asctime)s - %(levelname)s - %(message)s"
    logging.basicConfig(filename='./log/' + str(generate_time()) + '.log', level=logging.DEBUG, format=LOG_FORMAT)
    logging.info('today is ' + str(generate_time()))
    # 目前暂时这样设计，后面增加开始截至事件时再优化
    select_end_time = generate_time()
    logging.info("today will select " + generate_select_time(generate_time()) + 'to ' + str(select_end_time))
    # 取cookies
    get_cookies()
    # 执行主要操作
    w_excel(gen_fix_dict(generate_select_time(generate_time())))

