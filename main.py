import json
import logging

from RD_utils import *

if __name__ == '__main__':
    # LOG格式设置
    LOG_FORMAT = "%(asctime)s - %(levelname)s - %(message)s"
    logging.basicConfig(filename='./log/' + str(generate_time()) + '.log', level=logging.DEBUG, format=LOG_FORMAT)
    logging.info('today is ' + str(generate_time()))
    logging.info("today will select " + generate_select_time(generate_time()))
    # 取cookies
    # get_cookies()
    # 执行主要操作
    w_excel(gen_fix_dict(generate_select_time(generate_time())))

