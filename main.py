from RD_utils import *


def run_RD_test_xls(start_time=None, end_time=None):
    select_start_time, select_end_time = count_select_time(start_time, end_time)
    cookies = get_cookies_without_selenium()
    header = conf.get('header')
    header.update(cookies)
    new_url = conf.get('root_url') + 'changes/?O=81&S=0&n=50&q=status%3Amerged%20after%3A' + \
              str(select_start_time) + '%20before%3A' + str(select_end_time)
    data = get_data_from_gerrit(new_url, header)
    elpis, hos, other = adjust_original_data(data)
    make_excel(elpis, hos, other)


if __name__ == '__main__':
    # # 执行主要操作
    run_RD_test_xls()
    # run_RD_test_xls('2023-5-15','2023-5-17')
