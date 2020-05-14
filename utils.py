import xlrd
import os
import traceback

def check_old_info(old_info):
    """
    判断old_info中[度,分,秒]的异常数据，比如空值 非数字
    return: bool 是否异常，list 异常数字的index位置
    ATTENTION:  默认补0
    """
    err_list = []    # 异常的数字的位置
    for i in [2,3,4,6,7,8]:
        # 保证old_info中这六个地方是数字
        if type(old_info[i]) in [int, float]:
            continue
        # 默认补零
        old_info[i] = 0
        err_list.append(i+1)
    return len(err_list) != 0, err_list