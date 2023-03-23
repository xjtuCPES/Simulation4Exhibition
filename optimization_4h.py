'''
Author: gwyxjtu
Date: 2022-06-13 10:41:42
LastEditors: gwyxjtu 867718012@qq.com
LastEditTime: 2022-06-15 19:00:55
FilePath: /optimization/optimization_4h.py
Description: 人一生会遇到约2920万人,两个人相爱的概率是0.000049,所以你不爱我,我不怪你.

Copyright (c) 2022 by gwyxjtu 867718012@qq.com, All Rights Reserved. 
'''
'''
                       _oo0oo_
                      o8888888o
                      88" . "88
                      (| -_- |)
                      0\  =  /0
                    ___/`---'\___
                  .' \\|     |// '.
                 / \\|||  :  |||// \
                / _||||| -:- |||||- \
               |   | \\\  - /// |   |
               | \_|  ''\---/''  |_/ |
               \  .-\__  '-'  ___/-. /
             ___'. .'  /--.--\  `. .'___
          ."" '<  `.___\_<|>_/___.' >' "".
         | | :  `- \`.;`\ _ /`;.`/ - ` : | |
         \  \ `_.   \_ __\ /__ _/   .-` /  /
     =====`-.____`.___ \_____/___.-`___.-'=====
                       `=---='


     ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

           佛祖保佑     永不宕机     永无BUG
'''



import json
import pprint
import pandas as pd
from cpeslog.log_code import _logging
from Model.optimization_day import OptimizationDay,to_csv

if __name__ == '__main__':
    _logging.info('start')
    try:
        with open("Config/config.json", "r") as f:
            input_json = json.load(f)
    except BaseException as E:
        _logging.error('读取config失败,错误原因为{}'.format(E))
        raise Exception
    # 读取输入excel
    try:
        load = pd.read_excel('Input/input_4h.xls')
    except BaseException as E:
        _logging.error('读取input_4h的excel失败,错误原因为{}'.format(E))
        raise Exception

    try:
        sto = pd.read_excel('Input/input_now.xls')
    except BaseException as E:
        _logging.error('读取input_now的excel失败,错误原因为{}'.format(E))
        raise Exception
    # 确定终止时刻设备容量
    sto_4 = sto
    if sto['time'][0] >= 20:
        sto_4['end_slack'] = True
    else:
        sto_4['end_slack'] = False

    # 执行优化主程序
    try:
        dict_control,dict_plot = OptimizationDay(parameter_json=input_json, load_json=load, begin_time = sto['time'][0], time_scale=4, storage_begin_json=sto, storage_end_json=sto_4)
    except BaseException as E:
        _logging.error('优化住函数执行失败，错误原因为{}'.format(E))
        raise Exception

    # 写入输出到excel
    try:
        to_csv(dict_control,"dict_control_4h")
        to_csv(dict_plot,"dict_plot_4h")
    except BaseException as E:
        _logging.error('excel输出失败,错误原因为{}'.format(E))
        raise Exception
