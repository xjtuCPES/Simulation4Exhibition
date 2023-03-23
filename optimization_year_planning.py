'''
Author: gwyxjtu
Date: 2022-06-06 20:10:39
LastEditors: guo_win 867718012@qq.com
LastEditTime: 2023-03-23 15:12:02
FilePath: /proceeding-case/optimization_year_planning.py
Description: 人一生会遇到约2920万人,两个人相爱的概率是0.000049,所以你不爱我,我不怪你.

Copyright (c) 2022 by gwyxjtu 867718012@qq.com, All Rights Reserved. 
'''
'''
                       ::
                      :;J7, :,                        ::;7:
                      ,ivYi, ,                       ;LLLFS:
                      :iv7Yi                       :7ri;j5PL
                     ,:ivYLvr                    ,ivrrirrY2X,
                     :;r@Wwz.7r:                :ivu@kexianli.
                    :iL7::,:::iiirii:ii;::::,,irvF7rvvLujL7ur
                   ri::,:,::i:iiiiiii:i:irrv177JX7rYXqZEkvv17
                ;i:, , ::::iirrririi:i:::iiir2XXvii;L8OGJr71i
              :,, ,,:   ,::ir@mingyi.irii:i:::j1jri7ZBOS7ivv,
                 ,::,    ::rv77iiiriii:iii:i::,rvLq@huhao.Li
             ,,      ,, ,:ir7ir::,:::i;ir:::i:i::rSGGYri712:
           :::  ,v7r:: ::rrv77:, ,, ,:i7rrii:::::, ir7ri7Lri
          ,     2OBBOi,iiir;r::        ,irriiii::,, ,iv7Luur:
        ,,     i78MBBi,:,:::,:,  :7FSL: ,iriii:::i::,,:rLqXv::
        :      iuMMP: :,:::,:ii;2GY7OBB0viiii:i:iii:i:::iJqL;::
       ,     ::::i   ,,,,, ::LuBBu BBBBBErii:i:i:i:i:i:i:r77ii
      ,       :       , ,,:::rruBZ1MBBqi, :,,,:::,::::::iiriri:
     ,               ,,,,::::i:  @arqiao.       ,:,, ,:::ii;i7:
    :,       rjujLYLi   ,,:::::,:::::::::,,   ,:i,:,,,,,::i:iii
    ::      BBBBBBBBB0,    ,,::: , ,:::::: ,      ,,,, ,,:::::::
    i,  ,  ,8BMMBBBBBBi     ,,:,,     ,,, , ,   , , , :,::ii::i::
    :      iZMOMOMBBM2::::::::::,,,,     ,,,,,,:,,,::::i:irr:i:::,
    i   ,,:;u0MBMOG1L:::i::::::  ,,,::,   ,,, ::::::i:i:iirii:i:i:
    :    ,iuUuuXUkFu7i:iii:i:::, :,:,: ::::::::i:i:::::iirr7iiri::
    :     :rk@Yizero.i:::::, ,:ii:::::::i:::::i::,::::iirrriiiri::,
     :      5BMBBBBBBSr:,::rv2kuii:::iii::,:i:,, , ,,:,:i@petermu.,
          , :r50EZ8MBBBBGOBBBZP7::::i::,:::::,: :,:,::i;rrririiii::
              :jujYY7LS0ujJL7r::,::i::,::::::::::::::iirirrrrrrr:ii:
           ,:  :@kevensun.:,:,,,::::i:i:::::,,::::::iir;ii;7v77;ii;i,
           ,,,     ,,:,::::::i:iiiii:i::::,, ::::iiiir@xingjief.r;7:i,
        , , ,,,:,,::::::::iiiiiiiiii:,:,:::::::::iiir;ri7vL77rrirri::
         :,, , ::::::::i:::i:::i:i::,,,,,:,::i:i:::iir;@Secbone.ii:::
'''

import json
import pprint
import pandas as pd
from cpeslog.log_code import _logging
from Model.optimization_planning import OptimizationPlanning,to_csv,to_csv_2


if __name__ == '__main__':
    _logging.info('start')
    try:
        with open("Config/config.json", "r",encoding='utf-8') as f:
            input_json = json.load(f)
    except BaseException as E:
        _logging.error('读取config失败,错误原因为{}'.format(E))
        raise Exception
    # 读取输入excel
    
    try:
        load = pd.read_excel('Input/load.xlsx')
    except BaseException as E:
        _logging.error('读取input_24h的excel失败,错误原因为{}'.format(E))
        raise Exception
    load = {
        'ele_load':list(load['电负荷kW'].fillna(0)),
        'ele_newload':list(load['电负荷/kwh'].fillna(0)),
        'g_load':list(load['供暖热负荷(kW)'].fillna(0)),
        'water_load':list(load['生活热水负荷kW'].fillna(0)),
        'q_load':list(load['冷负荷(kW)'].fillna(0)),
    }
    
    # 优化主函数
    try:
        dict_plot = OptimizationPlanning(parameter_json=input_json, load_json=load)
    except BaseException as E:
        _logging.error('优化主函数执行失败，错误原因为{}'.format(E))
        raise Exception
    #print(dict_control)
    #print(dict_plot)
    
    #exit(0)
    # 写入输出Excel

    try:
        #to_csv(dict_control,"dict_control_24h")
        to_csv(dict_plot,"dict_plot_planning")
        to_csv_2(dict_plot,"twin_and_screen_sum",4992)
        to_csv_2(dict_plot,"twin_and_screen_win",24)
        to_csv_2(dict_plot,"twin_and_screen_non",2160)
    except BaseException as E:
        _logging.error('excel输出失败,错误原因为{}'.format(E))
        raise Exception