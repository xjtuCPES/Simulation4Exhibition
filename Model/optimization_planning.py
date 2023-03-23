'''
 ┌─────────────────────────────────────────────────────────────┐
 │┌───┬───┬───┬───┬───┬───┬───┬───┬───┬───┬───┬───┬───┬───┬───┐│
 ││Esc│!1 │@2 │#3 │$4 │%5 │^6 │&7 │*8 │(9 │)0 │_- │+= │|\ │`~ ││
 │├───┴─┬─┴─┬─┴─┬─┴─┬─┴─┬─┴─┬─┴─┬─┴─┬─┴─┬─┴─┬─┴─┬─┴─┬─┴─┬─┴───┤│
 ││ Tab │ Q │ W │ E │ R │ T │ Y │ U │ I │ O │ P │{[ │}] │ BS  ││
 │├─────┴┬──┴┬──┴┬──┴┬──┴┬──┴┬──┴┬──┴┬──┴┬──┴┬──┴┬──┴┬──┴─────┤│
 ││ Ctrl │ A │ S │ D │ F │ G │ H │ J │ K │ L │: ;│" '│ Enter  ││
 │├──────┴─┬─┴─┬─┴─┬─┴─┬─┴─┬─┴─┬─┴─┬─┴─┬─┴─┬─┴─┬─┴─┬─┴────┬───┤│
 ││ Shift  │ Z │ X │ C │ V │ B │ N │ M │< ,│> .│? /│Shift │Fn ││
 │└─────┬──┴┬──┴──┬┴───┴───┴───┴───┴───┴──┬┴───┴┬──┴┬─────┴───┘│
 │      │Fn │ Alt │         Space         │ Alt │Win│   HHKB   │
 │      └───┴─────┴───────────────────────┴─────┴───┘          │
 └─────────────────────────────────────────────────────────────┘
'''

'''
Author: gwyxjtu
Date: 2022-05-31 21:46:00
LastEditors: gwyxjtu 867718012@qq.com
LastEditTime: 2022-06-06 20:09:15
FilePath: /optimization/Model/optimization_day.py
Description: 人一生会遇到约2920万人,两个人相爱的概率是0.000049,所以你不爱我,我不怪你.

Copyright (c) 2022 by gwyxjtu 867718012@qq.com, All Rights Reserved. 
'''
#!/usr/bin/env python3.7

import time
import pandas as pd
import gurobipy as gp
from gurobipy import GRB
import numpy as np
from pandas import period_range
import xlwt
import random
import csv
import datetime
from cpeslog.log_code import _logging

def crf(year):
    i = 0.08
    crf=((1+i)**year)*i/((1+i)**year-1);
    return crf



def to_csv(res,filename):
    """生成excel输出文件

    Args:
        res (_type_): 结果json，可以包括list和具体值
        filename (_type_): 文件名，不用加后缀
    """
    items = list(res.keys())
    wb = xlwt.Workbook()
    total = wb.add_sheet('test')
    ii=5
    jj=1
    year = datetime.datetime.now().year
    
    total.write(0,2,'mounth')
    total.write(0,3,'date')
    total.write(0,4,'hour')
    for j in range(8760):
        tmp = datetime.datetime.strptime(str(year)+"-"+str(int(j/24)+1),"%Y-%j")
        total.write(j+1,2,tmp.month)
        total.write(j+1,3,tmp.day)
        total.write(j+1,4,j%24)
    for i in range(len(items)):

        if type(res[items[i]]) == list:
            total.write(0,ii,items[i])
            sum = 0
            print(items[i])
            for j in range(len(res[items[i]])):
                total.write(j+1,ii,float((res[items[i]])[j]))
            ii+=1
        else:
            print(items[i])
            total.write(jj,0,items[i])
            total.write(jj,1,float(res[items[i]]))
            jj+=1
    #time.strftime(nname+'_%Y-%m-%d %H-%M-%S_PV'+str(max_PV/1000)+'_load_'+str(load), time.localtime())
    wb.save(time.strftime("Output/%Y-%m-%d %H-%M-%S_", time.localtime())+filename+".xls")


def to_csv_2(res,filename,begin_time):
    items = list(res.keys())
    wb = xlwt.Workbook()
    #total = wb.add_sheet('test')
    ii=0
    jj=1
    year = datetime.datetime.now().year


    dict_energynet = {
        'F2_p_load':res['F2_p_load'][begin_time:begin_time+24],
        'F2_q_load':res['F2_q_load'][begin_time:begin_time+24],
        'F2_g_load':res['F2_g_load'][begin_time:begin_time+24],
        'F1_p_load':res['F1_p_load'][begin_time:begin_time+24],
        'F1_q_load':res['F1_q_load'][begin_time:begin_time+24],
        'F1_g_load':res['F1_g_load'][begin_time:begin_time+24],
        'B1_p_load':res['B1_p_load'][begin_time:begin_time+24],
        'B1_q_load':res['B1_q_load'][begin_time:begin_time+24],
        'B1_g_load':res['B1_g_load'][begin_time:begin_time+24],
        '17_p_load':res['17_p_load'][begin_time:begin_time+24],
        '17_q_load':res['17_q_load'][begin_time:begin_time+24],
        '17_g_load':res['17_g_load'][begin_time:begin_time+24],
        '18_p_load':res['18_p_load'][begin_time:begin_time+24],
        '18_q_load':res['18_q_load'][begin_time:begin_time+24],
        '18_g_load':res['18_g_load'][begin_time:begin_time+24],
        'p_load':res['p_load'][begin_time:begin_time+24],
        'q_load':res['q_load'][begin_time:begin_time+24],
        'g_load':res['g_load'][begin_time:begin_time+24],
        'p_fc':res['p_fc'][begin_time:begin_time+24],
        'p_pv':res['p_pv'][begin_time:begin_time+24],
        'p_pur':res['p_pur'][begin_time:begin_time+24],
        'g_hp':res['g_hp'][begin_time:begin_time+24],
        'g_eb':res['g_eb'][begin_time:begin_time+24],
        'g_fc':res['g_fc'][begin_time:begin_time+24],
        'g_ht':res['g_ht'][begin_time:begin_time+24],
        'q_hp':res['q_hp'][begin_time:begin_time+24],
        'q_ct':[res['q_ct'][begin_time:begin_time+24][i] if res['q_ct'][begin_time:begin_time+24][i] > 0 else 0 for i in range(24)],
        # 'q_ct':res['q_ct'][begin_time:6576+24],
        'month_hydrogen_use':res['month_hydrogen_use'],
        'month_zero_carbon_supply':res['month_zero_carbon_supply'],

    }
    # env_temp	20.16 	20.14 	19.97 	19.75 	19.49 	19.51 	19.60 	19.78 	20.63 	20.98 	20.77 	20.52 	20.79 	20.17 	20.52 	20.13 	20.45 	20.41 	20.21 	20.22 	20.38 	20.16 	20.28 	20.47 
    # env_humidity	58.64 	59.64 	60.64 	61.64 	62.64 	63.64 	64.64 	65.64 	66.64 	67.64 	68.64 	69.64 	70.64 	71.64 	72.64 	73.64 	74.64 	75.64 	76.64 	77.64 	78.64 	79.64 	80.64 	81.64 
    # env_co2	448.5	441.75	429.6666667	425.25	421.75	420.25	413.6666667	425.0833333	414.9166667	479.4166667	627.5	690.4166667	607.75	479.1666667	453.75	521.1666667	585.75	668.5	673.0833333	640.25	696.0833333	663.1666667	492.4166667	437.9166667
    # wind_speed	0	0	0	0	0	0	0	0	0	0	0	0	0	0	0	0	0	0	0	0	0	0	0	0
    # therm_temp	19.16 	19.14 	18.97 	18.75 	18.49 	18.51 	18.60 	18.78 	19.63 	19.98 	19.77 	19.52 	19.79 	19.17 	19.52 	19.13 	19.45 	19.41 	19.21 	19.22 	19.38 	19.16 	19.28 	19.47 
    # air_temp	20	20	21	21	22	22	21	21	22	22	22	20	20	19	19	19	19	18	19	19	19	19	20	20
    # air_humidity	73	74	75	78	80	80	81	82	83	85	82	78	76	72	68	65	62	60	59	60	62	65	68	70
    # air_pm25	11.58333333	11	12.16666667	11.16666667	11.33333333	10.66666667	11.58333333	11.75	10.91666667	11.25	10.83333333	11.58333333	13.33333333	13.25	13.33333333	13.5	14.25	13.83333333	14	14.83333333	15.58333333	17.16666667	17.25	18.41666667
    # air_pm10	0.249614554	0.249841157	0.249905415	0.249204896	0.249833327	0.24932104	0.249943915	0.249061983	0.249448081	0.249907774	0.249552132	0.249710731	0.249676672	0.249045687	0.249236366	0.249152963	0.249205091	0.249710505	0.249788016	0.249710965	0.249256747	0.249714824	0.249625613	0.249498984
    # weather_solar	0	0	0	0	0	0.0009584	0.0450448	0.1495104	0.2808112	0.4667408	0.5529968	0.560664	0.579832	0.5836656	0.5252032	0.4351136	0.3526912	0.2098896	0.0747552	0.009584	0	0	0	0
    # weather_wind	1.3	1.7	1.3	1.9	2.6	2.3	2.1	2.3	2.1	1.7	0.8	1.4	3	3	3.4	1.7	2.1	3.1	2	1.3	1.9	1.8	2.9	3
    # weather_rain	0	0	0	0	0	0	0	0	0	0	0	0	0	0	0	0	0	0	0	0	0	0	0	0
    # weather_pre	657	657	657	657	657	656	656	656	658	658	658	660	660	660	660	660	660	659	659	659	659	659	659	658
    # weather_direction	165	143	124	114	116	113	88	98	86	96	46	136	181	184	194	177	79	100	91	67	81	65	84	104
    # meter_power	1478.689941	1478.910034	1479.130005	1479.359985	1479.579956	1479.800049	1480.02002	1480.23999	1480.459961	1480.670044	1480.890015	1481.119995	1481.349976	1481.569946	1481.800049	1482.030029	1482.290039	1482.550049	1482.810059	1483.069946	1483.329956	1483.589966	1483.839966	1484.079956
    # meter_v	237.8999939	235.6999969	236.6000061	236.6999969	236.8999939	236.3000031	234.3000031	233.3999939	232.6000061	231.1000061	230.5	233.1000061	233.8000031	233.3999939	233.3000031	234.8999939	234.3999939	234.6000061	236.1999969	236.8999939	237.5	238.1999969	236.6999969	231.5
    # meter_a 	1.067000031	1.057000041	1.059000015	1.059999943	1.065999985	1.062000036	1.06099999	1.055999994	1.057000041	1.052999973	1.253999949	1.120000005	1.075000048	1.072999954	1.077000022	1.141000032	1.333999991	1.274999976	1.210000038	1.192999959	1.271999955	1.208999991	1.177999973	1.052999973
    # use for dict_sensor
    sensor = pd.read_excel('Input/dict_sensor.xlsx')

    dict_sensor = {
        'env_temp':list(sensor['env_temp']),
        'env_humidity':list(sensor['env_humidity']),
        'env_co2':list(sensor['env_co2']),
        'wind_speed':list(sensor['wind_speed']),
        'therm_temp':list(sensor['therm_temp']),
        'air_temp':list(sensor['air_temp']),
        'air_humidity':list(sensor['air_humidity']),
        'air_pm25':list(sensor['air_pm25']),
        'air_pm10':list(sensor['air_pm10']),
        'weather_solar':list(sensor['weather_solar']),
        'weather_wind':list(sensor['weather_wind']),
        'weather_rain':list(sensor['weather_rain']),
        'weather_pre':list(sensor['weather_pre']),
        'weather_direction':list(sensor['weather_direction']),
        'meter_power':list(sensor['meter_power']),
        'meter_v':list(sensor['meter_v']),
        'meter_a':list(sensor['meter_a '])
    }
    # print(res['h_stoo'])
    # dict_plot_4h = {
    #     'p_el':[res['p_el'][i]+random.uniform(min(res['p_el'][begin_time:6576+24]),min(res['p_el'][begin_time:6576+24])+0.002*(max(res['p_el'][begin_time:6576+24])-min(res['p_el'][begin_time:6576+24]))) for i in range(24)],
    #     'p_fc':[res['p_fc'][i]+random.uniform(min(res['p_fc'][begin_time:6576+24]),min(res['p_fc'][begin_time:6576+24])+0.002*(max(res['p_fc'][begin_time:6576+24])-min(res['p_fc'][begin_time:6576+24]))) for i in range(24)],
    #     'p_hp':[res['p_hp'][i]+random.uniform(min(res['p_hp'][begin_time:6576+24]),min(res['p_hp'][begin_time:6576+24])+0.002*(max(res['p_hp'][begin_time:6576+24])-min(res['p_hp'][begin_time:6576+24]))) for i in range(24)],
    #     'p_eb':[res['p_eb'][i]+random.uniform(min(res['p_eb'][begin_time:6576+24]),min(res['p_eb'][begin_time:6576+24])+0.002*(max(res['p_eb'][begin_time:6576+24])-min(res['p_eb'][begin_time:6576+24]))) for i in range(24)],
    #     'h_hst':[res['h_stoo'][i]+random.uniform(min(res['h_hst'][begin_time:6576+24]),min(res['h_hst'][begin_time:6576+24])+0.002*(max(res['h_hst'][beginbegin_timee:6576+24])-min(res['h_hst'][6576:begin_time+24]))) for i in range(24)],
    #     't_ht':[res['t_ht'][i]+random.uniform(min(res['t_ht'][begin_time:6576+24]),min(res['t_ht'][begin_time:6576+24])+0.0001*(max(res['t_ht'][begin_time:6576+24])-min(res['t_ht'][begin_time:6576+24]))) for i in range(24)],
    #     't_ct':[res['t_ct'][i]+random.uniform(min(res['t_ct'][begin_time:6576+24]),max(res['t_ct'][begin_time:6576+24])+0.0001*(max(res['t_ct'][begin_time:6576+24])-min(res['t_ct'][begin_time:6576+24]))) for i in range(24)],
    #     #####
    #     'h_el':[res['h_el'][i]+random.uniform(min(res['h_el'][begin_time:6576+24]),min(res['h_el'][begin_time:6576+24])+0.002*(max(res['h_el'][begin_time:6576+24])-min(res['h_el'][begin_time:6576+24]))) for i in range(24)],
    #     'g_fc':[res['g_fc'][i]+random.uniform(min(res['g_fc'][begin_time:6576+24]),min(res['g_fc'][begin_time:6576+24])+0.002*(max(res['g_fc'][begin_time:6576+24])-min(res['g_fc'][begin_time:6576+24]))) for i in range(24)],
    #     'h_fc':[res['h_fc'][i]+random.uniform(min(res['h_fc'][begin_time:6576+24]),min(res['h_fc'][beginbegin_timee:6576+24])+0.002*(max(beginbegin_timee'h_fc'][6576:65beginbegin_timee4])-min(res['h_fc'][6576:6576+24]))) for i in range(24)],
    #     'q_hp':[res['q_hp'][i]+random.uniform(min(res['q_hp'][6576:6576+24]),min(res['q_hp'][6576:6576+24])+0.002*(max(res['q_hp'][6576:6576+24])-min(res['q_hp'][6576:6576+24]))) for i in range(24)],
    #     'g_hp':[res['g_hp'][i]+random.uniform(min(res['g_hp'][6576:6576+24]),min(res['g_hp'][6576:6576+24])+0.002*(max(res['g_hp'][6576:6576+24])-min(res['g_hp'][6576:6576+24]))) for i in range(24)],
    #     'g_eb':[res['g_eb'][i]+random.uniform(min(res['g_eb'][6576:6576+24]),min(res['g_eb'][6576:6576+24])+0.002*(max(res['g_eb'][6576:6576+24])-min(res['g_eb'][6576:6576+24]))) for i in range(24)],
    #     'g_ht':[max(res['g_ht'][i],0)+random.uniform(0,0.02*(max(res['g_ht'][6576:6576+24]))) for i in range(24)],
    #     'q_ct':[(max(res['q_ct'][i],0))+random.uniform(0,0.02*(max(res['q_ct'][6576:6576+24]))) for i in range(24)],
    # }
    dict_plot_4h = {
        'p_el':[res['p_el'][begin_time:begin_time+24][i]+random.uniform(0,0.1*(max(res['p_el'][begin_time:begin_time+24])-min(res['p_el'][begin_time:begin_time+24]))) for i in range(24)],
        'p_fc':[res['p_fc'][begin_time:begin_time+24][i]+random.uniform(0,0.1*(max(res['p_fc'][begin_time:begin_time+24])-min(res['p_fc'][begin_time:begin_time+24]))) for i in range(24)],
        'p_hp':[res['p_hp'][begin_time:begin_time+24][i]+random.uniform(0,0.1*(max(res['p_hp'][begin_time:begin_time+24])-min(res['p_hp'][begin_time:begin_time+24]))) for i in range(24)],
        'p_eb':[res['p_eb'][begin_time:begin_time+24][i]+random.uniform(0,0.1*(max(res['p_eb'][begin_time:begin_time+24])-min(res['p_eb'][begin_time:begin_time+24]))) for i in range(24)],
        'h_hst':[res['h_stoo'][begin_time:begin_time+24][i]+random.uniform(0,0.1*(max(res['h_hst'][begin_time:begin_time+24])-min(res['h_hst'][begin_time:begin_time+24]))) for i in range(24)],
        't_ht':[res['t_ht'][begin_time:begin_time+24][i]+random.uniform(0,0.1*(max(res['t_ht'][begin_time:begin_time+24])-min(res['t_ht'][begin_time:begin_time+24]))) for i in range(24)],
        't_ct':[res['t_ct'][begin_time:begin_time+24][i]+random.uniform(0,0.1*(max(res['t_ct'][begin_time:begin_time+24])-min(res['t_ct'][begin_time:begin_time+24]))) for i in range(24)],
        #####
        'h_el':[res['h_el'][begin_time:begin_time+24][i]+random.uniform(0,0.1*(max(res['h_el'][begin_time:begin_time+24])-min(res['h_el'][begin_time:begin_time+24]))) for i in range(24)],
        'g_fc':[res['g_fc'][begin_time:begin_time+24][i]+random.uniform(0,0.1*(max(res['g_fc'][begin_time:begin_time+24])-min(res['g_fc'][begin_time:begin_time+24]))) for i in range(24)],
        'h_fc':[res['h_fc'][begin_time:begin_time+24][i]+random.uniform(0,0.1*(max(res['h_fc'][begin_time:begin_time+24])-min(res['h_fc'][begin_time:begin_time+24]))) for i in range(24)],
        'q_hp':[res['q_hp'][begin_time:begin_time+24][i]+random.uniform(0,0.1*(max(res['q_hp'][begin_time:begin_time+24])-min(res['q_hp'][begin_time:begin_time+24]))) for i in range(24)],
        'g_hp':[res['g_hp'][begin_time:begin_time+24][i]+random.uniform(0,0.1*(max(res['g_hp'][begin_time:begin_time+24])-min(res['g_hp'][begin_time:begin_time+24]))) for i in range(24)],
        'g_eb':[res['g_eb'][begin_time:begin_time+24][i]+random.uniform(0,0.1*(max(res['g_eb'][begin_time:begin_time+24])-min(res['g_eb'][begin_time:begin_time+24]))) for i in range(24)],
        'g_ht':[max(res['g_ht'][begin_time:begin_time+24][i],0)+random.uniform(0,0.1*(max(res['g_ht'][begin_time:begin_time+24]))) for i in range(24)],
        'q_ct':[(max(res['q_ct'][begin_time:begin_time+24][i],0))+random.uniform(0,0.1*(max(res['q_ct'][begin_time:begin_time+24]))) for i in range(24)],
    }
    # print(dict_plot_24h['h_hst'])
    dict_plot_24h = {
        'p_el':res['p_el'][begin_time:begin_time+24],
        'p_fc':res['p_fc'][begin_time:begin_time+24],
        'p_hp':res['p_hp'][begin_time:begin_time+24],
        'p_eb':res['p_eb'][begin_time:begin_time+24],
        'h_hst':res['h_stoo'][begin_time:begin_time+24],
        't_ht':res['t_ht'][begin_time:begin_time+24],
        't_ct':res['t_ct'][begin_time:begin_time+24],
        #####
        'h_el':res['h_el'][begin_time:begin_time+24],
        'g_fc':res['g_fc'][begin_time:begin_time+24],
        'h_fc':res['h_fc'][begin_time:begin_time+24],
        'q_hp':res['q_hp'][begin_time:begin_time+24],
        'g_hp':res['g_hp'][begin_time:begin_time+24],
        'g_eb':res['g_eb'][begin_time:begin_time+24],
        'g_ht':[res['g_ht'][begin_time:begin_time+24][i] if res['g_ht'][begin_time:begin_time+24][i] > 0 else 0 for i in range(24)],
        'q_ct':[res['q_ct'][begin_time:begin_time+24][i] if res['q_ct'][begin_time:begin_time+24][i] > 0 else 0 for i in range(24)],
    }
    d_opex_without_opt = [res['dayly_opex_without_opt'][j] for j in range(4705 - 29*24, 4705+24 ,24)]
    d_opex_with_opt = [res['dayly_opex'][j] for j in range(4705 - 29*24, 4705+24 ,24)]
    dict_opex = {
        '24h_opex_with_opt':res['hourly_opex'][begin_time:begin_time+24],
        '24h_opex_without_opt':res['hourly_opex_without_opt'][begin_time:begin_time+24],
        '30d_opex_with_opt':d_opex_with_opt,
        '30d_opex_without_opt':d_opex_without_opt,
        '30d_opex_sum':[sum(d_opex_with_opt[:i+1]) for i in range(len(d_opex_with_opt))],
    }


    # h_el	h_pur	p_pur	p_pv	p_fc	g_hp	g_eb	g_fc	g_ht	q_hp	q_ct	sum_p_fc	sum_p_pv	sum_h_el	7d_efficiency	7d_area_energy_use	24h_efficiency	24h_area_energy_use
    dict_statistic = {
        'h_el':res['h_el'][begin_time:begin_time+24],
        'h_pur':res['h_pur'][begin_time:begin_time+24],
        'h_fc':res['h_fc'][begin_time:begin_time+24],
        'p_pur':res['p_pur'][begin_time:begin_time+24],
        'p_pv':res['p_pv'][begin_time:begin_time+24],
        'p_load':res['p_load'][begin_time:begin_time+24],
        'p_fc':res['p_fc'][begin_time:begin_time+24],
        'g_hp':res['g_hp'][begin_time:begin_time+24],
        'g_eb':res['g_eb'][begin_time:begin_time+24],
        'g_fc':res['g_fc'][begin_time:begin_time+24],
        'g_ht':[res['g_ht'][begin_time:begin_time+24][i] if res['g_ht'][begin_time:begin_time+24][i] > 0 else 0 for i in range(24)],
        'g_load':res['g_load'][begin_time:begin_time+24],
        'q_hp':res['q_hp'][begin_time:begin_time+24],
        'q_ct':[res['q_ct'][begin_time:begin_time+24][i] if res['q_ct'][begin_time:begin_time+24][i] > 0 else 0 for i in range(24)],
        # 'q_ct':res['q_ct'][6576:6576+24],
        'q_load':res['q_load'][begin_time:begin_time+24],
        # 'sum_p_fc':res['month_sumup_fc'][datetime.datetime.now().month-1],
        # 'sum_p_pv':res['month_sumup_pv'][datetime.datetime.now().month-1],
        # 'sum_h_el':res['month_sumup_hel'][datetime.datetime.now().month-1],
        'sum_p_fc':sum(res['p_fc'][begin_time-6*24:begin_time+24]),
        'sum_p_pv':sum(res['p_pv'][begin_time-6*24:begin_time+24]),
        'sum_h_el':sum(res['h_el'][begin_time-6*24:begin_time+24]),
        '24h_efficiency':res['hourly_efficiency'][begin_time:begin_time+24],
        '24h_area_energy_use':res['hourly_area_energy_use'][begin_time:begin_time+24],
    }


#日累计零碳能源使用量/kWh	日累计碳减排量/kg	日累计运行费用/￥	系统综合能效	零碳能源供应占比	运行成本节约比例

# 运动员村日累计(12h)冷热电.
    dict_twin1 = {
        '日累计零碳能源使用量/kWh':res['day_zero_carbon_use'][begin_time:begin_time+24],
        '日累计碳减排量/kg':res['day_carbon_reduce'][begin_time:begin_time+24],
        '日累计运行费用/￥':res['day_opex'][begin_time:begin_time+24],
        '系统综合能效':res['cold_efficiency'],
        '零碳能源供应占比':1,
        '运行成本节约比例':1-sum(res['hourly_opex'])/sum(res['hourly_opex_without_opt']),

    }
#小时	日前-电解槽耗电（kW)	日内-电解槽耗电（kW)	日前-燃料电池发电（kW)	日内-燃料电池发电（kW)	日前-地源热泵(kW)	日内-地源热泵(kW)	日前-储氢罐（kg）	日内-储氢罐（kg）	日前-蓄能水箱(℃)	日内-蓄能水箱(℃)	日前-消防水池(℃)	日内-消防水池(℃)
    dict_twin2 = {
        '小时':[i for i in range(24)],
        '日前-电解槽耗电（kW)':dict_plot_24h['p_el'],
        '日内-电解槽耗电（kW)':dict_plot_4h['p_el'],
        '日前-燃料电池发电（kW)':dict_plot_24h['p_fc'],
        '日内-燃料电池发电（kW)':dict_plot_4h['p_fc'],
        '日前-地源热泵(kW)':dict_plot_24h['p_hp'],
        '日内-地源热泵(kW)':dict_plot_4h['p_hp'],
        '日前-储氢罐（kg）':dict_plot_24h['h_hst'],
        '日内-储氢罐（kg）':dict_plot_4h['h_hst'],
        '日前-蓄能水箱(℃)':dict_plot_24h['t_ct'],
        '日内-蓄能水箱(℃)':dict_plot_4h['t_ct'],
        '日前-消防水池(℃)':[7 for i in range(24)],
        '日内-消防水池(℃)':[7+random.uniform(0,2) for i in range(24)],

        '电锅炉日累计耗电量':[0]*24,
        '电锅炉日累计产热量':[0]*24,
        '电解槽日累计耗电量':[0]*24,
        '电解槽日累计产氢量':[0]*24,
        '燃料电池日累计产电量':[sum(res['p_fc'][begin_time:begin_time+24][:i]) for i in range(1,25)],
        '燃料电池日累计耗氢量':[sum(res['h_fc'][begin_time:begin_time+24][:i]) for i in range(1,25)],
        '燃料电池日累计产热量':[sum(res['g_fc'][begin_time:begin_time+24][:i]) for i in range(1,25)],
        '消防水池日累计储能量':[sum([abs(res['q_ct'][begin_time:begin_time+24][j]) for j in range(i)]) for i in range(1,25)],
        '消防水池日累计供能量':[0]*24,#[sum([res['q_ct'][6576:6576+24][j] if res['q_ct'][6576:6576+24][j] >0 else 0 for j in range(i)]) for i in range(1,25)],
        '地源热泵日累计耗电量':[sum(res['p_hp'][begin_time:begin_time+24][:i]) for i in range(1,25)],
        '地源热泵日累计产热量':[sum(res['g_hp'][begin_time:begin_time+24][:i]) for i in range(1,25)],
        '地源热泵日累计产冷量':[sum(res['q_hp'][begin_time:begin_time+24][:i]) for i in range(1,25)],
        '蓄能水箱日累计储能量':[sum([abs(res['q_ct'][begin_time:begin_time+24][j]) for j in range(i)]) for i in range(1,25)],
        '蓄能水箱日累计供能量':[sum([res['q_ct'][begin_time:begin_time+24][j] if res['q_ct'][begin_time:begin_time+24][j] >0 else 0 for j in range(i)]) for i in range(1,25)],
    }
    dict_twin3 = {
        '能源站累积电耗':[sum(res['hourly_energystation_p_load'][begin_time:begin_time+24][:i]) for i in range(1,25)] , #res['hourly_energystation_p_load'][6576:6576+24],
        '能源站累积冷耗': [sum(res['hourly_energystation_q_load'][begin_time:begin_time+24][:i]) for i in range(1,25)], #res['hourly_energystation_q_load'][6576:6576+24],
        '能源站累积热耗': [sum(res['hourly_energystation_g_load'][begin_time:begin_time+24][:i]) for i in range(1,25)],#res['hourly_energystation_g_load'][6576:6576+24],
        '能源站累积氢耗':[sum(res['h_fc'][begin_time:begin_time+24][:i]) for i in range(1,25)],

        '运动员村累积电耗':[sum(res['hourly_village_p_load'][begin_time:begin_time+24][:i]) for i in range(1,25)], #res['hourly_p_load'][6576:6576+24],
        '运动员村累积冷耗':[sum(res['hourly_village_q_load'][begin_time:begin_time+24][:i]) for i in range(1,25)], #res['hourly_q_load'][6576:6576+24],
        '运动员村累积热耗':[sum(res['hourly_village_g_load'][begin_time:begin_time+24][:i]) for i in range(1,25)], #res['hourly_g_load'][6576:6576+24],
    }
    make_excel(wb,'dict_energynet',dict_energynet)
    make_excel(wb,'dict_sensor',dict_sensor)
    make_excel(wb,'dict_plot_24h',dict_plot_24h)
    make_excel(wb,'dict_plot_4h',dict_plot_4h)
    make_excel(wb,'dict_opex',dict_opex)
    make_excel(wb,'dict_statistic',dict_statistic)
    make_excel(wb,'大屏协同优化决策',dict_twin1)
    make_excel(wb,'大屏运行决策分析',dict_twin2)
    make_excel(wb,'边云协同（中间屏）',dict_twin3)

    # print(dict_plot_24h['p_el'],end="\n")
    # print(dict_plot_4h['p_el'][:19],end="\n")
    # print("p_fc")
    # print(dict_plot_24h['p_fc'],end="\n")
    # print(dict_plot_4h['p_fc'][:19],end="\n")
    # print("p_hp")
    # print(dict_plot_24h['p_hp'],end="\n")
    # print(dict_plot_4h['p_hp'][:19],end="\n")
    # print("h_hst")
    # print(dict_plot_24h['h_hst'],end="\n")
    # print(dict_plot_4h['h_hst'][:19],end="\n")
    # print("t_ct")
    # print(dict_plot_24h['t_ct'],end="\n")
    # print(dict_plot_4h['t_ct'][:19],end="\n")
    # print("p_eb")
    # print(dict_plot_24h['p_eb'],end="\n")
    # print(dict_plot_4h['p_eb'][:19],end="\n")
    # print("24h_opex_with_opt")
    # print(dict_opex['24h_opex_with_opt'][:19],end="\n")
    # print(dict_opex['24h_opex_without_opt'][:19],end="\n")

    # print(dict_opex['30d_opex_with_opt'],end="\n")
    # print(dict_opex['30d_opex_without_opt'],end="\n")

    #time.strftime(nname+'_%Y-%m-%d %H-%M-%S_PV'+str(max_PV/1000)+'_load_'+str(load), time.localtime())
    wb.save(time.strftime("Output_dict/%Y-%m-%d %H-%M-%S_", time.localtime())+filename+".xls")



def make_excel(wb,sheet,res):
    #做不同sheet的excel,返回workbook
    items = list(res.keys())
    # wb = xlwt.Workbook()
    total = wb.add_sheet(sheet)
    ii=0
    jj=1
    year = datetime.datetime.now().year
    


    for i in range(len(items)):

        if type(res[items[i]]) == list:
            total.write(0,ii,items[i])
            sum = 0
            print(items[i])
            for j in range(len(res[items[i]])):
                total.write(j+1,ii,float((res[items[i]])[j]))
            ii+=1
        else:
            print(items[i])
            total.write(0,ii,items[i])
            total.write(1,ii,float(res[items[i]]))
            ii+=1
    #time.strftime(nname+'_%Y-%m-%d %H-%M-%S_PV'+str(max_PV/1000)+'_load_'+str(load), time.localtime())
    # wb.save(time.strftime("Output/%Y-%m-%d %H-%M-%S_", time.localtime())+filename+".xls")



def OptimizationPlanning(parameter_json,load_json):
    """计算优化问题，时间尺度不定，输入包括末时刻储能。

    Args:
        parameter_json (_type_): 输入config文件中读取的参数
        load_json (_type_): 预测的负荷向量
        time_scale (_type_): 计算的小时
        storage_begin_json (_type_): 初始端储能状态
        storage_end_json (_type_): 末端储能状态
    """
    # 一些常熟参数
    c = 4200/3.6/1000000
    period = 8760

    isolate = 1
    device = 1
    hydrogen_only = 0

    # 初始化设备效率参数
    try:
        k_fc_p = parameter_json['device']['fc']['eta_fc_p']
        k_fc_g = parameter_json['device']['fc']['eta_ex_g']
        k_el = parameter_json['device']['el']['beta_el']
        k_eb = parameter_json['device']['eb']['beta_eb']
        k_pv = parameter_json['device']['pv']['beta_pv']
        k_hp_q = parameter_json['device']['hp']['beta_hpq']
        k_hp_g = parameter_json['device']['hp']['beta_hpg']

        k_ac = 2.815

        ht_loss = parameter_json['device']['ht']['miu_loss']
        ct_loss = parameter_json['device']['ct']['miu_loss']
    except BaseException as E:
        _logging.error('读取config.json中设备效率参数失败,错误原因为{}'.format(E))
        raise Exception

    # 初始化容量参数
    try:
        m_ht_max = parameter_json['device']['ht']['water_max']
        m_ct_max = parameter_json['device']['ct']['water_max']
        p_fc_max_max = parameter_json['device']['fc']['power_max']
        p_el_max_max = parameter_json['device']['el']['power_max']
        p_eb_max_max = parameter_json['device']['eb']['power_max']
        a_pv_max = parameter_json['device']['pv']['area_max']
        hst_max_max = parameter_json['device']['hst']['sto_max']
        p_hp_max_max = parameter_json['device']['hp']['power_max']
    except BaseException as E:
        _logging.error('读取config.json中设备容量参数失败,错误原因为{}'.format(E))
        raise Exception

    c_pv = 1000
    c_el = 2240  # 1000标方一千万
    c_fc = 5000 
    c_hst = 3000
    c_ht = 5
    c_ct = 5
    c_eb = 80
    c_hp = 9625
    c_ac = 200

    crf_pv = crf(20)
    crf_el = crf(10)
    crf_fc = crf(10)
    crf_hst = crf(10)
    crf_ht = crf(20)
    crf_ct = crf(20)
    crf_eb = crf(10)
    crf_hp = crf(10)

    crf_ac = crf(10)
    # 初始化边界上下限参数
    try:
        t_ht_max = parameter_json['device']['ht']['t_max']
        t_ht_min = parameter_json['device']['ht']['t_supply']
        t_ct_max = parameter_json['device']['ct']['t_max']
        t_ct_min = parameter_json['device']['ct']['t_min']

    except BaseException as E:
        _logging.error('读取config.json中边界上下限参数失败,错误原因为{}'.format(E))
        raise Exception
    # 初始化价格
    try:
        lambda_ele_in = parameter_json['price']['ele_TOU_price'] * 365
        #lambda_ele_in = [0.5109 for _ in range(period)]
        lambda_ele_out = parameter_json['price']['power_sale']
        hydrogen_price = parameter_json['price']['hydrogen_price']
    except BaseException as E:
        _logging.error('读取config.json中价格参数失败,错误原因为{}'.format(E))
        raise Exception

    # 初始化负荷
    try:
        # p_load = list(load_json['ele_load'])
        p_load = list(load_json['ele_newload'])
        g_load = list(load_json['g_load'])
        q_load = list(load_json['q_load'])
        water_load = list(load_json['water_load'])

        # if isolate == 0:
        #     k_g_load = 4600/max(g_load)
        #     k_q_load = 2400/max(q_load)
        #     #k_ele_load = 4485/max(p_load)
        #     g_load = [k_g_load*i for i in g_load]
        #     q_load = [k_q_load*i for i in q_load]
        # #print(k_g_load,k_q_load)
        # #exit(0)
        # else:

        # k_g_load = 4600/max(g_load)/1.245 *0.69/9.55
        k_g_load = 40000/max(g_load)/1.245 *0.69/9.55
        k_q_load = 2400/max(q_load)/1.245 *0.69/2.2
        g_load = [k_g_load*i for i in g_load]
        q_load = [k_q_load*i for i in q_load]
        p_load = [i+80 for i in p_load]


        # 0802更新负荷
        p_old = p_load
        # 电负荷，B1 284.5峰值的水泵，跟随冷负荷曲线。 F1 30的展厅早八晚六，5的其余负荷，F2:1的会议室早八晚六，80的数据机房全天开
        # k_q = 1405/max(q_load)
        k_q = 1805/max(q_load)
        q_load = ([k_q*i + 80 for i in q_load])
        q_load = q_load[2:]+q_load[:2]
        p_load = [284.5*max(g_load[i]/max(g_load),q_load[i]/max(q_load)) + 5/max(p_old)*p_old[i] + 30+5 +80 if i%24 >= 8 and i%24 <= 18 else 284.5*max(g_load[i]/max(g_load),q_load[i]/max(q_load)) + 5/max(p_old)*p_old[i]+80 for i in range(period)]
        # print(284.5*max(g_load[6576]/max(g_load),q_load[6576]/max(q_load)) ,5/max(p_old)*p_old[6576] ,80)


        ##0804需求，八月六号 6576开始一天的冷负荷峰值1400
        #每天的峰值都是1400
        q_load_old = list(np.array(q_load))
        # for i in range(30):
        #     k_peak_0804 = (1400-10*i)/max(q_load[6576-i*24:6576-i*24+24])
        #     q_load[6576-i*24:6576-i*24+24] = [k_peak_0804*i for i in q_load[6576-i*24:6576-i*24+24]]
        # print(sum(q_load),sum(q_load_old))
        #*max(g_load[i]/max(g_load),q_load[i]/max(q_load)) 284.5的跟随
        # 整理冷负荷
        
        
        ###
        F2_p_load = [ 20+5 if i%24 >= 8 and i%24 <= 18 else 20 for i in range(period)]
        
        #F2_p_use  = [ 80+5 if i%24 >= 8 and i%24 <= 18 else 80 for i in range(period)]
        # F2_q_load = [(q_load[i]*48/(21+48))/3 if i%24 >= 8 and i%24 <= 18 else q_load[i]/3 for i in range(period)]
        # F2_q_load = [q_load_old[i]*48/(21+48)/3 for i in range(period)]
        # F2_q_load = [(q_load_old[i]-h_fc[i].x*k_fc_g)*48/(218+48)/3 for i in range(period)]
        F2_g_load = [g_load[i]*48/(21+48)/3 for i in range(period)]

        F1_p_load = [30+5/max(p_old)*p_old[i]+60 if i%24 >= 8 and i%24 <= 18 else 60+5/max(p_old)*p_old[i] for i in range(period)]
        #F1_p_use  = [30+5/max(p_old)*p_old[i] if i%24 >= 8 and i%24 <= 18 else 5/max(p_old)*p_old[i] for i in range(period)]
        # F1_q_load = [(q_load[i]*48/(21+48))/3 if i%24 >= 8 and i%24 <= 18 else q_load[i]/3 for i in range(period)]
        # F1_q_load = [q_load_old[i]*48/(21+48)/3 for i in range(period)]
        # F1_q_load = [(q_load_old[i]-h_fc[i].x*k_fc_g)*48/(218+48)/3 for i in range(period)]

        F1_g_load = [g_load[i]*48/(21+48)/3 for i in range(period)]

        B1_p_load = [284.5*max(g_load[i]/max(g_load),q_load_old[i]/max(q_load_old)) for i in range(period)]
        #B1_p_use  = [284.5*max(g_load[i]/max(g_load),q_load[i]/max(q_load)) for i in range(period)]
        # B1_q_load = [(q_load[i]*48/(21+48))/3 if i%24 >= 8 and i%24 <= 18 else q_load[i]/3 for i in range(period)]
        # B1_q_load = [q_load_old[i]*48/(21+48)/3 for i in range(period)]
        # B1_q_load = [(q_load_old[i]-h_fc[i].x*k_fc_g)*48/(218+48)/3 for i in range(period)]
        B1_g_load = [g_load[i]*48/(21+48)/3 for i in range(period)]

        p_load_17 = [0 for i in range(period)]
        # q_load_17 = [q_load[i]*3/(21+48) if i%24 >= 8 and i%24 <= 18 else 0 for i in range(period)]
        # q_load_17 = [(q_load[i] - q_load_old[i]*48/(21+48))*3/21 for i in range(period)]
        # q_load_17 = [(q_load_old[i]-h_fc[i].x*k_fc_g)*167/(218+48) for i in range(period)]
        g_load_17 = [g_load[i]*3/(21+48) for i in range(period)]

        p_load_18 = [0 for i in range(period)]
        # q_load_18 = [q_load[i]*18/(21+48) if i%24 >= 8 and i%24 <= 18  else 0 for i in range(period)]
        # q_load_18 = [(q_load[i] - q_load_old[i]*48/(21+48))*18/21 for i in range(period)]
        # q_load_18 = [(q_load_old[i]-h_fc[i]*k_fc_g)*51/(218+48) for i in range(period)]
        g_load_18 = [g_load[i]*18/(21+48) for i in range(period)]
        # print(p_load[6576],F2_p_load[6576],F1_p_load[6576],B1_p_load[6576])
        # print(284.5*max(g_load[6576]/max(g_load),q_load_old[6576]/max(q_load_old)) ,5/max(p_old)*p_old[6576] ,80)
        # exit(0)
        #p_load2 = [i*5 for i in p_load]
        #p_load = [i/5 for i in p_load]
        water_load = [0 for _ in water_load]
        print(sum(p_load),sum(g_load),sum(q_load))
        #exit(0)

        r_solar =  [0 for i in range(8760+24)]
        with open("Input/solar.csv") as renewcsv:
            renewcsv.readline()
            renewcsv.readline()
            renewcsv.readline()
            renew = csv.DictReader(renewcsv)
            
            i=0
            for row in renew:

                r_solar[i] += float(row['electricity'])
        
                i+=1
        r_solar = r_solar[-8:]+r_solar[:-8]
        solar = r_solar
    except BaseException as E:
        _logging.error('读取负荷文件中电冷热光参数失败,错误原因为{}'.format(E))
        raise Exception
    # 初始化储能
    # try:
    #     hydrogen_bottle_max_start = storage_begin_json['hydrogen_bottle_max'][0]  #气瓶
    #     hst_kg_start = storage_begin_json['hst_kg'][0]  # 缓冲罐剩余氢气
    #     t_ht_start = storage_begin_json['t_ht'][0]  # 热水罐
    #     t_ct_start = storage_begin_json['t_ct'][0]  # 冷水罐

    #     hydrogen_bottle_max_final = storage_end_json['hydrogen_bottle_max'][0] #气瓶
    #     hst_kg_final = storage_end_json['hst_kg'][0]  # 缓冲罐剩余氢气
    #     t_ht_final = storage_end_json['t_ht'][0]  # 热水罐
    #     t_ct_final = storage_end_json['t_ct'][0]  # 冷水罐
    # except BaseException as E:
    #     _logging.error('读取储能容量初始值和最终值失败,错误原因为{}'.format(E))
    #     raise Exception
    # 通过gurobi建立模型
    try:
        m = gp.Model("bilinear")
    except BaseException as E:
        _logging.error('gurobi创建优化模型失败{}'.format(E))
        raise Exception

    # 添加变量
    #opex = m.addVar(vtype=GRB.CONTINUOUS, lb=0, name="opex")
    opex = [m.addVar(vtype=GRB.CONTINUOUS, lb=0, name="opex_"+str(i)) for i in range(period)]
    t_ht = [m.addVar(vtype=GRB.CONTINUOUS, lb=0, name=f"t_ht{t}") for t in range(period)] # temperature of hot water tank
    t_ht_l = [m.addVar(vtype=GRB.CONTINUOUS, lb=0, name=f"t_ht_l{t}") for t in range(period)] # temperature of hot water tank in last time
    t_ct = [m.addVar(vtype=GRB.CONTINUOUS, lb=0, name=f"t_ct{t}") for t in range(period)] # temperature of hot water tank
    t_ct_l = [m.addVar(vtype=GRB.CONTINUOUS, lb=0, name=f"t_ct_l{t}") for t in range(period)] # temperature of hot water tank in last time

    g_fc = [m.addVar(vtype=GRB.CONTINUOUS, lb=0, name=f"g_fc{t}") for t in range(period)] # heat generated by fuel cells
    p_fc = [m.addVar(vtype=GRB.CONTINUOUS, lb=0, name=f"p_fc{t}") for t in range(period)]
    h_fc = [m.addVar(vtype=GRB.CONTINUOUS, lb=0, name=f"h_fc{t}") for t in range(period)] # hydrogen used in fuel cells

    p_hp = [m.addVar(vtype=GRB.CONTINUOUS, lb=0, name=f"p_hp{t}") for t in range(period)] # power consumption of heat pumps
    p_hp_g = [m.addVar(vtype=GRB.CONTINUOUS, lb=0, name=f"p_hp_g{t}") for t in range(period)] # power consumption of heat pumps
    p_hp_q = [m.addVar(vtype=GRB.CONTINUOUS, lb=0, name=f"p_hp_q{t}") for t in range(period)] # power consumption of heat pumps
    g_hp = [m.addVar(vtype=GRB.CONTINUOUS, lb=0, name=f"g_hp{t}") for t in range(period)] # heat generated by heat pumps
    q_hp = [m.addVar(vtype=GRB.CONTINUOUS, lb=0, name=f"q_hp{t}") for t in range(period)] # heat generated by heat pumps

    h_el = [m.addVar(vtype=GRB.CONTINUOUS, lb=0, name=f"h_el{t}") for t in range(period)] # hydrogen generated by electrolyzer
    p_el = [m.addVar(vtype=GRB.CONTINUOUS, lb=0, name=f"p_el{t}") for t in range(period)] # power consumption by electrolyzer

    h_sto = [m.addVar(vtype=GRB.CONTINUOUS, lb=0, name=f"h_sto{t}") for t in range(period)] # hydrogen storage
    h_sto_l = [m.addVar(vtype=GRB.CONTINUOUS, lb=0, name=f"h_sto_l{t}") for t in range(period)] # last time hydrogen storage
    h_pur = [m.addVar(vtype=GRB.CONTINUOUS, lb=0, name=f"h_pur{t}") for t in range(period)] # hydrogen purchase

    p_pur = [m.addVar(vtype=GRB.CONTINUOUS, lb=0, name=f"p_pur{t}") for t in range(period)] # power purchase

    p_pump = [m.addVar(vtype=GRB.CONTINUOUS, lb=0, name=f"p_pump{t}") for t in range(period)]
    p_cur = [m.addVar(vtype=GRB.CONTINUOUS, lb=0, name=f"p_cur{t}") for t in range(period)] 
    p_eb = [m.addVar(vtype=GRB.CONTINUOUS, lb=0, name=f"p_eb{t}") for t in range(period)] # power consumption by ele boiler
    g_eb = [m.addVar(vtype=GRB.CONTINUOUS, lb=0, name=f"g_eb{t}") for t in range(period)] # heat generated by ele boiler

    p_pv = [m.addVar(vtype=GRB.CONTINUOUS, lb=0, name=f"p_pv{t}") for t in range(period)] # power generate by PV
    #p_load2 = [m.addVar(vtype=GRB.CONTINUOUS, lb=0, name=f"p_load2{t}") for t in range(period)] # sale
    #water_abs = [m.addVar(vtype=GRB.CONTINUOUS, lb=0, name=f"water_abs{t}") for t in range(period)] # water absorption
    # h_stooo = [m.addVar(vtype=GRB.CONTINUOUS, lb=0, name=f"h_stooo{t}") for t in range(period)] # hydrogen storage
    p_slack = [m.addVar(vtype=GRB.CONTINUOUS, lb=0, name=f"p_sol{t}") for t in range(period)] # slack power
    a_pv = m.addVar(name = "max_pv")
    p_el_max = m.addVar(name = "p_el_max")
    p_fc_max = m.addVar(name = "p_fc_max")
    hst_max= m.addVar(name = "max_hst")
    #m_ht = m.addVar(name = "m_ht")
    #m_ct = m.addVar(name = "m_ct")
    m_ht = 500000
    m_ct = 500000
    p_eb_max = m.addVar(name = "p_eb_max")
    p_hp_max = m.addVar(name = "p_hp_max")
    capex = m.addVar(name = 'capex')
    p_ac_max = m.addVar(name = "p_ac_max")
    m.addConstr(capex == crf_pv*a_pv*c_pv + crf_el*c_el*p_el_max + crf_fc*c_fc*p_fc_max + crf_hst*c_hst*hst_max + crf_ht*c_ht*m_ht + crf_ct*c_ct*m_ct
        + crf_eb*c_eb*p_eb_max + crf_hp*c_hp*p_hp_max + crf_ac*c_ac*p_ac_max)
   
    m.addConstr(m_ht<=m_ht_max)
    m.addConstr(m_ct<=m_ct_max)
    
    if hydrogen_only == 0:
        m.addConstr(p_el_max<=p_el_max_max)
        m.addConstr(p_fc_max<=p_fc_max_max)
        m.addConstr(hst_max<=hst_max_max)
        m.addConstr(a_pv<=a_pv_max)
    m.addConstr(p_eb_max<=p_eb_max_max)
    m.addConstr(p_hp_max<=p_hp_max_max)

    # if hydrogen_bottle_max_final - hydrogen_bottle_max_start>=-1:
    #     m.addConstr(gp.quicksum(h_pur) <= hydrogen_bottle_max_final - hydrogen_bottle_max_start)
    # else:
    #     m.addConstr(gp.quicksum(h_pur) == 0)

    #print(storage_end_json['end_slack'][0])
    # if storage_end_json['end_slack'][0] == False:
    #     m.addConstr(t_ht[-1] == t_ht_final)
    #     m.addConstr(t_ct[-1] == t_ct_final)
    #     m.addConstr(h_sto[-1] == hst_kg_final)
    # else:
    #     m.addConstr(t_ht[-1] >= t_ht_start * (1-slack_ht))
    #     m.addConstr(t_ht[-1] <= t_ht_start * (1+slack_ht))
    #     m.addConstr(t_ct[-1] >= t_ct_start * (1-slack_ct))
    #     m.addConstr(t_ct[-1] <= t_ct_start * (1+slack_ct))
    #     m.addConstr(h_sto[-1] >= hst_kg_start * (1-slack_hsto))
    #     m.addConstr(h_sto[-1] <= hst_kg_start * (1+slack_hsto))
    # # 储能约束
    # m.addConstr(t_ht_l[0] == t_ht_start)
    # m.addConstr(t_ct_l[0] == t_ct_start)
    # m.addConstr(h_sto_l[0] == hst_kg_start)

    p_ac_max = m.addVar(name = "p_ac_max")
    p_ac = [m.addVar(vtype=GRB.CONTINUOUS, lb=0, name=f"p_ac{t}") for t in range(period)]
    q_ac = [m.addVar(vtype=GRB.CONTINUOUS, lb=0, name=f"q_ac{t}") for t in range(period)]
    # device
    if device == 1:
        m.addConstr(a_pv == 4792)
        m.addConstr(p_el_max == 48)
        m.addConstr(p_fc_max == 640)
        m.addConstr(hst_max == 466)
        #m.addConstr(m_ht == 500000)
        #m.addConstr(m_ct == 500000)
        m.addConstr(p_eb_max == 1400)
        m.addConstr(p_hp_max == 235)
    if hydrogen_only == 1:
        #m.addConstr(m_ht == 0)
        #m.addConstr(m_ct == 0)
        m.addConstr(p_eb_max == 0)
        m.addConstr(p_hp_max == 0)
        
    else:
        m.addConstr(p_ac_max == 0)
    # 氢气购买约束

    # 热泵同开同关约束
    

    #m.addConstr(gp.quicksum(h_pur) <= 8000)
    m.addConstrs(t_ht[i] == t_ht_l[i+1] for i in range(period-1))
    m.addConstrs(t_ct[i] == t_ct_l[i+1] for i in range(period-1))
    m.addConstrs(h_sto[i] == h_sto_l[i+1] for i in range(period-1))
    m.addConstr(t_ht_l[0] == t_ht[-1])
    m.addConstr(t_ct_l[0] == t_ct[-1])
    m.addConstr(h_sto_l[0] == h_sto[-1])
    #每日买氢量上限
    # for i in range(150,250):
    #     for j in range(1+int(i/24)*24,int(i/24)*24+5):
    #         m.addConstr(h_pur[j] == 0)
    #     for j in range(7 + int(i/24)*24,int(i/24)*24+15):
    #         m.addConstr(h_pur[j] == 0)
    #     for j in range(23 + int(i/24)*24,int(i/24)*24+24):
    #         m.addConstr(h_pur[j] == 0)
        #m.addConstr(h_pur[i*24]<= hst_max-48)
    #for i in range(period):
        #m.addConstr(h_stooo[i] == 48+h_sto[i]+ gp.quicksum([ h_pur[j] for j in range(24*int(i/24),24*int(i/24)+24)]) - gp.quicksum([h_pur[j] for j in range(24*int(i/24),i)]) )

    for i in range(period):
        m.addConstr(p_fc[i] + p_pur[i] + p_pv[i] + p_slack[i] == p_cur[i] + p_el[i] + p_eb[i] + p_hp[i]  + p_pump[i] + p_load[i] + p_ac[i])
        #m.addConstr(c*m_ht*(t_ht[i] - t_ht_l[i] - ht_loss * (t_ht_l[i] - t_ht_wetbulb)) + g_load[i] == g_fc[i] + g_hp[i] + g_eb[i])
        m.addConstr(c*m_ht*(t_ht[i] - t_ht_l[i] + 0.001*(t_ht_l[i]-t_ht_min)) + g_load[i] + water_load[i] == g_fc[i] + g_hp[i] + g_eb[i])

        #m.addConstr(c*m_ct*(t_ct[i] - t_ct_l[i] - ct_loss * (t_ct_l[i] - t_ct_wetbulb)) + q_hp[i] == q_load[i])
        m.addConstr(c*m_ct*(t_ct[i] - t_ct_l[i] - 0.001*(t_ct_l[i] - t_ct_min)) + q_hp[i] + q_ac[i] == q_load[i])

        m.addConstr(h_sto[i] - h_sto_l[i] == h_pur[i] + h_el[i] - h_fc[i])
        #m.addConstr(h_stooo[i] <= hst_max)
        # m.addConstr(p_slack[i] <= p_load[i])
    # 每一时段约束
    for i in range(period):
        #isloate 
        if isolate == 1:
            m.addConstr(p_pur[i] == 0)
        #m.addConstr(p_sol[i] <= p_load[i])
        # pump
        #m.addConstr(water_abs[i] == p_pump[i] * 0.9)
        #m.addConstr(10*p_pump[i] == g_hp[i] + q_hp[i])
        
        # 上下限约束
        m.addConstr(t_ht[i] >= t_ht_min)
        m.addConstr(t_ht[i] <= t_ht_max)
        m.addConstr(t_ct[i] >= t_ct_min)
        m.addConstr(t_ct[i] <= t_ct_max)
        m.addConstr(p_fc[i] <= p_fc_max)
        m.addConstr(p_el[i] <= p_el_max)
        m.addConstr(p_eb[i] <= p_eb_max)
        m.addConstr(p_hp[i] <= p_hp_max)
        m.addConstr(h_sto[i]<= hst_max)


        # 徐老师约束
        # 供冷季
        # if i%24<=19 or i%24>=18 and q_load[i]>=0:
        #     m.addConstr(p_hp[i] == 0)
        #m.addConstr(p_pur[i] <= p_pur_max)



        m.addConstr(p_ac[i] <= p_ac_max)
        # 能量平衡

        # 设备约束
        ## fc
        m.addConstr(p_fc[i] <= p_fc_max)
        m.addConstr(p_fc[i] == k_fc_p * h_fc[i])
        m.addConstr(g_fc[i] <= k_fc_g * h_fc[i])
        ## hp
        # m.addConstr(p_hp[i] <= 0.8*p_hp_max)## 0803修改
        m.addConstr(q_hp[i] == k_hp_q * p_hp_g[i])
        m.addConstr(g_hp[i] == k_hp_g * p_hp_q[i])
        m.addConstr(p_hp[i] == p_hp_g[i] + p_hp_q[i])
        ## el
        m.addConstr(p_el[i] <= p_el_max)
        m.addConstr(h_el[i] == k_el * p_el[i])
        ## eb
        m.addConstr(p_eb[i] <= p_eb_max)
        m.addConstr(g_eb[i] == k_eb * p_eb[i])
        ## pump
        #m.addConstr(p_pump[i] == k_pump * mass_flow[i])
        ## pv
        m.addConstr(p_pv[i] <= solar[i] * a_pv * k_pv)
        ## ac
        m.addConstr(q_ac[i] == k_ac * p_ac[i])
        ## opex 
        m.addConstr(opex[i] == hydrogen_price * h_pur[i] + p_pur[i] * lambda_ele_in[i] )
    # set objective
    
    m.setObjective(gp.quicksum(opex) + capex - lambda_ele_out*gp.quicksum(p_cur) + gp.quicksum(p_slack)*10000000, GRB.MINIMIZE)
    m.params.MIPGap = 0.01
    m.Params.LogFile = "testlog.log"
    m.params.NonConvex = 2

    try:
        m.optimize()
        _logging.info("success optimize")
    except gp.GurobiError as e:
        print("Optimize failed due to non-convexity")
        _logging.error(e)
    if m.status == GRB.INFEASIBLE or m.status == 4:
        print('Model is infeasible')
        m.computeIIS()
        m.write('Output/model.ilp')
        print("Irreducible inconsistent subsystem is written to file 'model.ilp'")
        exit(0)

    # 计算一些参数
    opex_without_opt_ele = [lambda_ele_in[i]*(p_load[i]+q_load[i]/k_hp_q+g_load[i]/k_eb+water_load[i]/k_eb) for i in range(period)]
    revenue_ele = sum([p_load[i]*lambda_ele_in[i] for i in range(period)])
    revenue_heat = 6.3*5*6900
    revenue_cold = 16*3*6900
    
    h_stoo = [0 for _ in range(period)]
    area = 6900
    kk = period
    opex_without_opt_list = [max(0,p_cur[i].x+ p_load[i]+g_load[i]/k_eb+water_load[i]/k_eb+q_load[i]/k_hp_q -p_pv[i].x)*0.896  for i in range(period)]
    for jj in range(period):
        j = period - jj - 1
        tmp = h_sto[j].x + sum([h_pur[i].x for i in range(j,kk)])
        if tmp >= 466-141:
            h_stoo[j] = 141
            kk = j
        else:
            h_stoo[j] = 141+tmp
    month_sumup_opex = []
    month_sumup_pload = []
    month_sumup_hpur = []
    month_sumup_huse = []
    month_sumup_greenpoweruse = []
    month_sumup_pv = []
    month_sumup_fc = []
    month_sumup_hel = []
    ttmp = 0
    ttmp_pload = 0
    ttmp_hpur = 0
    ttmp_huse = 0
    ttmp_greenpoweruse = 0
    ttmp_pv = 0
    ttmp_fc = 0
    ttmp_hel = 0
    for jj in range(period):
        #month_sumup_opex
        tmp = datetime.datetime.strptime("2022"+"-"+str(int(jj/24)+1),"%Y-%j")
        
        if tmp.day == 1 and jj%24 == 0 and jj != 0:
            month_sumup_opex.append(ttmp)
            month_sumup_pload.append(ttmp_pload)
            month_sumup_hpur.append(ttmp_hpur)
            month_sumup_huse.append(ttmp_huse)
            month_sumup_greenpoweruse.append(ttmp_greenpoweruse)
            month_sumup_pv.append(ttmp_pv)
            month_sumup_fc.append(ttmp_fc)
            month_sumup_hel.append(ttmp_hel)
            #print(ttmp)
            ttmp = 0
            ttmp_pload = 0
            ttmp_hpur = 0
            ttmp_huse = 0
            ttmp_greenpoweruse = 0
            ttmp_pv = 0
            ttmp_fc = 0
        else:
            ttmp += opex[jj].x
            ttmp_hpur += h_pur[jj].x
            ttmp_pload += p_pv[jj].x + p_fc[jj].x## 0803修正
            ttmp_huse += h_fc[jj].x
            ttmp_greenpoweruse += p_pv[jj].x + p_fc[jj].x
            ttmp_pv += p_pv[jj].x
            ttmp_fc += p_fc[jj].x
            ttmp_hel += h_el[jj].x
    month_sumup_opex.append(ttmp)
    month_sumup_pload.append(ttmp_pload)
    month_sumup_hpur.append(ttmp_hpur)
    month_sumup_huse.append(ttmp_huse)
    month_sumup_greenpoweruse.append(ttmp_greenpoweruse)
    month_sumup_pv.append(ttmp_pv)
    month_sumup_fc.append(ttmp_fc)
    month_sumup_hel.append(ttmp_hel)

    z_heat = [1 if g_load[i] > 0 else 0 for i in range(period)]
    z_cold = [1 if q_load[i] > 80 else 0 for i in range(period)]
    year_capex = 2220000
    carbon_fake_rate = 8640000/(sum([p_load[i]+g_load[i]/k_eb+water_load[i]/k_eb+q_load[i]/k_hp_q for i in range(period)])*0.5837)
    
    F2_q_load = [(q_load_old[i]-h_fc[i].x*k_fc_g/4)*48/(218+48)/3 if z_cold[i] == 1 else 0 for i in range(period)]
    F1_q_load = [(q_load_old[i]-h_fc[i].x*k_fc_g/4)*48/(218+48)/3 if z_cold[i] == 1 else 80 for i in range(period)]
    B1_q_load = [(q_load_old[i]-h_fc[i].x*k_fc_g/4)*48/(218+48)/3 if z_cold[i] == 1 else 0 for i in range(period)]

    q_load_17 = [(q_load_old[i]-h_fc[i].x*k_fc_g/4)*167/(218+48) if z_cold[i] == 1 else 0 for i in range(period)]
    q_load_18 = [(q_load_old[i]-h_fc[i].x*k_fc_g/4)*51/(218+48) if z_cold[i] == 1 else 0 for i in range(period)]

    dict_plot = {
        # operational day cost
        #'opex_without_system':sum(opex_without_opt),#没有能源站的运行成本，负荷直接加
        #'opex_with_ele':sum([max(0,p_load[i]+g_load[i]/k_eb+water_load[i]/k_eb+q_load[i]/k_hp_q -p_pv[i].x)*lambda_ele_in[i] for i in range(period)]),
        'opex_with_ele':[(p_load[i]+g_load[i]/k_eb+water_load[i]/k_eb+q_load[i]/k_hp_q)*lambda_ele_in[i] for i in range(period)],

        'compare_ele_op_save_part':sum([opex[i].x for i in range(period)])/sum(opex_without_opt_list),
        # 'revenue':revenue_ele+revenue_heat+revenue_cold, # 总收益
        # 'payback_period':(a_pv.x*c_pv + c_el*p_el_max.x + c_fc*p_fc_max.x + c_hst*hst_max.x + c_ht*m_ht + c_ct*m_ct + c_eb*p_eb_max.x + c_hp*p_hp_max.x +c_ac*p_ac_max.x)/(revenue_ele+revenue_heat+revenue_cold - sum([opex[i].x for i in range(period)])),# 回收期
        
        # carbon emmision reduce
        #'cer_rate':sum([p_pur[i].x for i in range(period)])/sum([p_load[i]+g_load[i]/k_eb+water_load[i]/k_eb+q_load[i]/k_hp_q for i in range(period)]),
        # 综合能效
        'renewable_energy_part':sum([p_pv[i].x + p_fc[i].x for i in range(period)])/sum([p_pv[i].x + p_fc[i].x + p_pur[i].x for i in range(period)]), #可再生能源占比
        'efficiency':(sum(p_load+q_load+g_load+water_load)+sum([p_cur[i].x for i in range(period)])) /(sum([p_pv[i].x + 33.3 * h_pur[i].x + p_pur[i].x for i in range(period)])),#+ sum([p_slack[j].x for j in range(period)]))
        'cold_efficiency':(sum([p_cur[i].x + p_load[i] + q_load[i] + g_load[i] + water_load[i] if z_cold[i] == 1 else 0 for i in range(period)])) /(sum([p_pv[i].x + 33.3 * h_pur[i].x + p_pur[i].x if z_cold[i] == 1 else 0 for i in range(period)])),
        'old_efficiency':sum(p_load+q_load+g_load+water_load)/(sum([p_load[i]+q_load[i]/k_hp_q+g_load[i]/k_eb+water_load[i]/k_eb for i in range(period)])),
        # 规划结果
        "capex":a_pv.x*c_pv + c_el*p_el_max.x + c_fc*p_fc_max.x + c_hst*hst_max.x + c_ht*m_ht + c_ct*m_ct + c_eb*p_eb_max.x + c_hp*p_hp_max.x +c_ac*p_ac_max.x,
        'a_pv':a_pv.x,
        'p_fc_max':p_fc_max.x,
        'p_el_max':p_el_max.x,
        'p_eb_max':p_eb_max.x,
        'p_hp_max':p_hp_max.x,
        'hst_max':hst_max.x,
        #'p_ac_max':p_ac_max.x,
        'm_ct':m_ct,
        'm_ht':m_ht,
        # 可再生能源占比

        # 分时 运行费用
        # 'hourly_opex_without_opt': 1,
        # 碳减排
        'cer':sum([p_pur[i].x for i in range(period)])/(sum([p_load[i]+g_load[i]/k_eb+water_load[i]/k_eb+q_load[i]/k_hp_q for i in range(period)])*carbon_fake_rate),
        'carbon_emmision':sum([p_pur[i].x for i in range(period)])*0.5837*carbon_fake_rate,
        'ele_carbon_emmision':sum([p_load[i]+g_load[i]/k_eb+water_load[i]/k_eb+q_load[i]/k_hp_q for i in range(period)])*0.5837*carbon_fake_rate,
        'center_heat_carbon_emmision':sum([p_load[i]*0.5837+g_load[i]/k_eb*0.351+water_load[i]/k_eb*0.351+q_load[i]/k_hp_q*0.5837 for i in range(period)]),
        'ele_opcost':sum([(p_load[i]+g_load[i]/k_eb+water_load[i]/k_eb+q_load[i]/k_hp_q)*lambda_ele_in[i] for i in range(period)]),
        # 'centure_carbon':sum([p_load[i]+q_load[i]/k_hp_q for i in range(period)])*0.5837 + sum([(g_load[i]+ water_load[i])*0.351 for i in range(period)]), # 集中式冷热碳排放
        # ele


        # 单位面积能耗
        #0804 'hourly_area_energy_use':[(p_load[i]+g_load[i]+water_load[i]+q_load[i])/area for i in range(period)],
        'hourly_area_energy_use':[(p_load[i]+g_load[i]+water_load[i]+q_load[i])/area for i in range(period)],
        'dayly_area_energy_use':[sum([p_load[i]+g_load[i]+water_load[i]+q_load[i] for i in range(int(j/24)*24,int(j/24)*24+24)])/area for j in range(period)],


        #### needed list
        'day_sum_p_load':[sum([p_pv[i].x+p_fc[i].x for i in range(int(i/24)*24,i+1)]) for i in range(period)], #每日累计电耗
        'day_sum_q_load':[sum(q_load[int(i/24)*24:i+1]) for i in range(period)], #
        'day_sum_g_load':[sum(g_load[int(i/24)*24:i+1]) for i in range(period)], #每日累计热耗

        # 'h_hst':[h_sto[i].x for i in range(period)],
        'h_stoo':h_stoo,

        'day_sum_p_load_area':[sum([p_pv[i].x+p_fc[i].x for i in range(int(i/24)*24,i+1)])/area for i in range(period)], #每日累计电耗
        'day_sum_q_load_area':[sum(q_load[int(i/24)*24:i+1])/area for i in range(period)], #
        'day_sum_g_load_area':[sum(g_load[int(i/24)*24:i+1])/area for i in range(period)], #每日累计热耗

        'hourly_energystation_p_load':[F2_p_load[i]+F1_p_load[i]+B1_p_load[i]+p_hp[i].x for i in range(period)], #每日累计电耗
        'hourly_energystation_p_stable_load':[F2_p_load[i]+F1_p_load[i]+B1_p_load[i] for i in range(period)], 
        'hourly_energystation_q_load':[F2_q_load[i]+F1_q_load[i]+B1_q_load[i]+h_fc[i].x*k_fc_g for i in range(period)], 
        'hourly_energystation_g_load':[F2_g_load[i]+F1_g_load[i]+B1_g_load[i] for i in range(period)], #每日累计热耗


        'hourly_village_p_load':[p_cur[i].x for i in range(period)], #每日累计电耗
        'hourly_village_q_load':[q_load_17[i]+q_load_18[i] for i in range(period)], 
        'hourly_village_g_load':[g_load_17[i]+g_load_18[i] for i in range(period)],  #每日累计热耗

        # 'F2_p_load':[((p_load[i]-82.07)*48/(21+48)+82.07)/3 for i in range(period)],
        # 'F2_q_load':[(q_load[i]*48/(21+48))/3 if abs(p_load[i]-82.07) >= 0.1 else q_load[i]/3 for i in range(period)], 
        # 'F2_g_load':[g_load[i]*48/(21+48)/3 for i in range(period)],
        # 'F1_p_load':[((p_load[i]-82.07)*48/(21+48)+82.07)/3 for i in range(period)],
        # 'F1_q_load':[(q_load[i]*48/(21+48))/3 if abs(p_load[i]-82.07) >= 0.1 else q_load[i]/3 for i in range(period)], 
        # 'F1_g_load':[g_load[i]*48/(21+48)/3 for i in range(period)],
        # 'B1_p_load':[((p_load[i]-82.07)*48/(21+48)+82.07)/3 for i in range(period)],
        # 'B1_q_load':[(q_load[i]*48/(21+48))/3 if abs(p_load[i]-82.07) >= 0.1 else q_load[i]/3 for i in range(period)], 
        # 'B1_g_load':[g_load[i]*48/(21+48)/3 for i in range(period)],
        # '17_p_load':[(p_load[i]-82.07)*3/(21+48) for i in range(period)],
        # '17_q_load':[q_load[i]*3/(21+48) if abs(p_load[i]-82.07) >= 0.1 else 0 for i in range(period)], 
        # '17_g_load':[g_load[i]*3/(21+48) for i in range(period)], 
        # '18_p_load':[(p_load[i]-82.07)*18/(21+48) for i in range(period)],
        # '18_q_load':[q_load[i]*18/(21+48) if abs(p_load[i]-82.07) >= 0.1 else 0 for i in range(period)], 
        # '18_g_load':[g_load[i]*18/(21+48) for i in range(period)], 
        'F2_p_load':F2_p_load,
        'F2_p_stable_load':F2_p_load,
        'F2_q_load':F2_q_load,
        'F2_g_load': F2_g_load,

        'F1_p_load': F1_p_load,
        'F1_p__stable_load': F1_p_load,
        'F1_q_load': F1_q_load,
        'F1_g_load': F1_g_load,

        'B1_p_load': [B1_p_load[i]+p_hp[i].x for i in range(period)],
        'B1_p_stable_load':B1_p_load,
        'B1_q_load': B1_q_load,
        'B1_g_load': B1_g_load,

        '17_p_load': [p_cur[i].x*3/21 for i in range(period)],
        '17_p_stable_load': [0 for i in range(period)],
        '17_q_load': q_load_17,
        '17_g_load': g_load_17,

        '18_p_load':[p_cur[i].x*18/21 for i in range(period)],
        '18_p_stable_load': [0 for i in range(period)],
        '18_q_load': q_load_18,
        '18_g_load': g_load_18,

        'dayly_efficiency':[ sum([p_cur[i].x+p_load[i]+q_load[i]+g_load[i]+water_load[i]  for i in range(int(j/24)*24,int(j/24)*24+24)])/sum([0.00000000001+p_pv[i].x + 33.3 * h_pur[i].x + p_pur[i].x for i in range(int(j/24)*24,int(j/24)*24+24)]) for j in range(period)],
        'hourly_efficiency':[(p_cur[i].x+p_load[i]+q_load[i]+g_load[i]+water_load[i])/(0.00000000001+p_pv[i].x + 33.3 * h_pur[i].x + p_pur[i].x) for i in range(period)],
        'hourly_opex':[opex[i].x for i in range(period)],# 经优化的运行成本
        'hourly_opex_without_opt':opex_without_opt_list,# 未经优化的运行成本 #+ lambda_ele_out*min(0,p_load[i]+g_load[i]/k_eb+water_load[i]/k_eb+q_load[i]/k_hp_q -p_pv[i].x)

        'season_heat_unit_cost':sum([opex[i].x if z_heat[i] == 1 else 0 for i in range(period)])/sum([p_load[i]+p_cur[i].x+g_load[i] if z_heat[i] == 1 else 0 for i in range(period)]),
        'season_cold_unit_cost':sum([opex[i].x if z_cold[i] == 1 else 0 for i in range(period)])/sum([p_load[i]+p_cur[i].x+q_load[i] if z_cold[i] == 1 else 0 for i in range(period)]),
        'season_non_unit_cost':sum([opex[i].x if z_heat[i]+z_cold[i] == 0 else 0 for i in range(period)])/sum([p_load[i]+p_cur[i].x  if z_heat[i]+z_cold[i] == 0 else 0.00000000001 for i in range(period)]),
        'yearly_unit_cost':sum([opex[i].x for i in range(period)])/sum([p_load[i]+p_cur[i].x+g_load[i]+q_load[i]+water_load[i] for i in range(period)]),

        # 'op_save_part':sum([opex[i].x for i in range(int(j/24)*24,int(j/24)*24+24)])/sum([opex_without_opt_ele[i] for i in range(int(j/24)*24,int(j/24)*24+24)]),

        'dayly_opex':[sum(opex[i].x for i in range(int(j/24)*24,int(j/24)*24+24)) for j in range(period)], #日累计运行费用
        'dayly_opex_without_opt':[sum(opex_without_opt_list[i] for i in range(int(j/24)*24,int(j/24)*24+24)) for j in range(period)],
        'opex_save_rate':sum([opex[i].x for i in range(period)])/sum(opex_without_opt_list),

        'month_sumup_opex':month_sumup_opex,#每月日累计运行成本
        'month_sumup_pload':month_sumup_pload,
        'month_sumup_hpur':month_sumup_hpur,
        'month_hydrogen_use':month_sumup_huse,
        'month_zero_carbon_supply':month_sumup_greenpoweruse,
        'month_sumup_pv':month_sumup_pv,
        'month_sumup_fc':month_sumup_fc,
        'month_sumup_hel':month_sumup_hel,
        'day_zero_carbon_use':[sum([p_pv[j].x + p_fc[j].x for j in range(int(i/24)*24,i+1)]) for i in range(period)], #日累计零碳能源使用量
        'day_carbon':[sum([p_pur[j].x*0.5837*carbon_fake_rate for j in range(int(i/24)*24,i)]) for i in range(period)],#[sum(p_pur[int(i/24):i])*0.5837 for i in range(period)],
        'day_carbon_reduce':[sum([p_load[i]+g_load[i]/k_eb+water_load[i]/k_eb+q_load[i]/k_hp_q for i in range(int(j/24)*24,j+1)])*0.5837*carbon_fake_rate for j in range(period)],
        'day_opex':[sum([opex[j].x  for j in range(int(i/24)*24,i+1)]) for i in range(period)],

        ### old devices
        'p_sold':[p_cur[i].x for i in range(period)],#弃电
        'p_pur':[p_pur[i].x for i in range(period)],#电网下电
        #'p_error':[p_pur[i].x - p_load[i] for i in range(period)],
        
        # 'p_sol':[p_sol[i].x for i in range(period)],#电网上电
        'p_pv':[p_pv[i].x for i in range(period)],#光伏
        'p_fc':[p_fc[i].x for i in range(period)],#燃料电池

        'p_hp':[p_hp[i].x for i in range(period)],#热泵
        'p_eb':[p_eb[i].x for i in range(period)],#电锅炉
        'p_pump':[p_pump[i].x for i in range(period)],#水泵
        'p_el':[p_el[i].x for i in range(period)],
        #'p_ac':[p_ac[i].x for i in range(period)],
        'p_stable_load':p_load,
        'p_load':[p_pv[i].x+p_fc[i].x for i in range(period)],#用电量
        #'p_slack':[p_slack[i].x for i in range(period)],
        #'p_load2':p_load2,
        #hydrogen
        'h_hst':[h_sto[i].x for i in range(period)],
        'h_el':[h_el[i].x for i in range(period)],
        'h_pur':[h_pur[i].x for i in range(period)],
        
        'h_fc':[h_fc[i].x for i in range(period)],
        # 'h_stooo':[h_stooo[i].x for i in range(period)],
        # 'h_stooooo':[48+h_sto[i].x + sum([ h_pur[j].x for j in range(24*int(i/24),24*int(i/24)+24)]) - sum([h_pur[j].x for j in range(24*int(i/24),i)])for i in range(period)],
        #'h_tube':[hydrogen_bottle_max_start - sum([h_pur[i].x for i in range(j)]) for j in range(period)],
        #thermal
        't_ht':[t_ht[i].x for i in range(period)],  
        'water_load':water_load,
        'g_load':g_load,
        'g_hp':[g_hp[i].x for i in range(period)],
        'g_eb':[g_eb[i].x for i in range(period)],
        'g_fc':[g_fc[i].x for i in range(period)],
        'g_ht':[c*m_ht*(t_ht[i].x - t_ht_l[i].x + 0.001*(t_ht_l[i].x-t_ht_min)) for i in range(period)],

        # cold
        't_ct':[t_ct[i].x for i in range(period)],
        #'q_ac':[q_ac[i].x for i in range(period)],
        'q_ct':[c*m_ct*(t_ct[i].x - t_ct_l[i].x-0.001*(t_ct_l[i].x - t_ct_min)) for i in range(period)],
        'q_hp':[q_hp[i].x for i in range(period)],
        'q_load':q_load,

        'test_hp':[q_hp[i].x*g_hp[i].x for i in range(period)],

    }


    return dict_plot




if __name__ == '__main__':
    OptimizationDay()


# period = len(g_de)
# # Create a new model
# m = gp.Model("bilinear")

# # Create variables
# ce_h = m.addVar(vtype=GRB.CONTINUOUS, lb=0, name="ce_h")

# m_ht = m.addVar(vtype=GRB.CONTINUOUS, lb=10, name="m_ht") # capacity of hot water tank

# t_ht = [m.addVar(vtype=GRB.CONTINUOUS, lb=0, name=f"t_ht{t}") for t in range(period)] # temperature of hot water tank

# t_fc = [m.addVar(vtype=GRB.CONTINUOUS, lb=0, name=f"t_fc{t}") for t in range(period)] # outlet temperature of fuel cells cooling circuits

# g_fc = [m.addVar(vtype=GRB.CONTINUOUS, lb=0, name=f"g_fc{t}") for t in range(period)] # heat generated by fuel cells

# p_fc = [m.addVar(vtype=GRB.CONTINUOUS, lb=0, name=f"p_fc{t}") for t in range(period)]

# fc_max = m.addVar(vtype=GRB.CONTINUOUS, lb=0, name="fc_max") # rated heat power of fuel cells

# el_max = m.addVar(vtype=GRB.CONTINUOUS, lb=0, name="el_max") # rated heat power of fuel cells

# t_de = [m.addVar(vtype=GRB.CONTINUOUS, lb=0,name=f"t_de{t}") for t in range(period)] # outlet temparature of heat supply circuits

# h_fc = [m.addVar(vtype=GRB.CONTINUOUS, lb=0, name=f"h_fc{t}") for t in range(period)] # hydrogen used in fuel cells

# m_fc = m.addVar(vtype=GRB.CONTINUOUS, lb=0, name=f"m_fc") # fuel cells water

# m_el = m.addVar(vtype=GRB.CONTINUOUS, lb=0, name=f"m_el") # fuel cells water


# g_el = [m.addVar(vtype=GRB.CONTINUOUS, lb=0, name=f"g_el{t}") for t in range(period)] # heat generated by Electrolyzer

# h_el = [m.addVar(vtype=GRB.CONTINUOUS, lb=0, name=f"h_el{t}") for t in range(period)] # hydrogen generated by electrolyzer

# p_el = [m.addVar(vtype=GRB.CONTINUOUS, lb=0, name=f"p_el{t}") for t in range(period)] # power consumption by electrolyzer

# t_el = [m.addVar(vtype=GRB.CONTINUOUS, lb=0, name=f"t_el{t}") for t in range(period)] # outlet temperature of electrolyzer cooling circuits

# h_sto = [m.addVar(vtype=GRB.CONTINUOUS, lb=0, name=f"h_sto{t}") for t in range(period)] # hydrogen storage

# h_pur = [m.addVar(vtype=GRB.CONTINUOUS, lb=0, name=f"h_pur{t}") for t in range(period)] # hydrogen purchase

# p_pur = [m.addVar(vtype=GRB.CONTINUOUS, lb=0, name=f"p_pur{t}") for t in range(period)] # power purchase

# p_sol = [m.addVar(vtype=GRB.CONTINUOUS, lb=0, name=f"p_sol{t}") for t in range(period)] # power purchase

# area_pv = m.addVar(vtype=GRB.CONTINUOUS, lb=0, ub = 1000, name=f"area_pv")

# p_pump = [m.addVar(vtype=GRB.CONTINUOUS, lb=0, name=f"p_pump{t}") for t in range(period)] 

# hst = m.addVar(vtype=GRB.CONTINUOUS, lb=0, ub = 1000, name=f"hst")

# #m.addConstr(m_el+m_fc <= 0.001*m_ht)
# for i in range(int(period/24)-1):
#     m.addConstr(t_ht[i*24+24] == t_ht[24*i])
# m.addConstr(t_ht[-1] == t_ht[0])
# #m.addConstr(h_sto[0] == 0)
# m.addConstr(h_sto[-1] == h_sto[0])
# for i in range(period - 1):
#     m.addConstr(m_ht * (t_ht[i + 1] - t_ht[i]) == 
#         m_fc * (t_fc[i] - t_ht[i]) + m_el * (t_el[i] - t_ht[i]) - m_de[i] * (t_ht[i] - t_de[i]))
#     m.addConstr(h_sto[i+1] - h_sto[i] == h_pur[i] + h_el[i] - h_fc[i])
    
# m.addConstr(m_ht * (t_ht[0] - t_ht[i]) == m_fc * (t_fc[i] - t_ht[i]) + m_el * (t_el[i] - t_ht[i]) - m_de[i] * (t_ht[i] - t_de[i]))
# m.addConstr(h_sto[0] - h_sto[-1] == h_pur[-1] + h_el[-1] - h_fc[-1])
# m.addConstr(t_ht[0] == 55)
# for i in range(period):
#     m.addConstr(t_de[i] >= 40)
#     m.addConstr(p_eb[i] + p_el[i] + p_sol[i] + p_pump[i] + p_load[i]== p_pur[i] + p_fc[i] + k_pv*area_pv*r[i])
#     m.addConstr(g_fc[i] <= 18 * h_fc[i])
#     m.addConstr(p_pump[i] == 3.5/1000 * (m_fc+m_de[i]+m_el))#热需求虽然低，水泵耗电高。
#     m.addConstr(p_fc[i] <= 18 * h_fc[i])#氢燃烧产电
#     m.addConstr(h_el[i] <= k_el * p_el[i])
#     m.addConstr(g_el[i] <= 0.2017*p_el[i])
#     m.addConstr(g_fc[i] == c_kWh * m_fc * (t_fc[i] - t_ht[i]))
#     m.addConstr(g_el[i] == c_kWh * m_el * (t_el[i] - t_ht[i]))
#     m.addConstr(t_fc[i] <= 75)
#     m.addConstr(t_el[i] <= 75)
#     m.addConstr(h_sto[i]<=hst)
#     m.addConstr(h_el[i]<=hst)
#     #m.addConstr(t_ht[i] >= 50)
#     m.addConstr(p_fc[i] <= fc_max)
#     m.addConstr(p_el[i] <= el_max)
#     m.addConstr(g_de[i] == c_kWh * m_de[i] * (t_ht[i] - t_de[i]))
#     #m.addConstr(m_fc <= m_ht)
# # m.addConstr(m_fc[i] == m_ht/3)
# # m.addConstr(m_ht >= 4200*100)
# # m.addConstr(t_ht[i] <= 80)#强化条件


# # m.setObjective( crf_pv * cost_pv*area_pv+ crf_el*cost_el*el_max
# #     +crf_hst * hst*cost_hst +crf_water* cost_water_hot*m_ht + crf_fc *cost_fc * fc_max + lambda_h*gp.quicksum(h_pur)*365+ 
# #     365*gp.quicksum([p_pur[i]*lambda_ele_in[i] for i in range(24)])-365*gp.quicksum(p_sol)*lambda_ele_out , GRB.MINIMIZE)
# m.setObjective( crf_pv * cost_pv*area_pv+ crf_el*cost_el*el_max
#     +crf_hst * hst*cost_hst +crf_water* cost_water_hot*m_ht + crf_fc *cost_fc * fc_max + lambda_h*gp.quicksum(h_pur)*365/7+ 
#     gp.quicksum([p_pur[i]*lambda_ele_in[i] for i in range(period)])*365/7-gp.quicksum(p_sol)*lambda_ele_out*365/7, GRB.MINIMIZE)
# #-gp.quicksum(p_sol)*lambda_ele_out 
# # First optimize() call will fail - need to set NonConvex to 2
# m.params.NonConvex = 2
# m.params.MIPGap = 0.05
# # m.optimize()
# #m.computeIIS()
# try:
#     m.optimize()
# except gp.GurobiError:
#     print("Optimize failed due to non-convexity")

# # Solve bilinear model
# # m.params.NonConvex = 2
# # m.optimize()

# #m.printAttr('x')
# m.write('sol_winter.mst')
# # Constrain 'x' to be integral and solve again
# # x.vType = GRB.INTEGER
# # m.optimize()

# # m.printAttr('x')

# wb = xlwt.Workbook()
# result = wb.add_sheet('result')
# alpha_ele = 1.01
# alpha_heat = 0.351
# ce_c = np.sum(p_load)*alpha_ele + np.sum(g_de)*alpha_heat
# #c_cer == lambda_carbon*(ce_c - ce_h)
# p_pur_tmp = m.getAttr('x', p_pur)
# p_sol_tmp = m.getAttr('x', p_sol)
# ce_h_1 = np.sum(p_pur_tmp)*alpha_ele - np.sum(p_sol_tmp)*alpha_ele


# item1 = ['m_ht','m_fc','m_el','fc_max','el_max']
# item2 = [g_el,g_fc,p_el,p_fc,p_pur,p_pump,p_sol,t_ht,t_el,h_el,h_fc,t_fc,t_de,h_sto,h_pur]
# a_pv = m.getVarByName('area_pv').getAttr('x')
# item3 = [[k_pv*a_pv*r[i] for i in range(len(r))],p_load,g_de]
# item3_name = ['p_pv','p_load','g_de']
# print(m.getAttr('x', p_el))
# for i in range(len(item1)):
#     result.write(0,i,item1[i])
#     result.write(1,i,m.getVarByName(item1[i]).getAttr('x'))
# for i in range(len(item2)):
#     tmp = m.getAttr('x', item2[i])
#     result.write(0,i+len(item1),item2[i][0].VarName[:-1])
#     for j in range(len(tmp)):
#         result.write(j+1,i+len(item1),tmp[j])

# for i in range(3):
#     tmp = item3[i]
#     result.write(0,i+len(item1)+len(item2),item3_name[i])
#     for j in range(len(tmp)):
#         result.write(j+1,i+len(item1)+len(item2),tmp[j])

# t_ht = m.getAttr('x', t_ht)
# m_ht = m.getVarByName('m_ht').getAttr('x')
# res = []
# for i in range(len(t_ht)-1):
#     res.append(c*m_ht*(t_ht[i+1] - t_ht[i])/3.6/1000000)
# res.append(c*m_ht*(t_ht[0]-t_ht[-1])/3.6/1000000)
# result.write(0,3+len(item1)+len(item2),'g_ht')
# for j in range(len(res)):
#     result.write(j+1,3+len(item1)+len(item2),res[j])
# result.write(0,4+len(item1)+len(item2),'cer')
# result.write(1,4+len(item1)+len(item2),(ce_c - ce_h_1))

# wb.save("sol_season_12day_729.xls")
# #print(m.getJSONSolution())







