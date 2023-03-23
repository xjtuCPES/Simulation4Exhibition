
[![forthebadge](https://forthebadge.com/images/badges/made-with-python.svg)](https://forthebadge.com)


[![forthebadge](https://forthebadge.com/images/badges/built-with-love.svg)](https://forthebadge.com)

[![forthebadge](Img/powered-by-guo.svg)](https://github.com/gwyxjtu)

# **设备运行决策**
- 代码负责人：郭王懿
- 主要内容：分为日前决策和日内决策，其中日前决策根据后一天24h的负荷预测和光伏预测，计算出下一日的基策略；日内决策根据当前时刻的储能状态和未来四小时的负荷预测和当日的基策略，做出4小时的决策。
目前还可以计算全年仿真数据，用于仿真展示。

## **0. 代码功能解释**
本代码负责优化决策部分的设备运行决策。日前决策在每日0时运行，根据下一日的光照强度和负荷预测水平，决策得到下一日设备运行策略；日内决策在日内每小时运行，根据未来四小时负荷和光照，决策得到4小时的设备运行策略。


## **1. 代码环境**
 
```
pandas==1.4.2
gurobipy==9.5.1
matplotlib==3.3.4
numpy==1.22.3
pandas==1.4.2
PuLP==2.6.0
requests==2.25.1
xlrd==1.2.0
xlwt==1.3.0
```

- 搭环境的话请按照目录下的requirements.txt 进行安装
```pip install -r requirements.txt```
## **2. demo示例**
- 注：运行代码如下   
``` python optimization_year_planning.py```   
- 运行结果代码段如下：
```
% python optimization_year_planning.py

1233840.5249999547 381650.51368328545 389943.0153125662
Set parameter Username
Academic license - for non-commercial use only - expires 2023-03-04
Set parameter MIPGap to value 0.01
Set parameter LogFile to value "testlog.log"
Set parameter NonConvex to value 2
Gurobi Optimizer version 9.5.1 build v9.5.1rc2 (mac64[arm])
Thread count: 8 physical cores, 8 logical processors, using up to 8 threads
Optimize a model with 280336 rows, 236529 columns and 670241 nonzeros
Model fingerprint: 0xaf4ef3b9
Coefficient statistics:
  Matrix range     [2e-04, 1e+03]
  Objective range  [6e-01, 1e+07]
  Bounds range     [0e+00, 0e+00]
  RHS range        [9e-03, 5e+05]

Concurrent LP optimizer: primal simplex, dual simplex, and barrier
Showing barrier log only...

Presolve removed 227776 rows and 131409 columns
Presolve time: 0.31s
Presolved: 52560 rows, 105120 columns, 192720 nonzeros

Ordering time: 0.01s

Barrier statistics:
 AA' NZ     : 9.636e+04
 Factor NZ  : 6.568e+05 (roughly 70 MB of memory)
 Factor Ops : 8.835e+06 (less than 1 second per iteration)
 Threads    : 6

                  Objective                Residual
Iter       Primal          Dual         Primal    Dual     Compl     Time
   0   1.39462013e+14 -1.26325482e+14  2.44e+02 4.48e+00  3.39e+09     1s
   1   1.41185327e+13 -3.63655869e+13  2.35e+01 2.98e-08  4.09e+08     1s

Barrier performed 1 iterations in 0.78 seconds (0.38 work units)
Barrier solve interrupted - model solved by another algorithm


Solved with primal simplex
Solved in 26013 iterations and 0.78 seconds (1.48 work units)
Optimal objective  1.961087156e+06
F2_p_load
F2_q_load
F2_g_load
F1_p_load
F1_q_load
F1_g_load
B1_p_load
B1_q_load
B1_g_load
17_p_load
17_q_load
17_g_load
18_p_load
18_q_load
18_g_load
p_load
p_fc
p_pv
p_pur
g_hp
g_eb
g_fc
g_ht
q_hp
q_ct
month_hydrogen_use
month_zero_carbon_supply
env_temp
env_humidity
env_co2
wind_speed
therm_temp
air_temp
air_humidity
air_pm25
air_pm10
weather_solar
weather_wind
weather_rain
weather_pre
weather_direction
meter_power
meter_v
meter_a
p_el
p_fc
p_hp
p_eb
h_hst
t_ht
t_ct
p_el
p_fc
p_hp
p_eb
h_hst
t_ht
t_ct
24h_opex_with_opt
24h_opex_without_opt
30d_opex_with_opt
30d_opex_without_opt
30d_opex_sum
h_el
h_pur
p_pur
p_pv
p_fc
g_hp
g_eb
g_fc
g_ht
q_hp
q_ct
sum_p_fc
sum_p_pv
sum_h_el
7d_efficiency
7d_area_energy_use
24h_efficiency
24h_area_energy_use
```   
## **3.输入解释**

- 设备配置文件在`Config`文件夹下的`config.json`，由管理员进行配置。
- 模型的输入在`Input`文件夹下，主要由三部分组成

### **load.xlsx**：主要负责日前决策程序的输入

			
+ 日期:当前日期
+ 冷负荷(kW):8760h冷负荷
+ 供暖热负荷(kW):8760小时每个小时的热负荷
+ 电负荷/kwh:8760小时每个小时的电负荷
  
### **solar.csv**：主要负责日内决策程序的输入

+ ele_load:未来4h每个小时的电负荷
+ g_load:未来4h每个小时的热负荷
+ q_load:未来4h每个小时的冷负荷
+ soalr:未来4h每个小时的光照强度
+ time:当前时间，例如：下午13点，填入13



### **dict_sensor.xlsx**
		

+ time:世界时间
+ local_time:当地时间
+ electricity:光照强度



## **4.输出格式**
模型运行生成1个excel：   
### **2022-07-22 10-13-14_dict_plot_planning.xls**
详见文档
共有下述sheet
### sheet1:dict_energynet

+ F2_p_load:边缘节点 能源站F2的24h电负荷
+ F2_q_load:边缘节点 能源站F2的24h冷负荷
+ F2_g_load:边缘节点 能源站F2的24h热负荷
+ F1_p_load:边缘节点 能源站F1的24h电负荷
+ F1_q_load:边缘节点 能源站F1的24h冷负荷
+ F1_g_load:边缘节点 能源站F1的24h热负荷
+ B1_p_load:边缘节点 能源站B1的24h电负荷
+ B1_q_load:边缘节点 能源站B1的24h冷负荷
+ B1_g_load:边缘节点 能源站B1的24h热负荷
+ 17_p_load:边缘节点 DK2-17的24h电负荷
+ 17_q_load:边缘节点 DK2-17的24h冷负荷
+ 17_g_load:边缘节点 DK2-17的24h热负荷
+ 18_p_load:边缘节点 DK2-18的24h电负荷
+ 18_q_load:边缘节点 DK2-18的24h冷负荷
+ 18_g_load:边缘节点 DK2-18的24h热负荷
+ p_load:24h电负荷
+ p_fc:24h燃料电池分时电功率
+ p_pv:24h光伏分时供电量
+ p_pur:24h电量购买量
+ g_hp:24h热泵分时供热量
+ g_eb:24h电锅炉分时供电量
+ g_fc:24h燃料电池分时供热量
+ g_ht:24h分时热水箱供热量
+ q_hp:24h热泵分时供冷量
+ q_ct:24h蓄能水箱分时供冷量
+ month_hydrogen_use:12个月分月用氢量
+ month_zero_carbon_supply:12个月分月零碳电能供应量


### sheet2:dict_sensor

+ env_temp:环境温度
+ env_humidity:环境湿度
+ env_co2:环境CO2浓度
+ wind_speed:风速
+ therm_temp:热电偶检测器温度
+ air_temp:室外空气质量监测温度
+ air_humidity:室外空气质量监测湿度
+ air_pm25:室外空气质量监测PM2.5浓度
+ air_pm10:室外空气质量监测PM10浓度
+ weather_solar:室外气象监测光照强度
+ weather_wind:室外气象监测风速
+ weather_rain:室外气象监测降雨量
+ weather_pre:室外气象监测大气压
+ weather_direction:室外气象监测风向
+ meter_power:智能电表电量
+ meter_v:智能电表电压
+ meter_a:智能电表电流

### sheet3:dict_plot_24h
+ p_el:电解槽耗电的日前调度策略
+ p_fc:燃料电池产电的日前调度策略
+ p_hp:热泵耗电的日前调度策略
+ p_eb:电锅炉耗电的日前调度策略
+ h_hst:储氢罐储氢量的日前调度策略
+ t_ht:蓄能水箱储水温度的日前调度策略
+ t_ct:消防水池储能温度的日前调度策略

### sheet3:dict_plot_4h

+ p_el:电解槽耗电的日内调度策略
+ p_fc:燃料电池产电的日内调度策略
+ p_hp:热泵耗电的日内调度策略
+ p_eb:电锅炉耗电的日内调度策略
+ h_hst:储氢罐储氢量的日内调度策略
+ t_ht:蓄能水箱储水温度的日内调度策略
+ t_ct:消防水池储能温度的日内调度策略

值得注意的是，日内策略只展示到当日的当前时刻，比如现在是下午三点，日内策略只展示到下午三点。
24h_opex_with_opt	24h_opex_without_opt	30d_opex_with_opt	30d_opex_without_opt	30d_opex_sum
### sheet4:dict_opex

+ 24h_opex_with_opt:24h一日经优化的运行成本
+ 24h_opex_without_opt:24h一日未经优化的运行成本
+ 30d_opex_with_opt:往前推30天经优化的运行成本
+ 30d_opex_without_opt:往前推30天未经优化的运行成本


### sheet5:dict_statistic

+ h_el:电解槽耗电量
+ h_pur:市场购氢量
+ p_pur:电网购电量
+ p_pv:光伏产电量
+ p_fc:燃料电池产电量
+ g_hp:热泵供热量
+ g_eb:电锅炉供热量
+ g_fc:燃料电池供热量
+ g_ht:蓄能水箱供热量
+ q_hp:热泵供冷量
+ q_ct:蓄能水箱供冷量
+ sum_p_fc:累计燃料电池产电量
+ sum_p_pv:累计光伏产电量
+ sum_h_el:累计电解槽制氢量
+ 7d_efficiency:近7天经优化的综合能效
+ 7d_area_energy_use:近7天经优化的单位面积综合能耗
+ 24h_efficiency:近24h经优化的综合能效
+ 24h_area_energy_use:近24h经优化的单位面积综合能耗


## **5. 代码目录**

```
optimization
├─ Config
│  └─ config.json
├─ Input
│  ├─ generate_input_data.py
│  ├─ input_24h.xls
│  ├─ input_4h.xls
│  └─ input_now.xls
├─ Model
│  ├─ __pycache__
│  │  └─ optimization_day.cpython-310.pyc
│  └─ optimization_day.py
├─ Output
│  ├─ .DS_Store
│  ├─ control_output_24.xls
│  ├─ control_output_4.xls
│  ├─ dict_control_24h.xls
│  ├─ dict_control_4h.xls
│  ├─ dict_plot_24h.xls
│  ├─ dict_plot_4h.xls
│  └─ generate_out_data.py
├─ config.py
├─ cpeslog
│  ├─ __init__.py
│  ├─ __pycache__
│  │  ├─ __init__.cpython-310.pyc
│  │  ├─ __init__.cpython-37.pyc
│  │  ├─ log_code.cpython-310.pyc
│  │  └─ log_code.cpython-37.pyc
│  └─ log_code.py
├─ model.py
├─ my_test_log
│  ├─ _error.log
│  ├─ _info.log
│  ├─ _info.log.1
│  └─ _info.log.2
├─ optimization_24h.py
├─ optimization_4h.py
├─ readme.md
└─ requirements.txt

```

## **6. 可能存在的问题**

### 模型不可行的解决办法
如果无论是日前或者日内，出现模型不可行，或者模型结果计算错误。如何继续控制

### 末端时刻储能的调配
末端时刻储能调配很重要，调配好可以较大程度解决弃弃光。
