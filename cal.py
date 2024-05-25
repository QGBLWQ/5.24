# -*- coding: GBK -*-
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D
from tqdm import tqdm
WindProducePrice = 0.5 #风电成本 元/kwh
LightProducePrice = 0.4 #光电成本 元/kwh
BuyInPrice = 1 #主电网购入成本 元/kwh
batteryCapacity = 100 #电池容量 kwh
batteryPower = 50 #电池功率 kw
data1 = pd.read_excel('5.24/A题附件1：各园区典型日负荷数据.xlsx')
data2 = pd.read_excel('5.24/data2.xlsx')#读入数据，附件2
data2 = data2.iloc[30:,]
data2 = data2.reset_index(drop=True)
LightProduceByA = data2['Unnamed: 1'].values.tolist()
WinProduceByB = data2['Unnamed: 2'].values.tolist()
LightProduceByC = data2['Unnamed: 3'].values.tolist()
WinProduceByC = data2['Unnamed: 4'].values.tolist()
LoadA = data1['园区A负荷(kW)'].values.tolist()
LoadB = data1['园区B负荷(kW)'].values.tolist()
LoadC = data1['园区C负荷(kW)'].values.tolist()
EA_buy = np.maximum((np.array(LoadA) - np.array(LightProduceByA)), 0) #shape(24,)
EA_curt = np.maximum((np.array(LightProduceByA) - np.array(LoadA)), 0)
A_LightProduce = np.array(LightProduceByA)

EB_buy = np.maximum((np.array(LoadB) - WinProduceByB), 0)
EB_curt = np.maximum((np.array(WinProduceByB) - np.array(LoadB)), 0)
B_WindProduce = np.array(WinProduceByB)

EC_buy = np.maximum((np.array(LoadC) - np.array(LightProduceByC) - np.array(WinProduceByC)), 0)
EC_curt = np.maximum((np.array(LightProduceByC) + np.array(WinProduceByC) - np.array(LoadC)), 0)
C_LightProduce = np.array(LightProduceByC)
C_WindProduce = np.array(WinProduceByC)

def Compute_Single_Cost(E_buy, E_curt, LightProduce, LightPrice, WindProduce, WindPrice):
    #array : E_buy, E_curt
    #const : LPro, LPri, WPro, WPri
    return E_buy * BuyInPrice + WindProduce * WindPrice + LightProduce * LightPrice

ACost = Compute_Single_Cost(EA_buy, EA_curt, A_LightProduce, LightProducePrice, 0, WindProducePrice)
BCost = Compute_Single_Cost(EB_buy, EB_curt, 0, A_LightProduce, B_WindProduce, WindProducePrice)
CCost = Compute_Single_Cost(EC_buy, EC_curt, C_LightProduce, LightProducePrice, C_WindProduce, WindProducePrice)

APer = ACost.sum() / np.sum(LoadA)
BPer = BCost.sum() / np.sum(LoadB)
CPer = CCost.sum() / np.sum(LoadC)

APer, BPer, CPer

def calculate_costs(Produce,Loads,Price,TypedPrice,NoBattery):
    # 初始电池电量设置为电池最大容量的10%
    # 最大电池电量设置为电池最大容量的90%
    min_battery_Capacity = 0.1 * batteryCapacity
    max_battery_Capacity = 0.9 * batteryCapacity
    CurrentBatteryCapacity = 10 #初始化电池容量
    TotalData = pd.DataFrame(columns=[
        '时间',
        '电池电量',
        '负荷-发电',
        '买电花费',
        '产电成本',
        '总花费',
        '买电花费(无电池)',
        '产电成本(无电池)',
        '总花费(无电池)'
        ])
    CheckData = pd.DataFrame(columns=[
    '园区购电量',
    '园区弃电量',
    '园区负荷',
    '园区总供电成本',
    '园区平均供电成本'
    ])
    FirstFlag = True
    iteration = 0
    totalcost = 0
    temptotal = 0
    for PowerProduce,Load in zip(Produce,Loads):#PowerProduce为发电量，Load为负荷量
        # Check_row = pd.DataFrame({
        #'园区购电量':[0],
        #'园区弃电量':[0],
        #'园区负荷':[Load],
        #'园区总供电成本':[0],
        #'园区平均供电成本':[0]
        #})
        Demand = Load-PowerProduce #需电量
        BuyInElcCost = 0 #初始化购电费用
        ProduceElcCost = 0 #初始化发电费用
        # 判断负荷量是否大于发电量
        if Demand>0:#需要放电
            if CurrentBatteryCapacity>min_battery_Capacity:#判断电池是否可以放电
                CurrentBatteryAvalibleCapacity = min(batteryPower,CurrentBatteryCapacity-min_battery_Capacity) #当前可用电量
                if(Demand-(CurrentBatteryAvalibleCapacity)*0.95>0):#判断是否需要额外买电
                    #Check_row['园区购电量'] = (Demand-(CurrentBatteryAvalibleCapacity)*0.95)
                    BuyInElcCost = (Demand-(CurrentBatteryAvalibleCapacity)*0.95)/Price
                    ProduceElcCost = (PowerProduce*TypedPrice)
                    CurrentBatteryCapacity-=CurrentBatteryAvalibleCapacity
                else:
                    ProduceElcCost = (PowerProduce*TypedPrice)
                    CurrentBatteryCapacity = CurrentBatteryCapacity - CurrentBatteryAvalibleCapacity
            else:
                #Check_row['园区购电量'] = Demand
                BuyInElcCost = Demand/Price
                ProduceElcCost = (PowerProduce*TypedPrice)
        else:#可以充电
            if (CurrentBatteryCapacity<max_battery_Capacity):#判断电池是否可以充电
                CurrentBatteryAvalibleCapacity = min(batteryPower,max_battery_Capacity-CurrentBatteryCapacity) #当前可充电量
                if(abs(Demand)-(CurrentBatteryAvalibleCapacity)/0.95>0):#判断是否弃电
                    #Check_row['园区弃电量'] = abs(Demand)-(CurrentBatteryAvalibleCapacity)/0.95
                    ProduceElcCost = (PowerProduce*TypedPrice)
                    CurrentBatteryCapacity += CurrentBatteryAvalibleCapacity
                else:
                    ProduceElcCost = (PowerProduce*TypedPrice)
                    CurrentBatteryCapacity = CurrentBatteryCapacity + abs(CurrentBatteryAvalibleCapacity)*0.95
            else:
                #Check_row['园区弃电量'] = abs(Demand)
                ProduceElcCost = (PowerProduce*TypedPrice)
        totalcost = BuyInElcCost+ProduceElcCost
        #Check_row['园区总供电成本'] = totalcost
        temptotal +=totalcost
         # 记录每次循环的数据
        new_row = pd.DataFrame({
        '时间':[iteration],
        '电池电量':[CurrentBatteryCapacity],
        '负荷-发电':[Demand],
        '买电花费':[BuyInElcCost],
        '产电成本':[ProduceElcCost],
        '总花费':[totalcost],
        '买电花费(无电池)':[NoBattery[0][iteration]],
        '产电成本(无电池)':[NoBattery[1][iteration]],
        '总花费(无电池)':[NoBattery[2][iteration]]
        })
        TotalData = pd.concat([TotalData,new_row], ignore_index=False)
        #CheckData = pd.concat([CheckData,Check_row], ignore_index=False)
        iteration +=1
    CheckData['园区平均供电成本'] = temptotal/np.sum(np.array(Loads))
    return temptotal


def calculate_costs_for_multy(LightProduce, WindProduce, Loads, Price, LightTypedPrice, WindTypedPrice, NoBattery):
    # 初始电池电量设置为电池最大容量的10%
    # 最大电池电量设置为电池最大容量的90%
    min_battery_Capacity = 0.1 * batteryCapacity
    max_battery_Capacity = 0.9 * batteryCapacity
    CurrentBatteryCapacity = 10  # 初始化电池容量
    TotalData = pd.DataFrame(columns=[
        '时间',
        '电池电量',
        '负荷-发电',
        '买电花费',
        '产电成本',
        '总花费',
        '买电花费(无电池)',
        '产电成本(无电池)',
        '总花费(无电池)'
    ])
    CheckData = pd.DataFrame(columns=[
        '园区购电量',
        '园区弃电量',
        '园区负荷',
        '园区总供电成本',
        '园区平均供电成本'
    ])
    FirstFlag = True
    iteration = 0
    totalcost = 0
    temptotal = 0
    for LightPowerProduce, WindPowerProduce, Load in zip(LightProduce, WindProduce, Loads):  # PowerProduce为发电量，Load为负荷量
        # Check_row = pd.DataFrame({
        # '园区购电量':[0],
        # '园区弃电量':[0],
        # '园区负荷':[Load],
        # '园区总供电成本':[0],
        # '园区平均供电成本':[0]
        # })
        PowerProduce = (LightPowerProduce + WindPowerProduce)
        Demand = Load - (LightPowerProduce + WindPowerProduce)  # 需电量
        BuyInElcCost = 0  # 初始化购电费用
        ProduceElcCost = 0  # 初始化发电费用
        # 判断负荷量是否大于发电量
        if Demand > 0:  # 需要放电
            if CurrentBatteryCapacity > min_battery_Capacity:  # 判断电池是否可以放电
                CurrentBatteryAvalibleCapacity = min(batteryPower,
                                                     CurrentBatteryCapacity - min_battery_Capacity)  # 当前可用电量
                if (Demand - (CurrentBatteryAvalibleCapacity) * 0.95 > 0):  # 判断是否需要额外买电
                    #Check_row['园区购电量'] = (Demand - (CurrentBatteryAvalibleCapacity) * 0.95)
                    BuyInElcCost = (Demand - (CurrentBatteryAvalibleCapacity) * 0.95) / Price
                    ProduceElcCost = (LightPowerProduce * LightTypedPrice) + (WindPowerProduce * WindTypedPrice)
                    CurrentBatteryCapacity -= CurrentBatteryAvalibleCapacity
                else:
                    ProduceElcCost = (LightPowerProduce * LightTypedPrice) + (WindPowerProduce * WindTypedPrice)
                    CurrentBatteryCapacity = CurrentBatteryCapacity - CurrentBatteryAvalibleCapacity
            else:
                #Check_row['园区购电量'] = Demand
                BuyInElcCost = Demand / Price
                ProduceElcCost = (LightPowerProduce * LightTypedPrice) + (WindPowerProduce * WindTypedPrice)
        else:  # 可以充电
            if (CurrentBatteryCapacity < max_battery_Capacity):  # 判断电池是否可以充电
                CurrentBatteryAvalibleCapacity = min(batteryPower,
                                                     max_battery_Capacity - CurrentBatteryCapacity)  # 当前可充电量
                if (abs(Demand) - (CurrentBatteryAvalibleCapacity) / 0.95 > 0):  # 判断是否弃电
                    #Check_row['园区弃电量'] = abs(Demand) - (CurrentBatteryAvalibleCapacity) / 0.95
                    ProduceElcCost = (LightPowerProduce * LightTypedPrice) + (WindPowerProduce * WindTypedPrice)
                    CurrentBatteryCapacity += CurrentBatteryAvalibleCapacity
                else:
                    ProduceElcCost = (LightPowerProduce * LightTypedPrice) + (WindPowerProduce * WindTypedPrice)
                    CurrentBatteryCapacity = CurrentBatteryCapacity + abs(CurrentBatteryAvalibleCapacity) * 0.95
            else:
                #Check_row['园区弃电量'] = abs(Demand)
                ProduceElcCost = (LightPowerProduce * LightTypedPrice) + (WindPowerProduce * WindTypedPrice)
        totalcost = BuyInElcCost + ProduceElcCost
        #Check_row['园区总供电成本'] = totalcost
        temptotal += totalcost
        # 记录每次循环的数据
        new_row = pd.DataFrame({
            '时间': [iteration],
            '电池电量': [CurrentBatteryCapacity],
            '负荷-发电': [Demand],
            '买电花费': [BuyInElcCost],
            '产电成本': [ProduceElcCost],
            '总花费': [totalcost],
            '买电花费(无电池)': [NoBattery[0][iteration]],
            '产电成本(无电池)': [NoBattery[1][iteration]],
            '总花费(无电池)': [NoBattery[2][iteration]]
        })
        TotalData = pd.concat([TotalData, new_row], ignore_index=False)
        #CheckData = pd.concat([CheckData, Check_row], ignore_index=False)
        iteration += 1
    CheckData['园区平均供电成本'] = temptotal / np.sum(np.array(Loads))

    return temptotal


import numpy as np
from tqdm import tqdm
from concurrent.futures import ProcessPoolExecutor, as_completed
import logging

# 设置日志记录
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# 假设 calculate_costs 和 calculate_costs_for_multy 函数已定义

def calculate_total_cost(batteryPower, batteryCapacity):
    try:
        SetBatteryCostPer = (batteryCapacity * 1800 + batteryPower * 800) / (10 * 365)  # 分十年均摊电池费用到每天
        logging.info(f"Calculating cost for power {batteryPower} and capacity {batteryCapacity}")
        total_cost = (calculate_costs(LightProduceByA, LoadA, BuyInPrice, LightProducePrice, [EA_buy * BuyInPrice,
                                                                                             A_LightProduce * LightProducePrice,
                                                                                             ACost]) +
                      calculate_costs(WinProduceByB, LoadB, BuyInPrice, WindProducePrice, [EA_buy * BuyInPrice,
                                                                                           A_LightProduce * LightProducePrice,
                                                                                           ACost]) +
                      calculate_costs_for_multy(LightProduceByC, WinProduceByC, LoadC, BuyInPrice, LightProducePrice, WindProducePrice, [EA_buy * BuyInPrice,
                                                                                                                                   A_LightProduce * LightProducePrice,
                                                                                                                                   ACost]) +
                      SetBatteryCostPer)
        logging.info(f"Finished calculating cost for power {batteryPower} and capacity {batteryCapacity}")
        return (batteryPower, batteryCapacity, total_cost)
    except Exception as e:
        logging.error(f"Error calculating cost for power {batteryPower} and capacity {batteryCapacity}: {e}")
        return (batteryPower, batteryCapacity, np.nan)  # 返回 NaN 表示计算失败

def main():
    st_c = 80
    end_c = 250
    step_c = 1
    st_p = 50
    end_p = 200
    step_p = 1

    CapacityRange = np.arange(st_c, end_c, step_c)
    PowerRange = np.arange(st_p, end_p, step_p)
    TotalCost_Q_1_3 = np.zeros((len(PowerRange), len(CapacityRange)))

    tasks = []
    with ProcessPoolExecutor(max_workers=8) as executor:  # 限制并行度
        for batteryPower in tqdm(PowerRange, desc="Power"):
            for batteryCapacity in CapacityRange:
                tasks.append(executor.submit(calculate_total_cost, batteryPower, batteryCapacity))

        for future in tqdm(as_completed(tasks), total=len(tasks), desc="Total Progress"):
            try:
                batteryPower, batteryCapacity, total_cost = future.result()
                x = np.where(PowerRange == batteryPower)[0][0]
                y = np.where(CapacityRange == batteryCapacity)[0][0]
                TotalCost_Q_1_3[x, y] = total_cost
            except Exception as e:
                logging.error(f"Error retrieving result: {e}")

    # 使用 TotalCost_Q_1_3 做进一步处理
    return TotalCost_Q_1_3

if __name__ == "__main__":
    TotalCost_Q_1_3 = main()
