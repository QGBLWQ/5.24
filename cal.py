# -*- coding: GBK -*-
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D
from tqdm import tqdm
WindProducePrice = 0.5 #���ɱ� Ԫ/kwh
LightProducePrice = 0.4 #���ɱ� Ԫ/kwh
BuyInPrice = 1 #����������ɱ� Ԫ/kwh
batteryCapacity = 100 #������� kwh
batteryPower = 50 #��ع��� kw
data1 = pd.read_excel('5.24/A�⸽��1����԰�������ո�������.xlsx')
data2 = pd.read_excel('5.24/data2.xlsx')#�������ݣ�����2
data2 = data2.iloc[30:,]
data2 = data2.reset_index(drop=True)
LightProduceByA = data2['Unnamed: 1'].values.tolist()
WinProduceByB = data2['Unnamed: 2'].values.tolist()
LightProduceByC = data2['Unnamed: 3'].values.tolist()
WinProduceByC = data2['Unnamed: 4'].values.tolist()
LoadA = data1['԰��A����(kW)'].values.tolist()
LoadB = data1['԰��B����(kW)'].values.tolist()
LoadC = data1['԰��C����(kW)'].values.tolist()
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
    # ��ʼ��ص�������Ϊ������������10%
    # ����ص�������Ϊ������������90%
    min_battery_Capacity = 0.1 * batteryCapacity
    max_battery_Capacity = 0.9 * batteryCapacity
    CurrentBatteryCapacity = 10 #��ʼ���������
    TotalData = pd.DataFrame(columns=[
        'ʱ��',
        '��ص���',
        '����-����',
        '��绨��',
        '����ɱ�',
        '�ܻ���',
        '��绨��(�޵��)',
        '����ɱ�(�޵��)',
        '�ܻ���(�޵��)'
        ])
    CheckData = pd.DataFrame(columns=[
    '԰��������',
    '԰��������',
    '԰������',
    '԰���ܹ���ɱ�',
    '԰��ƽ������ɱ�'
    ])
    FirstFlag = True
    iteration = 0
    totalcost = 0
    temptotal = 0
    for PowerProduce,Load in zip(Produce,Loads):#PowerProduceΪ��������LoadΪ������
        # Check_row = pd.DataFrame({
        #'԰��������':[0],
        #'԰��������':[0],
        #'԰������':[Load],
        #'԰���ܹ���ɱ�':[0],
        #'԰��ƽ������ɱ�':[0]
        #})
        Demand = Load-PowerProduce #�����
        BuyInElcCost = 0 #��ʼ���������
        ProduceElcCost = 0 #��ʼ���������
        # �жϸ������Ƿ���ڷ�����
        if Demand>0:#��Ҫ�ŵ�
            if CurrentBatteryCapacity>min_battery_Capacity:#�жϵ���Ƿ���Էŵ�
                CurrentBatteryAvalibleCapacity = min(batteryPower,CurrentBatteryCapacity-min_battery_Capacity) #��ǰ���õ���
                if(Demand-(CurrentBatteryAvalibleCapacity)*0.95>0):#�ж��Ƿ���Ҫ�������
                    #Check_row['԰��������'] = (Demand-(CurrentBatteryAvalibleCapacity)*0.95)
                    BuyInElcCost = (Demand-(CurrentBatteryAvalibleCapacity)*0.95)/Price
                    ProduceElcCost = (PowerProduce*TypedPrice)
                    CurrentBatteryCapacity-=CurrentBatteryAvalibleCapacity
                else:
                    ProduceElcCost = (PowerProduce*TypedPrice)
                    CurrentBatteryCapacity = CurrentBatteryCapacity - CurrentBatteryAvalibleCapacity
            else:
                #Check_row['԰��������'] = Demand
                BuyInElcCost = Demand/Price
                ProduceElcCost = (PowerProduce*TypedPrice)
        else:#���Գ��
            if (CurrentBatteryCapacity<max_battery_Capacity):#�жϵ���Ƿ���Գ��
                CurrentBatteryAvalibleCapacity = min(batteryPower,max_battery_Capacity-CurrentBatteryCapacity) #��ǰ�ɳ����
                if(abs(Demand)-(CurrentBatteryAvalibleCapacity)/0.95>0):#�ж��Ƿ�����
                    #Check_row['԰��������'] = abs(Demand)-(CurrentBatteryAvalibleCapacity)/0.95
                    ProduceElcCost = (PowerProduce*TypedPrice)
                    CurrentBatteryCapacity += CurrentBatteryAvalibleCapacity
                else:
                    ProduceElcCost = (PowerProduce*TypedPrice)
                    CurrentBatteryCapacity = CurrentBatteryCapacity + abs(CurrentBatteryAvalibleCapacity)*0.95
            else:
                #Check_row['԰��������'] = abs(Demand)
                ProduceElcCost = (PowerProduce*TypedPrice)
        totalcost = BuyInElcCost+ProduceElcCost
        #Check_row['԰���ܹ���ɱ�'] = totalcost
        temptotal +=totalcost
         # ��¼ÿ��ѭ��������
        new_row = pd.DataFrame({
        'ʱ��':[iteration],
        '��ص���':[CurrentBatteryCapacity],
        '����-����':[Demand],
        '��绨��':[BuyInElcCost],
        '����ɱ�':[ProduceElcCost],
        '�ܻ���':[totalcost],
        '��绨��(�޵��)':[NoBattery[0][iteration]],
        '����ɱ�(�޵��)':[NoBattery[1][iteration]],
        '�ܻ���(�޵��)':[NoBattery[2][iteration]]
        })
        TotalData = pd.concat([TotalData,new_row], ignore_index=False)
        #CheckData = pd.concat([CheckData,Check_row], ignore_index=False)
        iteration +=1
    CheckData['԰��ƽ������ɱ�'] = temptotal/np.sum(np.array(Loads))
    return temptotal


def calculate_costs_for_multy(LightProduce, WindProduce, Loads, Price, LightTypedPrice, WindTypedPrice, NoBattery):
    # ��ʼ��ص�������Ϊ������������10%
    # ����ص�������Ϊ������������90%
    min_battery_Capacity = 0.1 * batteryCapacity
    max_battery_Capacity = 0.9 * batteryCapacity
    CurrentBatteryCapacity = 10  # ��ʼ���������
    TotalData = pd.DataFrame(columns=[
        'ʱ��',
        '��ص���',
        '����-����',
        '��绨��',
        '����ɱ�',
        '�ܻ���',
        '��绨��(�޵��)',
        '����ɱ�(�޵��)',
        '�ܻ���(�޵��)'
    ])
    CheckData = pd.DataFrame(columns=[
        '԰��������',
        '԰��������',
        '԰������',
        '԰���ܹ���ɱ�',
        '԰��ƽ������ɱ�'
    ])
    FirstFlag = True
    iteration = 0
    totalcost = 0
    temptotal = 0
    for LightPowerProduce, WindPowerProduce, Load in zip(LightProduce, WindProduce, Loads):  # PowerProduceΪ��������LoadΪ������
        # Check_row = pd.DataFrame({
        # '԰��������':[0],
        # '԰��������':[0],
        # '԰������':[Load],
        # '԰���ܹ���ɱ�':[0],
        # '԰��ƽ������ɱ�':[0]
        # })
        PowerProduce = (LightPowerProduce + WindPowerProduce)
        Demand = Load - (LightPowerProduce + WindPowerProduce)  # �����
        BuyInElcCost = 0  # ��ʼ���������
        ProduceElcCost = 0  # ��ʼ���������
        # �жϸ������Ƿ���ڷ�����
        if Demand > 0:  # ��Ҫ�ŵ�
            if CurrentBatteryCapacity > min_battery_Capacity:  # �жϵ���Ƿ���Էŵ�
                CurrentBatteryAvalibleCapacity = min(batteryPower,
                                                     CurrentBatteryCapacity - min_battery_Capacity)  # ��ǰ���õ���
                if (Demand - (CurrentBatteryAvalibleCapacity) * 0.95 > 0):  # �ж��Ƿ���Ҫ�������
                    #Check_row['԰��������'] = (Demand - (CurrentBatteryAvalibleCapacity) * 0.95)
                    BuyInElcCost = (Demand - (CurrentBatteryAvalibleCapacity) * 0.95) / Price
                    ProduceElcCost = (LightPowerProduce * LightTypedPrice) + (WindPowerProduce * WindTypedPrice)
                    CurrentBatteryCapacity -= CurrentBatteryAvalibleCapacity
                else:
                    ProduceElcCost = (LightPowerProduce * LightTypedPrice) + (WindPowerProduce * WindTypedPrice)
                    CurrentBatteryCapacity = CurrentBatteryCapacity - CurrentBatteryAvalibleCapacity
            else:
                #Check_row['԰��������'] = Demand
                BuyInElcCost = Demand / Price
                ProduceElcCost = (LightPowerProduce * LightTypedPrice) + (WindPowerProduce * WindTypedPrice)
        else:  # ���Գ��
            if (CurrentBatteryCapacity < max_battery_Capacity):  # �жϵ���Ƿ���Գ��
                CurrentBatteryAvalibleCapacity = min(batteryPower,
                                                     max_battery_Capacity - CurrentBatteryCapacity)  # ��ǰ�ɳ����
                if (abs(Demand) - (CurrentBatteryAvalibleCapacity) / 0.95 > 0):  # �ж��Ƿ�����
                    #Check_row['԰��������'] = abs(Demand) - (CurrentBatteryAvalibleCapacity) / 0.95
                    ProduceElcCost = (LightPowerProduce * LightTypedPrice) + (WindPowerProduce * WindTypedPrice)
                    CurrentBatteryCapacity += CurrentBatteryAvalibleCapacity
                else:
                    ProduceElcCost = (LightPowerProduce * LightTypedPrice) + (WindPowerProduce * WindTypedPrice)
                    CurrentBatteryCapacity = CurrentBatteryCapacity + abs(CurrentBatteryAvalibleCapacity) * 0.95
            else:
                #Check_row['԰��������'] = abs(Demand)
                ProduceElcCost = (LightPowerProduce * LightTypedPrice) + (WindPowerProduce * WindTypedPrice)
        totalcost = BuyInElcCost + ProduceElcCost
        #Check_row['԰���ܹ���ɱ�'] = totalcost
        temptotal += totalcost
        # ��¼ÿ��ѭ��������
        new_row = pd.DataFrame({
            'ʱ��': [iteration],
            '��ص���': [CurrentBatteryCapacity],
            '����-����': [Demand],
            '��绨��': [BuyInElcCost],
            '����ɱ�': [ProduceElcCost],
            '�ܻ���': [totalcost],
            '��绨��(�޵��)': [NoBattery[0][iteration]],
            '����ɱ�(�޵��)': [NoBattery[1][iteration]],
            '�ܻ���(�޵��)': [NoBattery[2][iteration]]
        })
        TotalData = pd.concat([TotalData, new_row], ignore_index=False)
        #CheckData = pd.concat([CheckData, Check_row], ignore_index=False)
        iteration += 1
    CheckData['԰��ƽ������ɱ�'] = temptotal / np.sum(np.array(Loads))

    return temptotal


import numpy as np
from tqdm import tqdm
from concurrent.futures import ProcessPoolExecutor, as_completed
import logging

# ������־��¼
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# ���� calculate_costs �� calculate_costs_for_multy �����Ѷ���

def calculate_total_cost(batteryPower, batteryCapacity):
    try:
        SetBatteryCostPer = (batteryCapacity * 1800 + batteryPower * 800) / (10 * 365)  # ��ʮ���̯��ط��õ�ÿ��
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
        return (batteryPower, batteryCapacity, np.nan)  # ���� NaN ��ʾ����ʧ��

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
    with ProcessPoolExecutor(max_workers=8) as executor:  # ���Ʋ��ж�
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

    # ʹ�� TotalCost_Q_1_3 ����һ������
    return TotalCost_Q_1_3

if __name__ == "__main__":
    TotalCost_Q_1_3 = main()
