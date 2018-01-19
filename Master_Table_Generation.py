# -*- coding: utf-8 -*-
import pandas as pd

home_dir = "E:/Belle/Staccato/"
prop = pd.read_excel(home_dir+"思加图商品属性.xlsx", encoding='gbk')
prop = prop[['商品编号', '货号', '商品款号', '三级分类', '商品销售季', '商品季', '首次上架时间', '首次上架年份', '首次上架月份']]

def GENERATE_MASTER(channel):
    trans_16 = pd.read_csv(home_dir+"ST_"+channel+"_2016.csv", encoding='gbk')
    trans_17 = pd.read_csv(home_dir+"ST_"+channel+"_2017.csv", encoding='gbk')
    trans_1617 = trans_16.append(trans_17)
    max_pricetag = trans_1617.groupby('供应商款色编号').max()['牌价']
    trans_1617['最高牌价'] = max_pricetag[trans_1617['供应商款色编号']].values
    trans_1617 = trans_1617[['日期', '供应商款色编号', '成交金额', '成交件数', '最高牌价']]
    trans_1617['折扣率'] = round(trans_1617['成交金额'] /    trans_1617['成交件数'] / trans_1617['最高牌价'],2)
    if channel == "TB":
        trans_1617['渠道'] = "淘宝"
    elif channel == "JD":
        trans_1617['渠道'] = "京东"
    elif channel == "YG":
        trans_1617['渠道'] = "优购"
    elif channel == "VIP":
        trans_1617['渠道'] = "唯品会"  
    master = pd.merge(trans_1617, prop, how='left', left_on='供应商款色编号', right_on='货号', suffixes=['','_1']).drop('货号', axis=1)
    master.to_csv(home_dir+"Master_"+channel+"_1617.csv")

GENERATE_MASTER("TB")
GENERATE_MASTER("JD")
GENERATE_MASTER("YG")
GENERATE_MASTER("VIP")