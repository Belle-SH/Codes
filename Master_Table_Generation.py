# -*- coding: utf-8 -*-
import pandas as pd
import numpy as np

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
    master.loc[master['首次上架年份'] < 2015, '类别'] = "常青款"
    sale_cnt_rank = pd.pivot_table(master, values='成交件数', index='商品款号', aggfunc=np.sum).sort_values('成交件数', ascending=False)
    evergreen_list = master[master['类别'] == "常青款"]['商品款号'].drop_duplicates()
    for i in evergreen_list:
        sale_cnt_rank = sale_cnt_rank.drop(i)
    count = len(sale_cnt_rank)
    for id in sale_cnt_rank[0:int(0.2*count)].index.values:
        master.loc[master['商品款号'] == id, '类别'] = '畅销款'
    for id in sale_cnt_rank[int(0.2*count)+1:int(0.8*count)].index.values:
        master.loc[master['商品款号'] == id, '类别'] = '平销款'
    for id in sale_cnt_rank[int(0.8*count)+1:count+1].index.values:
        master.loc[master['商品款号'] == id, '类别'] = '滞销款'
    master.to_csv(home_dir+"Master_"+channel+"_1617.csv")

GENERATE_MASTER("TB")
GENERATE_MASTER("JD")
GENERATE_MASTER("YG")
GENERATE_MASTER("VIP")