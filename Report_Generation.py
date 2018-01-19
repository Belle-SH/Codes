# -*- coding: utf-8 -*-
import pandas as pd
import xlwings as xw
import numpy as np
import datetime

starttime = datetime.datetime.now()

home_dir = "E:/Belle/Staccato/"

# 读取报表header里的筛选条件
filter_condition = []
for i in range(5):
    if i != 3:
        filter_condition.append(pd.read_excel(home_dir+"报表.xlsx", sheetname=i, encoding='gbk'))
    else:
        filter_condition.append(pd.read_excel(home_dir+"报表.xlsx", sheetname=i, encoding='gbk', skiprows=[0,1]))

# 读取线上2016,2017的销售数据
def read_master(channel):
    master = pd.read_csv(home_dir+"Master_"+channel+"_1617.csv", encoding='gbk')
    master = master[(master['商品季'] == filter_condition[0]['季节'].values[0]) & (master['三级分类'] == filter_condition[0]['款式'].values[0]) & (pd.to_datetime(master['日期']) <= pd.to_datetime(filter_condition[0]['检查日期'].values[0]))]
    master.loc[master['首次上架年份'] < 2016, '2016类别'] = "常青款"
    master_2016 = master[master['日期'].str.slice(0,4) == "2016"]
    sale_cnt_rank = pd.pivot_table(master_2016, values='成交件数', index='商品款号', aggfunc=np.sum).sort_values('成交件数', ascending=False)
    evergreen_list = master[master['2016类别'] == "常青款"]['商品款号'].drop_duplicates()
    for i in evergreen_list:
        if i in sale_cnt_rank.index:
            sale_cnt_rank = sale_cnt_rank.drop(i)
    count = len(sale_cnt_rank)
    for id in sale_cnt_rank[0:int(0.2*count)].index.values:
        master.loc[master['商品款号'] == id, '2016类别'] = '畅销款'
    for id in sale_cnt_rank[int(0.2*count)+1:int(0.8*count)].index.values:
        master.loc[master['商品款号'] == id, '2016类别'] = '平销款'
    for id in sale_cnt_rank[int(0.8*count)+1:count+1].index.values:
        master.loc[master['商品款号'] == id, '2016类别'] = '滞销款'
    master.loc[master['首次上架年份'] < 2017, '2017类别'] = "常青款"
    master_2017 = master[master['日期'].str.slice(0,4) == "2017"]
    sale_cnt_rank = pd.pivot_table(master_2017, values='成交件数', index='商品款号', aggfunc=np.sum).sort_values('成交件数', ascending=False)
    evergreen_list = master[master['2017类别'] == "常青款"]['商品款号'].drop_duplicates()
    for i in evergreen_list:
        if i in sale_cnt_rank.index:
            sale_cnt_rank = sale_cnt_rank.drop(i)
    count = len(sale_cnt_rank)
    for id in sale_cnt_rank[0:int(0.2*count)].index.values:
        master.loc[master['商品款号'] == id, '2017类别'] = '畅销款'
    for id in sale_cnt_rank[int(0.2*count)+1:int(0.8*count)].index.values:
        master.loc[master['商品款号'] == id, '2017类别'] = '平销款'
    for id in sale_cnt_rank[int(0.8*count)+1:count+1].index.values:
        master.loc[master['商品款号'] == id, '2017类别'] = '滞销款'
    return master
master_tb = read_master('TB')
master_yg = read_master('YG')
master_jd = read_master('JD')
master_vip = read_master('VIP')
master_all = pd.concat([master_tb, master_yg, master_jd, master_vip], axis=0)

wb = xw.Book(home_dir+"报表.xlsx")
sht = []
for i in range(1,6):
    sht.append(wb.sheets['Sheet'+str(i)])
    
master_id_cate = master_all[['商品款号','2016类别','2017类别']].drop_duplicates()
item_effect = pd.read_csv('E:/Belle/Sycm/items_effect_staccato_GBK.csv', encoding='GBK')
item_effect['title'] = item_effect['title'].str.replace('STACCATO','')
item_effect['title'] = item_effect['title'].str.replace('staccato','')
item_effect['title'] = item_effect['title'].str.replace('Staccato','')
item_effect['product_id'] = item_effect['title'].str.extract("([A-Za-z0-9]{8})", expand=False)
item_effect = item_effect[['itemuv', 'payrate', 'avgstaytime', 'addcartitemcnt', 'favbuyercnt', 'select_date_begin', 'product_id']]
item_effect = item_effect[item_effect['product_id'].notnull()]
prop = pd.read_excel(home_dir+"思加图商品属性.xlsx", encoding='gbk')
prop = prop[['商品编号', '货号', '商品款号', '三级分类', '商品销售季', '商品季', '首次上架时间', '首次上架年份', '首次上架月份']]
prop_dedup = prop.drop_duplicates('商品款号', keep='first')
master = pd.merge(item_effect, prop_dedup, how='left', left_on='product_id', right_on='商品款号', suffixes=['','_1']).drop('商品款号', axis=1)
master = pd.merge(master, master_id_cate, how='left', left_on='product_id', right_on='商品款号', suffixes=['','_1']).drop('商品款号', axis=1)

#filter_res = master_all[(master_all['商品季'] == filter_condition[0]['季节'].values[0]) & (master_all['三级分类'] == filter_condition[0]['款式'].values[0]) & (pd.to_datetime(master_all['日期']) <= pd.to_datetime(filter_condition[0]['检查日期'].values[0]))]
filter_res = master_all.copy()

i = 0
for category in ['常青款','畅销款','平销款','滞销款']:
    filter_category = filter_res[filter_res[str(pd.to_datetime(filter_condition[0]['检查日期'].values[0]).year)+'类别'] == category]
    spu = len(filter_category['商品款号'].unique())
    sku = len(filter_category['供应商款色编号'].unique())
    sale_cnt = int(sum(filter_category['成交件数']))
    sale_amt = int(sum(filter_category['成交金额']))
    dscnt = round(sum(filter_category['折扣率'])/len(filter_category['折扣率']),2)
    sht[0].range('C'+str(7+i)+':G'+str(7+i)).value = [spu, sku, sale_cnt, sale_amt, dscnt]
    i += 1

# 粒度设定 = 周
tw_list_2016 = pd.date_range(start="2016-01-01",end="2016-12-31",freq='W-MON')
tw_list_2017 = pd.date_range(start="2017-01-01",end="2017-12-31",freq='W-MON')

# 继续按照类别筛选（常青，畅销，平销，滞销）
filter_res_2016 = filter_res[filter_res['2016类别'] == filter_condition[1]['类别'].values[0]]
filter_res_2017 = filter_res[filter_res['2017类别'] == filter_condition[1]['类别'].values[0]]

# 2016加总
tw_sale_cnt_2016 = {}
tw_dscnt_2016 = {}
tw_sale_amt_2016 = {}
for tw_number in range(1,len(tw_list_2016)):
    tw_start = tw_list_2016[tw_number - 1]
    tw_end = tw_list_2016[tw_number]
    tw_sale_cnt_2016[tw_start] = int(np.sum(filter_res_2016['成交件数'][(pd.to_datetime(filter_res_2016['日期']) >= tw_start) & (pd.to_datetime(filter_res_2016['日期']) < tw_end)]))
    tw_dscnt_2016[tw_start] = np.average(filter_res_2016['折扣率'][(pd.to_datetime(filter_res_2016['日期']) >= tw_start) & (pd.to_datetime(filter_res_2016['日期']) < tw_end)])*100
    tw_sale_amt_2016[tw_start] = int(np.sum(filter_res_2016['成交金额'][(pd.to_datetime(filter_res_2016['日期']) >= tw_start) & (pd.to_datetime(filter_res_2016['日期']) < tw_end)]))

# 2017加总
tw_sale_cnt_2017 = {}
tw_dscnt_2017 = {}
tw_sale_amt_2017 = {}
for tw_number in range(1,len(tw_list_2017)):
    line_number = tw_number + 33
    line_number2 = tw_number + 132
    tw_start = tw_list_2017[tw_number - 1]
    tw_end = tw_list_2017[tw_number]
    tw_sale_cnt_2017[tw_start] = int(np.sum(filter_res_2017['成交件数'][(pd.to_datetime(filter_res_2017['日期']) >= tw_start) & (pd.to_datetime(filter_res_2017['日期']) < tw_end)]))
    tw_dscnt_2017[tw_start] = np.average(filter_res_2017['折扣率'][(pd.to_datetime(filter_res_2017['日期']) >= tw_start) & (pd.to_datetime(filter_res_2017['日期']) < tw_end)])*100
    tw_sale_amt_2017[tw_start] = int(np.sum(filter_res_2017['成交金额'][(pd.to_datetime(filter_res_2017['日期']) >= tw_start) & (pd.to_datetime(filter_res_2017['日期']) < tw_end)]))
    sht[1].range('A'+str(line_number)).value = tw_number
    sht[1].range('A'+str(line_number2)).value = tw_number                      
    sht[1].range('H'+str(line_number2)).value = tw_number
    sht[1].range('B'+str(line_number)).value = tw_start
    if tw_number < len(tw_list_2016):
        sht[1].range('C'+str(line_number)).value = list(tw_sale_cnt_2016.values())[tw_number-1]
        sht[1].range('I'+str(line_number)).value = str(list(tw_dscnt_2016.values())[tw_number-1]) + '%'
        sht[1].range('I'+str(line_number)).number_format = '0%'
        if sht[1].range('I'+str(line_number)).value == 'nan%':
            sht[1].range('I'+str(line_number)).value = ""
        sht[1].range('B'+str(line_number2)).value = list(tw_sale_amt_2016.values())[tw_number-1]
        sht[1].range('B'+str(line_number2)).number_format = "0"
    sht[1].range('D'+str(line_number)).value = list(tw_sale_cnt_2017.values())[tw_number-1]
    sht[1].range('J'+str(line_number)).value = str(list(tw_dscnt_2017.values())[tw_number-1]) + '%'
    sht[1].range('J'+str(line_number)).number_format = '0%'
    if sht[1].range('J'+str(line_number)).value == 'nan%':
        sht[1].range('J'+str(line_number)).value = ""
    sht[1].range('C'+str(line_number2)).value = list(tw_sale_amt_2017.values())[tw_number-1]
    sht[1].range('C'+str(line_number2)).number_format = "0"
        
master_filter = master[(master['商品季'] == filter_condition[1]['季节'].values[0]) & (master['三级分类'] == filter_condition[1]['款式'].values[0]) & (master[str(pd.to_datetime(filter_condition[1]['检查日期'].values[0]).year)+'类别'] == filter_condition[1]['类别'].values[0])]

addcartitemcnt = []
itemuv = []
avgstaytime = []
payrate = []
def CALCULATE_ITEM_EFFECT(today):
    thirty_days_before = today + pd.Timedelta('-30 days')
    for date in pd.date_range(start=thirty_days_before, end=today):
        master_filter_today = master_filter[pd.to_datetime(master_filter['select_date_begin']) == date]
        filter_res_2017_today = filter_res_2017[pd.to_datetime(filter_res_2017['日期']) == date]
        join_res = pd.merge(master_filter_today, filter_res_2017_today, how='left', left_on = 'product_id', right_on = '商品款号', suffixes=['','_1'])
        addcartitemcnt.append(np.average(master_filter_today['addcartitemcnt']))
        itemuv.append(np.average(master_filter_today['itemuv']))
        avgstaytime.append(round(np.average(master_filter_today['avgstaytime']),1))
        total = 0
        for pid in master_filter_today['product_id']:
            total += (sum(join_res[join_res['product_id'] == pid]['成交件数']) / master_filter_today[master_filter_today['product_id'] == pid]['itemuv']).values[0]
        if len(master_filter_today['product_id']) != 0:
            payrate.append(round(total/len(master_filter_today['product_id'])*100,2))
        else:
            payrate.append(0)
        
def UPDATE_ITEM_EFFECT(sht_number):
    today = pd.to_datetime(filter_condition[sht_number]['检查日期'].values[0])
    CALCULATE_ITEM_EFFECT(today)
    web_tw_dict = {'addcartitemcnt':addcartitemcnt, 'itemuv':itemuv, 'avgstaytime':avgstaytime, 'payrate':payrate}
    web_tw = pd.DataFrame(web_tw_dict)
    
    for i in range(-7,0):
        line_number1 = 215
        line_number2 = 245
        sht[sht_number].range('A'+str(line_number1+i)).value = web_tw['itemuv'].fillna(0)[30+i]
        sht[sht_number].range('B'+str(line_number1+i)).value = np.sum(web_tw['itemuv'][23:30].dropna())/7
        sht[sht_number].range('C'+str(line_number1+i)).value = np.sum(web_tw['itemuv'][0:30].dropna())/30
        sht[sht_number].range('H'+str(line_number1+i)).value = str(web_tw['payrate'].fillna(0)[30+i]) + '%'
        sht[sht_number].range('I'+str(line_number1+i)).value = str(round(np.sum(web_tw['payrate'][23:30].dropna())/7,2)) + '%'
        sht[sht_number].range('J'+str(line_number1+i)).value = str(round(np.sum(web_tw['payrate'][0:30].dropna())/30,2)) + '%'
        sht[sht_number].range('A'+str(line_number2+i)).value = web_tw['avgstaytime'].fillna(0)[30+i]
        sht[sht_number].range('B'+str(line_number2+i)).value = np.sum(web_tw['avgstaytime'][23:30].dropna())/7
        sht[sht_number].range('C'+str(line_number2+i)).value = np.sum(web_tw['avgstaytime'][0:30].dropna())/30
        sht[sht_number].range('H'+str(line_number2+i)).value = web_tw['addcartitemcnt'].fillna(0)[30+i]
        sht[sht_number].range('I'+str(line_number2+i)).value = np.sum(web_tw['addcartitemcnt'][23:30].dropna())/7
        sht[sht_number].range('J'+str(line_number2+i)).value = np.sum(web_tw['addcartitemcnt'][0:30].dropna())/30

UPDATE_ITEM_EFFECT(1)
    
# 计算到款Level的详细信息
#today = pd.to_datetime(filter_condition[1]['检查日期'].values[0])
#for i in range(len(filter_condition[3])):
#    line_number = str(i + 4)
#    sku_id = filter_condition[3]['商品编码'][i]
#    spu_id = prop[prop['货号'] == sku_id]['商品款号']
#    if len(spu_id.values) == 0:
#        continue
#    master_id = master[master['product_id'] == spu_id.values[0]]
#    master_tb_id = master_tb[master_tb['供应商款色编号'] == sku_id][['成交件数', '日期']]
#    addcartitemcnt = []
#    itemuv = []
#    avgstaytime = []
#    payrate = []
#    sale_cnt = []
#    for j in range(4):
#        day_diff_start = 28 - 7 * j
#        day_diff_end = 21 - 7 * j
#        tw_start = today + pd.Timedelta('-'+str(day_diff_start)+' days')
#        tw_end = today + pd.Timedelta('-'+str(day_diff_end)+' days')
#        master_id_tw = master_id[(pd.to_datetime(master_id['select_date_begin']) >= tw_start) & (pd.to_datetime(master_id['select_date_begin']) < tw_end)]
#        master_tb_id_tw = master_tb_id[(pd.to_datetime(master_tb_id['日期']) >= tw_start) & (pd.to_datetime(master_tb_id['日期']) < tw_end)]
#        addcartitemcnt.append(np.sum(master_id_tw['addcartitemcnt'].dropna())/7)
#        itemuv.append(np.sum(master_id_tw['itemuv'].dropna())/7)
#        avgstaytime.append(np.sum(master_id_tw['avgstaytime'].dropna())/7)
#        sale_cnt.append(np.sum(master_tb_id_tw['成交件数'].dropna())/7)
#    sht[3].range('S'+line_number+':'+'V'+line_number).value = pd.Series(itemuv).values
#    sht[3].range('W'+line_number+':'+'Z'+line_number).value = pd.Series(sale_cnt).values
#    ratio = pd.Series(sale_cnt)/pd.Series(itemuv)
#    ratio = ratio.replace([np.inf, -np.inf, np.nan], 0) 
#    sht[3].range('AC'+line_number+':'+'AF'+line_number).value = ratio.values
#    sht[3].range('AG'+line_number+':'+'AJ'+line_number).value = pd.Series(avgstaytime).values
#    sht[3].range('AK'+line_number+':'+'AN'+line_number).value = pd.Series(addcartitemcnt).values    
    
# 计算到款Level的曲线
#filter_res = master_all[(master_all['供应商款色编号'] == filter_condition[4]['货号'].values[0]) & (pd.to_datetime(master_all['日期']) <= pd.to_datetime(filter_condition[4]['检查日期'].values[0]))]
#
#tw_sale_cnt_2016 = {}
#tw_dscnt_2016 = {}
#tw_sale_amt_2016 = {}
#for tw_number in range(1,len(tw_list_2016)):
#    tw_start = tw_list_2016[tw_number - 1]
#    tw_end = tw_list_2016[tw_number]
#    tw_sale_cnt_2016[tw_start] = int(np.sum(filter_res['成交件数'][(pd.to_datetime(filter_res['日期']) >= tw_start) & (pd.to_datetime(filter_res['日期']) < tw_end)]))
#    tw_dscnt_2016[tw_start] = np.average(filter_res['折扣率'][(pd.to_datetime(filter_res['日期']) >= tw_start) & (pd.to_datetime(filter_res['日期']) < tw_end)])*100
#    tw_sale_amt_2016[tw_start] = int(np.sum(filter_res['成交金额'][(pd.to_datetime(filter_res['日期']) >= tw_start) & (pd.to_datetime(filter_res['日期']) < tw_end)]))
#
## 2017加总
#tw_sale_cnt_2017 = {}
#tw_dscnt_2017 = {}
#tw_sale_amt_2017 = {}
#for tw_number in range(1,len(tw_list_2017)):
#    line_number = tw_number + 33
#    line_number2 = tw_number + 132
#    tw_start = tw_list_2017[tw_number - 1]
#    tw_end = tw_list_2017[tw_number]
#    tw_sale_cnt_2017[tw_start] = int(np.sum(filter_res['成交件数'][(pd.to_datetime(filter_res['日期']) >= tw_start) & (pd.to_datetime(filter_res['日期']) < tw_end)]))
#    tw_dscnt_2017[tw_start] = np.average(filter_res['折扣率'][(pd.to_datetime(filter_res['日期']) >= tw_start) & (pd.to_datetime(filter_res['日期']) < tw_end)])*100
#    tw_sale_amt_2017[tw_start] = int(np.sum(filter_res['成交金额'][(pd.to_datetime(filter_res['日期']) >= tw_start) & (pd.to_datetime(filter_res['日期']) < tw_end)]))
#    sht[4].range('A'+str(line_number)).value = tw_number
#    sht[4].range('A'+str(line_number2)).value = tw_number                      
#    sht[4].range('H'+str(line_number2)).value = tw_number
#    sht[4].range('B'+str(line_number)).value = tw_start
#    if tw_number < len(tw_list_2016):
#        sht[4].range('C'+str(line_number)).value = list(tw_sale_cnt_2016.values())[tw_number-1]
#        sht[4].range('I'+str(line_number)).value = str(list(tw_dscnt_2016.values())[tw_number-1]) + '%'
#        sht[4].range('I'+str(line_number)).number_format = '0%'
#        if sht[4].range('I'+str(line_number)).value == 'nan%':
#            sht[4].range('I'+str(line_number)).value = ""
#        sht[4].range('B'+str(line_number2)).value = list(tw_sale_amt_2016.values())[tw_number-1]
#        sht[4].range('B'+str(line_number2)).number_format = "0"
#    sht[4].range('D'+str(line_number)).value = list(tw_sale_cnt_2017.values())[tw_number-1]
#    sht[4].range('J'+str(line_number)).value = str(list(tw_dscnt_2017.values())[tw_number-1]) + '%'
#    sht[4].range('J'+str(line_number)).number_format = '0%'
#    if sht[4].range('J'+str(line_number)).value == 'nan%':
#        sht[4].range('J'+str(line_number)).value = ""
#    sht[4].range('C'+str(line_number2)).value = list(tw_sale_amt_2017.values())[tw_number-1]
#    sht[4].range('C'+str(line_number2)).number_format = "0"
#
#master_filter = master[(master['货号'] == filter_condition[4]['货号'].values[0]) & (pd.to_datetime(master['select_date_begin']) <= pd.to_datetime(filter_condition[4]['检查日期'].values[0]))]
#
#addcartitemcnt = []
#itemuv = []
#avgstaytime = []
#payrate = []
#UPDATE_ITEM_EFFECT(4)

endtime = datetime.datetime.now()
print('Running Time:' + str((endtime - starttime).seconds) + 's')