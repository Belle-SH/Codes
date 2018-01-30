# -*- coding: utf-8 -*-
import pandas as pd
import xlwings as xw
import numpy as np
import datetime

starttime = datetime.datetime.now()

home_dir = "E:/Belle/Staccato/"

# 读取报表header里的筛选条件
filter_condition = []
for i in range(2):
    filter_condition.append(pd.read_excel(home_dir+"报表.xlsx", sheetname=i, encoding='gbk'))
#for i in range(5):
#    if i != 3:
#        filter_condition.append(pd.read_excel(home_dir+"报表全.xlsx", sheetname=i, encoding='gbk'))
#    else:
#        filter_condition.append(pd.read_excel(home_dir+"报表全.xlsx", sheetname=i, encoding='gbk', skiprows=[0,1]))

# 读取商品属性表      
prop = pd.read_csv(home_dir+'Staccato_Women_Prop_1617.csv',encoding='gbk')

# 读取线上2016,2017的销售数据
master = pd.read_csv(home_dir+'Staccato_Master_1617_Spring_Summer.csv')
master = master[(master['季节'] == filter_condition[0]['季节'].values[0]) & (master['三级分类'] == filter_condition[0]['款式'].values[0]) & (pd.to_datetime(master['日期']) <= pd.to_datetime(filter_condition[0]['检查日期'].values[0]))]
master_sub = master[['款号','成交件数','日期']]
master_all = pd.merge(master,prop,how='left',left_on='供应商款色编号',right_on='供应商款色编码',suffixes=['','_1']).drop(['商品编号_1','供应商款色编号','款号_1','销售季','三级分类_1','季节_1'], axis=1)
master_all['2016新旧'] = master_all['款号'].str.slice(-1) < '6'
master_all['2017新旧'] = master_all['款号'].str.slice(-1) < '7'

# 读取2016,2017库存数据
inv = pd.read_csv(home_dir+'Staccato_Inventory_1617.csv',names=['日期','商品编号','库存'])

# 设定sht
wb = xw.Book(home_dir+"报表.xlsx")
sht = []
for i in range(2):
    sht.append(wb.sheets[i])

# 读入生意参谋-商品效果表，并与商品属性表整合
item_effect = pd.read_csv('E:/Belle/Sycm/items_effect_staccato_GBK.csv',encoding='GBK')
item_effect['title'] = item_effect['title'].str.replace('Staccato','')
item_effect['title'] = item_effect['title'].str.replace('STACCATO','')
item_effect['title'] = item_effect['title'].str.replace('staccato','')
item_effect['product_id'] = item_effect['title'].str.extract("([A-Za-z0-9]{8})", expand=False)
item_effect = item_effect[['itemuv', 'payrate', 'avgstaytime', 'addcartitemcnt', 'favbuyercnt', 'select_date_begin', 'product_id']]
item_effect = item_effect[item_effect['product_id'].notnull()]
prop_dedup = prop.drop_duplicates('款号', keep='first')
master = pd.merge(item_effect, prop_dedup, how='left', left_on='product_id', right_on='款号', suffixes=['','_1']).drop('款号', axis=1)
master['2016新旧'] = master['product_id'].str.slice(-1) < '6'
master['2017新旧'] = master['product_id'].str.slice(-1) < '7'

filter_res = master_all.copy()
one_year_ago = pd.to_datetime(filter_condition[0]['检查日期'].values[0]) + pd.Timedelta('-365 days')
filter_res_2016 = master_all[pd.to_datetime(master_all['日期']) <= one_year_ago]

i = 0
for category in ['畅','平','滞']:
    filter_2017 = filter_res[(pd.to_datetime(filter_res['日期']) >= pd.to_datetime('2017-01-01')) & (filter_res['2017类别'] == category)]
    filter_2017_impute = filter_2017[filter_2017['成交件数'] != 0]
    filter_2016 = filter_res_2016[filter_res_2016['2016类别'] == category]
    # 在2016的SKU中去掉张亚萍表中误标的17年SKU
    filter_2016 = filter_2016[filter_res_2016['2017新旧']]
    filter_2016_impute = filter_2016[filter_2016['成交件数'] != 0]
    spu_2017 = len(filter_2017['款号'].unique())
    sku_2017 = len(filter_2017['供应商款色编码'].unique())
    sale_cnt_2017 = int(sum(filter_2017['成交件数']))
    sale_amt_2017 = int(sum(filter_2017['成交金额']))
    inv_total = 0
    for item_id in filter_2017['商品编号'].unique():
        if len(inv[inv['商品编号'] == item_id]) != 0:
            inv_total += inv[inv['商品编号'] == item_id].sort_values(['日期'])['库存'].values[-1]
    if len(filter_2017_impute['折扣率']) != 0:
        dscnt_2017 = round(sum(filter_2017_impute['折扣率'])/len(filter_2017_impute['折扣率']),2)
    else:
        dscnt_2017 = '0'
    gross_profit_2017 = np.average(filter_2017_impute['毛利率'])
    sold_out_2017 = sale_cnt_2017/(sale_cnt_2017+inv_total)
    spu_2016 = len(filter_2016['款号'].unique())
    sku_2016 = len(filter_2016['供应商款色编码'].unique())
    sale_cnt_2016 = int(sum(filter_2016['成交件数']))
    sale_amt_2016 = int(sum(filter_2016['成交金额']))
    inv_total = 0
    for item_id in filter_2016['商品编号'].unique():
        if len(inv[inv['商品编号'] == item_id]) != 0:
            inv_total += inv[inv['商品编号'] == item_id].sort_values(['日期'])['库存'].values[-1]
    if len(filter_2016_impute['折扣率']) != 0:
        dscnt_2016 = round(sum(filter_2016_impute['折扣率'])/len(filter_2016_impute['折扣率']),2)
    else:
        dscnt_2016 = '0'
    gross_profit_2016 = np.average(filter_2016_impute['毛利率'])
    sold_out_2016 = sale_cnt_2016/(sale_cnt_2016+inv_total)                        
#    filter_2017_new = filter_2017[filter_2017['2017新旧'] == False]
#    filter_2017_old = filter_2017[filter_2017['2017新旧'] == True]
#    filter_2016_new = filter_2016[filter_2016['2016新旧'] == False]
#    filter_2016_old = filter_2016[filter_2016['2016新旧'] == True]
#    spu_2017_new = len(filter_2017_new['款号'].unique())
#    sku_2017_new = len(filter_2017_new['供应商款色编码'].unique())
#    sale_cnt_2017_new = int(sum(filter_2017_new['成交件数']))
#    sale_amt_2017_new = int(sum(filter_2017_new['成交金额']))
#    if len(filter_2017_new['折扣率']) != 0:
#        dscnt_2017_new = round(sum(filter_2017_new['折扣率'])/len(filter_2017_new['折扣率']),2)
#    else:
#        dscnt_2017_new = '0'
#    sht[0].range('C'+str(7+i)+':G'+str(7+i)).value = [spu_2017_new, sku_2017_new, sale_cnt_2017_new, sale_amt_2017_new, dscnt_2017_new]
#    spu_2017_old = len(filter_2017_old['款号'].unique())
#    sku_2017_old = len(filter_2017_old['供应商款色编码'].unique())
#    sale_cnt_2017_old = int(sum(filter_2017_old['成交件数']))
#    sale_amt_2017_old = int(sum(filter_2017_old['成交金额']))
#    if len(filter_2017_old['折扣率']) != 0:
#        dscnt_2017_old = round(sum(filter_2017_old['折扣率'])/len(filter_2017_old['折扣率']),2)
#    else:
#        dscnt_2017_old = '0'
#    sht[0].range('J'+str(7+i)+':N'+str(7+i)).value = [spu_2017_old, sku_2017_old, sale_cnt_2017_old, sale_amt_2017_old, dscnt_2017_old]
#    spu_2016_new = len(filter_2016_new['款号'].unique())
#    sku_2016_new = len(filter_2016_new['供应商款色编码'].unique())
#    sale_cnt_2016_new = int(sum(filter_2016_new['成交件数']))
#    sale_amt_2016_new = int(sum(filter_2016_new['成交金额']))
#    if len(filter_2016_new['折扣率']) != 0:
#        dscnt_2016_new = round(sum(filter_2016_new['折扣率'])/len(filter_2016_new['折扣率']),2)
#    else:
#        dscnt_2016_new = '0'
#    sht[0].range('C'+str(15+i)+':G'+str(15+i)).value = [spu_2016_new, sku_2016_new, sale_cnt_2016_new, sale_amt_2016_new, dscnt_2016_new]
#    spu_2016_old = len(filter_2016_old['款号'].unique())
#    sku_2016_old = len(filter_2016_old['供应商款色编码'].unique())
#    sale_cnt_2016_old = int(sum(filter_2016_old['成交件数']))
#    sale_amt_2016_old = int(sum(filter_2016_old['成交金额']))
#    if len(filter_2016_old['折扣率']) != 0:
#        dscnt_2016_old = round(sum(filter_2016_old['折扣率'])/len(filter_2016_old['折扣率']),2)
#    else:
#        dscnt_2016_old = '0'
#    sht[0].range('J'+str(15+i)+':N'+str(15+i)).value = [spu_2016_old, sku_2016_old, sale_cnt_2016_old, sale_amt_2016_old, dscnt_2016_old]
    sht[0].range('C'+str(7+i)+':G'+str(7+i)).value = [spu_2017,sku_2017,sale_cnt_2017,sale_amt_2017,dscnt_2017,gross_profit_2017,sold_out_2017]
    sht[0].range('L'+str(15+i)+':R'+str(15+i)).value = [spu_2016,sku_2016,sale_cnt_2016,sale_amt_2016,dscnt_2016,gross_profit_2016,sold_out_2016]
    i += 1

# 粒度设定 = 周
tw_list_2016 = pd.date_range(start="2016-01-01",end="2016-12-31",freq='W-MON')
tw_list_2017 = pd.date_range(start="2017-01-01",end="2017-12-31",freq='W-MON')

# 继续按照类别筛选（常青，畅销，平销，滞销）
if (filter_condition[1]['新旧'].values[0] == '新'):
    filter_res_2016 = filter_res[(filter_res['2016类别'] == filter_condition[1]['类别'].values[0]) & (filter_res['2016新旧'] == False) & (filter_res['2017新旧'] == True)]
    filter_res_2017 = filter_res[(filter_res['2017类别'] == filter_condition[1]['类别'].values[0]) & (filter_res['2017新旧'] == False)]
else:
    filter_res_2016 = filter_res[(filter_res['2016类别'] == filter_condition[1]['类别'].values[0]) & (filter_res['2016新旧'] == True) & (filter_res['2017新旧'] == True)]
    filter_res_2017 = filter_res[(filter_res['2017类别'] == filter_condition[1]['类别'].values[0]) & (filter_res['2017新旧'] == True)]

# 2016加总
tw_sale_cnt_2016 = {}
tw_dscnt_2016 = {}
tw_sale_amt_2016 = {}
tw_sold_out_2016 = {}
sku_cnt_2016 = len(filter_res_2016['供应商款色编码'].unique())
sku_cnt_2017 = len(filter_res_2017['供应商款色编码'].unique())

for tw_number in range(1,len(tw_list_2016)):
    tw_start = tw_list_2016[tw_number - 1]
    tw_end = tw_list_2016[tw_number]
    tw_sale_cnt_2016[tw_start] = np.sum(filter_res_2016['成交件数'][(pd.to_datetime(filter_res_2016['日期']) >= tw_start) & (pd.to_datetime(filter_res_2016['日期']) < tw_end)]) / sku_cnt_2016
    tw_dscnt_2016[tw_start] = np.average(filter_res_2016['折扣率'][(pd.to_datetime(filter_res_2016['日期']) >= tw_start) & (pd.to_datetime(filter_res_2016['日期']) < tw_end)])*100
    tw_sale_amt_2016[tw_start] = np.sum(filter_res_2016['成交金额'][(pd.to_datetime(filter_res_2016['日期']) >= tw_start) & (pd.to_datetime(filter_res_2016['日期']) < tw_end)]) / sku_cnt_2016
    #累计售罄率 = 累计销量 / (累计销量 + 期末库存)
    sale_total = np.sum(filter_res_2016['成交件数'][(pd.to_datetime(filter_res_2016['日期']) >= pd.to_datetime('2016-01-01')) & (pd.to_datetime(filter_res_2016['日期']) < tw_end)])
    inv_total = 0
    for pid in filter_res_2016['商品编号'].unique():
        inv_pid = inv[inv['商品编号'] == pid].sort_values('日期')
        if sum(pd.to_datetime(inv_pid['日期']) <= tw_end) > 0:
            inv_total += inv_pid[pd.to_datetime(inv_pid['日期']) <= tw_end]['库存'].values[-1]
    if (sale_total + inv_total) != 0:
        tw_sold_out_2016[tw_start] = sale_total / (sale_total + inv_total)
    else:
        tw_sold_out_2016[tw_start] = 0

# 2017加总
tw_sale_cnt_2017 = {}
tw_dscnt_2017 = {}
tw_sale_amt_2017 = {}
tw_sold_out_2017 = {}
for tw_number in range(1,len(tw_list_2017)):
    line_number = tw_number + 27
    line_number2 = tw_number + 104
    tw_start = tw_list_2017[tw_number - 1]
    tw_end = tw_list_2017[tw_number]
    tw_sale_cnt_2017[tw_start] = np.sum(filter_res_2017['成交件数'][(pd.to_datetime(filter_res_2017['日期']) >= tw_start) & (pd.to_datetime(filter_res_2017['日期']) < tw_end)]) / sku_cnt_2017
    tw_dscnt_2017[tw_start] = np.average(filter_res_2017['折扣率'][(pd.to_datetime(filter_res_2017['日期']) >= tw_start) & (pd.to_datetime(filter_res_2017['日期']) < tw_end)])*100
    tw_sale_amt_2017[tw_start] = np.sum(filter_res_2017['成交金额'][(pd.to_datetime(filter_res_2017['日期']) >= tw_start) & (pd.to_datetime(filter_res_2017['日期']) < tw_end)]) / sku_cnt_2017
    #累计售罄率 = 累计销量 / (累计销量 + 期末库存)
    sale_total = np.sum(filter_res_2017['成交件数'][(pd.to_datetime(filter_res_2017['日期']) >= pd.to_datetime('2017-01-01')) & (pd.to_datetime(filter_res_2017['日期']) < tw_end)])
    inv_total = 0
    inv_total_2 = 0
    for pid in filter_res_2017['商品编号'].unique():
        inv_pid = inv[inv['商品编号'] == pid].sort_values('日期')
        if sum(pd.to_datetime(inv_pid['日期']) <= tw_end) > 0:
            inv_total += inv_pid[pd.to_datetime(inv_pid['日期']) <= tw_end]['库存'].values[-1]
        if sum(pd.to_datetime(inv_pid['日期']) <= tw_start) > 0:
            inv_total_2 += inv_pid[pd.to_datetime(inv_pid['日期']) <= tw_start]['库存'].values[-1]
    if (sale_total + inv_total) != 0:
        tw_sold_out_2017[tw_start] = sale_total / (sale_total + inv_total)
    else:
        tw_sold_out_2017[tw_start] = 0
    sht[1].range('B'+str(line_number)).value = tw_start
    sht[1].range('I'+str(line_number)).value = tw_start
    if tw_number < len(tw_list_2016):
        sht[1].range('C'+str(line_number)).value = list(tw_sale_cnt_2016.values())[tw_number-1]
        sht[1].range('C'+str(line_number2)).value = str(list(tw_dscnt_2016.values())[tw_number-1]) + '%'
        sht[1].range('C'+str(line_number2)).number_format = '0%'
        if sht[1].range('C'+str(line_number2)).value == 'nan%':
            sht[1].range('C'+str(line_number2)).value = ""
        sht[1].range('J'+str(line_number2)).value = list(tw_sale_amt_2016.values())[tw_number-1]
        sht[1].range('J'+str(line_number2)).number_format = "0"
    sht[1].range('L'+str(line_number)).value = inv_total
    sht[1].range('M'+str(line_number)).value = inv_total_2
    sht[1].range('J'+str(line_number)).value = list(tw_sold_out_2016.values())[tw_number-1]
    sht[1].range('K'+str(line_number)).value = list(tw_sold_out_2017.values())[tw_number-1]
    sht[1].range('D'+str(line_number)).value = list(tw_sale_cnt_2017.values())[tw_number-1]
    sht[1].range('D'+str(line_number2)).value = str(list(tw_dscnt_2017.values())[tw_number-1]) + '%'
    sht[1].range('D'+str(line_number2)).number_format = '0%'
    if sht[1].range('D'+str(line_number2)).value == 'nan%':
        sht[1].range('D'+str(line_number2)).value = ""
    sht[1].range('K'+str(line_number2)).value = list(tw_sale_amt_2017.values())[tw_number-1]
    sht[1].range('K'+str(line_number2)).number_format = "0"
        
master_filter = master[(master['季节'] == filter_condition[1]['季节'].values[0]) & (master['三级分类'] == filter_condition[1]['款式'].values[0]) & (master['2017类别'] == filter_condition[1]['类别'].values[0]) & (master['2017新旧'] == (filter_condition[1]['新旧'].values[0] == '旧'))]

addcartitemcnt = []
itemuv = []
avgstaytime = []
payrate = []
today = pd.to_datetime(filter_condition[1]['检查日期'].values[0])

for i in range(12,0,-1):
        tw_start = today + pd.Timedelta('-'+str(i*7)+' days')
        tw_end = today + pd.Timedelta('-'+str(i*7-7)+' days')
        master_sub_tmp = master_sub[(pd.to_datetime(master_sub['日期']) > pd.to_datetime(tw_start)) & (pd.to_datetime(master_sub['日期']) <= pd.to_datetime(tw_end))]
        master_filter_tw = master_filter[(pd.to_datetime(master_filter['select_date_begin']) > pd.to_datetime(tw_start)) & (pd.to_datetime(master_filter['select_date_begin']) <= pd.to_datetime(tw_end))]
        addcartitemcnt.append(np.sum(master_filter_tw['addcartitemcnt']))
        itemuv.append(np.sum(master_filter_tw['itemuv']))
        avgstaytime.append(round(np.average(master_filter_tw['avgstaytime']),1))
        total = 0
        for pid in master_filter_tw['product_id'].unique():
            total += sum(master_sub_tmp[master_sub_tmp['款号'] == pid]['成交件数'])
        if len(master_filter_tw['product_id']) != 0:
            payrate.append(round(total /sum(master_filter_tw['itemuv'])*100,2))
        else:
            payrate.append(0)

web_tw_dict = {'addcartitemcnt':addcartitemcnt, 'itemuv':itemuv, 'avgstaytime':avgstaytime, 'payrate':payrate}
web_tw = pd.DataFrame(web_tw_dict)
for i in range(0,12):
        line_number1 = 180
        line_number2 = 214
        start_date = today + pd.Timedelta('-'+str((12-i)*7)+' days')
        sht[1].range('A'+str(line_number1+i)).value = start_date
        sht[1].range('H'+str(line_number1+i)).value = start_date
        sht[1].range('A'+str(line_number2+i)).value = start_date
        sht[1].range('H'+str(line_number2+i)).value = start_date
        sht[1].range('B'+str(line_number1+i)).value = web_tw['itemuv'].fillna(0)[i]
        sht[1].range('I'+str(line_number1+i)).value = str(web_tw['payrate'].fillna(0)[i]) + '%'
        sht[1].range('B'+str(line_number2+i)).value = web_tw['avgstaytime'].fillna(0)[i]
        sht[1].range('I'+str(line_number2+i)).value = web_tw['addcartitemcnt'].fillna(0)[i]
    
# 计算到款Level的详细信息
#today = pd.to_datetime(filter_condition[1]['检查日期'].values[0])
#for i in range(len(filter_condition[3])):
#    line_number = str(i + 4)
#    sku_id = filter_condition[3]['商品编码'][i]
#    spu_id = prop[prop['供应商款色编码'] == sku_id]['商品款号']
#    if len(spu_id.values) == 0:
#        continue
#    master_id = master[master['product_id'] == spu_id.values[0]]
#    master_tb_id = master_tb[master_tb['供应商款色编码'] == sku_id][['成交件数', '日期']]
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
#sht[4].range('D2').value = master_all[master_all['供应商款色编码'] == filter_condition[4]['供应商款色编码'].values[0]]['2017类别'].values[0]
#sht[4].range('E2').value = master_all[master_all['供应商款色编码'] == filter_condition[4]['供应商款色编码'].values[0]]['2016类别'].values[0]
#sht[4].range('F2').value = master_all[master_all['供应商款色编码'] == filter_condition[4]['供应商款色编码'].values[0]]['三级分类'].values[0]
#sht[4].range('G2').value = master_all[master_all['供应商款色编码'] == filter_condition[4]['供应商款色编码'].values[0]]['季节'].values[0]
#filter_res = master_all[(master_all['供应商款色编码'] == filter_condition[4]['供应商款色编码'].values[0]) & (pd.to_datetime(master_all['日期']) <= pd.to_datetime(filter_condition[4]['检查日期'].values[0]))]
#filter_res_2016 = filter_res[pd.to_datetime(filter_res['日期']) < pd.to_datetime('2017/01/01')]
#filter_res_2017 = filter_res[pd.to_datetime(filter_res['日期']) >= pd.to_datetime('2017/01/01')]
#
#tw_sale_cnt_2016 = {}
#tw_dscnt_2016 = {}
#tw_sale_amt_2016 = {}
#tw_sold_out_2016 = {}
#
#for tw_number in range(1,len(tw_list_2016)):
#    tw_start = tw_list_2016[tw_number - 1]
#    tw_end = tw_list_2016[tw_number]
#    tw_sale_cnt_2016[tw_start] = int(np.sum(filter_res_2016['成交件数'][(pd.to_datetime(filter_res_2016['日期']) >= tw_start) & (pd.to_datetime(filter_res_2016['日期']) < tw_end)]))
#    tw_dscnt_2016[tw_start] = np.average(filter_res_2016['折扣率'][(pd.to_datetime(filter_res_2016['日期']) >= tw_start) & (pd.to_datetime(filter_res_2016['日期']) < tw_end)])*100
#    tw_sale_amt_2016[tw_start] = int(np.sum(filter_res_2016['成交金额'][(pd.to_datetime(filter_res_2016['日期']) >= tw_start) & (pd.to_datetime(filter_res_2016['日期']) < tw_end)]))
#    #累计售罄率 = 累计销量 / (累计销量 + 期末库存)
#    sale_total = np.sum(filter_res_2016['成交件数'][(pd.to_datetime(filter_res_2016['日期']) >= pd.to_datetime('2016-01-01')) & (pd.to_datetime(filter_res_2016['日期']) < tw_end)])
#    inv_total = 0
#    for pid in filter_res_2016['商品编号'].unique():
#        inv_pid = inv[inv['商品编号'] == pid].sort_values('日期')
#        if sum(pd.to_datetime(inv_pid['日期']) <= tw_end) > 0:
#            inv_total += inv_pid[pd.to_datetime(inv_pid['日期']) <= tw_end]['库存'].values[-1]
#    if (sale_total + inv_total) != 0:
#        tw_sold_out_2016[tw_start] = sale_total / (sale_total + inv_total)
#    else:
#        tw_sold_out_2016[tw_start] = 0
#
## 2017加总
#tw_sale_cnt_2017 = {}
#tw_dscnt_2017 = {}
#tw_sale_amt_2017 = {}
#tw_sold_out_2017 = {}
#for tw_number in range(1,len(tw_list_2017)):
#    line_number = tw_number + 27
#    line_number2 = tw_number + 104
#    tw_start = tw_list_2017[tw_number - 1]
#    tw_end = tw_list_2017[tw_number]
#    tw_sale_cnt_2017[tw_start] = int(np.sum(filter_res_2017['成交件数'][(pd.to_datetime(filter_res_2017['日期']) >= tw_start) & (pd.to_datetime(filter_res_2017['日期']) < tw_end)]))
#    tw_dscnt_2017[tw_start] = np.average(filter_res_2017['折扣率'][(pd.to_datetime(filter_res_2017['日期']) >= tw_start) & (pd.to_datetime(filter_res_2017['日期']) < tw_end)])*100
#    tw_sale_amt_2017[tw_start] = int(np.sum(filter_res_2017['成交金额'][(pd.to_datetime(filter_res_2017['日期']) >= tw_start) & (pd.to_datetime(filter_res_2017['日期']) < tw_end)]))
#    #累计售罄率 = 累计销量 / (累计销量 + 期末库存)
#    sale_total = np.sum(filter_res_2017['成交件数'][(pd.to_datetime(filter_res_2017['日期']) >= pd.to_datetime('2017-01-01')) & (pd.to_datetime(filter_res_2017['日期']) < tw_end)])
#    inv_total = 0
#    inv_total_2 = 0
#    for pid in filter_res_2017['商品编号'].unique():
#        inv_pid = inv[inv['商品编号'] == pid].sort_values('日期')
#        if sum(pd.to_datetime(inv_pid['日期']) <= tw_end) > 0:
#            inv_total += inv_pid[pd.to_datetime(inv_pid['日期']) <= tw_end]['库存'].values[-1]
#        if sum(pd.to_datetime(inv_pid['日期']) <= tw_start) > 0:
#            inv_total_2 += inv_pid[pd.to_datetime(inv_pid['日期']) <= tw_start]['库存'].values[-1]
#    if (sale_total + inv_total) != 0:
#        tw_sold_out_2017[tw_start] = sale_total / (sale_total + inv_total)
#    else:
#        tw_sold_out_2017[tw_start] = 0
#    sht[4].range('B'+str(line_number)).value = tw_start
#    sht[4].range('I'+str(line_number)).value = tw_start
#    sht[4].range('C'+str(line_number)).value = list(tw_sale_cnt_2016.values())[tw_number-1]
#    sht[4].range('C'+str(line_number2)).value = str(list(tw_dscnt_2016.values())[tw_number-1]) + '%'
#    sht[4].range('C'+str(line_number2)).number_format = '0%'
#    if sht[4].range('C'+str(line_number2)).value == 'nan%':
#        sht[4].range('C'+str(line_number2)).value = ""
#    sht[4].range('J'+str(line_number2)).value = list(tw_sale_amt_2016.values())[tw_number-1]
#    sht[4].range('J'+str(line_number2)).number_format = "0"
#    sht[4].range('L'+str(line_number)).value = inv_total
#    sht[4].range('M'+str(line_number)).value = inv_total_2
#    sht[4].range('J'+str(line_number)).value = list(tw_sold_out_2016.values())[tw_number-1]
#    sht[4].range('K'+str(line_number)).value = list(tw_sold_out_2017.values())[tw_number-1]
#    sht[4].range('D'+str(line_number)).value = list(tw_sale_cnt_2017.values())[tw_number-1]
#    sht[4].range('D'+str(line_number2)).value = str(list(tw_dscnt_2017.values())[tw_number-1]) + '%'
#    sht[4].range('D'+str(line_number2)).number_format = '0%'
#    if sht[4].range('D'+str(line_number2)).value == 'nan%':
#        sht[4].range('D'+str(line_number2)).value = ""
#    sht[4].range('K'+str(line_number2)).value = list(tw_sale_amt_2017.values())[tw_number-1]
#    sht[4].range('K'+str(line_number2)).number_format = "0"
#
#master_filter = master[(master['供应商款色编码'] == filter_condition[4]['供应商款色编码'].values[0]) & (pd.to_datetime(master['select_date_begin']) <= pd.to_datetime(filter_condition[4]['检查日期'].values[0]))]
#
#addcartitemcnt = []
#itemuv = []
#avgstaytime = []
#payrate = []
#today = pd.to_datetime(filter_condition[1]['检查日期'].values[0])
#
#for i in range(12,0,-1):
#        tw_start = today + pd.Timedelta('-'+str(i*7)+' days')
#        tw_end = today + pd.Timedelta('-'+str(i*7-7)+' days')
#        master_sub_tmp = master_sub[(pd.to_datetime(master_sub['日期']) > pd.to_datetime(tw_start)) & (pd.to_datetime(master_sub['日期']) <= pd.to_datetime(tw_end))]
#        master_filter_tw = master_filter[(pd.to_datetime(master_filter['select_date_begin']) > pd.to_datetime(tw_start)) & (pd.to_datetime(master_filter['select_date_begin']) <= pd.to_datetime(tw_end))]
#        addcartitemcnt.append(np.sum(master_filter_tw['addcartitemcnt']))
#        itemuv.append(np.sum(master_filter_tw['itemuv']))
#        avgstaytime.append(round(np.average(master_filter_tw['avgstaytime']),1))
#        total = 0
#        for pid in master_filter_tw['product_id'].unique():
#            total += sum(master_sub_tmp[master_sub_tmp['款号'] == pid]['成交件数'])
#        if len(master_filter_tw['product_id']) != 0:
#            payrate.append(round(total /sum(master_filter_tw['itemuv'])*100,2))
#        else:
#            payrate.append(0)
#
#web_tw_dict = {'addcartitemcnt':addcartitemcnt, 'itemuv':itemuv, 'avgstaytime':avgstaytime, 'payrate':payrate}
#web_tw = pd.DataFrame(web_tw_dict)
#for i in range(0,12):
#        line_number1 = 180
#        line_number2 = 214
#        start_date = today + pd.Timedelta('-'+str((12-i)*7)+' days')
#        sht[4].range('A'+str(line_number1+i)).value = start_date
#        sht[4].range('H'+str(line_number1+i)).value = start_date
#        sht[4].range('A'+str(line_number2+i)).value = start_date
#        sht[4].range('H'+str(line_number2+i)).value = start_date
#        sht[4].range('B'+str(line_number1+i)).value = web_tw['itemuv'].fillna(0)[i]
#        sht[4].range('I'+str(line_number1+i)).value = str(web_tw['payrate'].fillna(0)[i]) + '%'
#        sht[4].range('B'+str(line_number2+i)).value = web_tw['avgstaytime'].fillna(0)[i]
#        sht[4].range('I'+str(line_number2+i)).value = web_tw['addcartitemcnt'].fillna(0)[i]

endtime = datetime.datetime.now()
print('Running Time:' + str((endtime - starttime).seconds) + 's')