# -*- coding: utf-8 -*-
import pandas as pd
import xlwings as xw
import numpy as np

home_dir = "E:/Belle/Belle/"
wb = xw.Book(home_dir+"报表到款数据_春季满帮鞋.xlsx")
sht = wb.sheets[0]

# 读取商品属性表      
prop = pd.read_csv(home_dir+'Belle_Women_Prop_1617.csv',encoding='gbk')

# 读取2016,2017库存数据
inv = pd.read_csv(home_dir+'Belle_Inventory_1617.csv',names=['日期','商品编号','库存'])

# 读取线上2016,2017的销售数据
#cat_list = pd.Series(['春季满帮鞋常青旧','春季满帮鞋畅销新','春季满帮鞋平销新','春季满帮鞋滞销旧','春季满帮鞋滞销新','春季浅口鞋常青旧','春季浅口鞋畅销新','春季浅口鞋平销新','春季浅口鞋滞销旧','春季浅口鞋滞销新','夏季纯凉鞋常青旧','夏季纯凉鞋畅销新','夏季纯凉鞋平销新','夏季纯凉鞋滞销旧','夏季纯凉鞋滞销新'])
cat_list = pd.Series(['春季满帮鞋滞销新'])
master = pd.read_csv(home_dir+'Belle_Master_1617_Spring_Summer.csv')
master['2017新旧'] = master['款号'].str.slice(-1) < '7'
master_all = pd.merge(master,prop,how='left',left_on='供应商款色编号',right_on='供应商款色编码',suffixes=['','_1']).drop(['商品编号_1','供应商款色编号','款号_1','销售季','三级分类_1','季节_1'], axis=1)

# 读入生意参谋-商品效果表，并与商品属性表整合
item_effect = pd.read_csv('E:/Belle/Sycm/items_effect_belle.csv')
item_effect['product_id'] = item_effect['title'].str.extract("([A-Za-z0-9]{8})", expand=False)
item_effect = item_effect[['itemUv', 'payRate', 'avgStayTime', 'addCartItemCnt', 'favBuyerCnt', 'select_date_begin', 'product_id']]
item_effect = item_effect[item_effect['product_id'].notnull()]
prop_dedup = prop.drop_duplicates('款号', keep='first')
sycm = pd.merge(item_effect, prop_dedup, how='left', left_on='product_id', right_on='款号', suffixes=['','_1']).drop('款号', axis=1)
tw_list_2017 = pd.date_range(start="2017-01-01",end="2018-1-1",freq='W-MON')
line = 1

for cat in cat_list:
    season = cat[0:2]
    class3 = cat[2:5]
    category = cat[5:7]
    if '销' in category:
        category = cat[5:6]
    ind_new = cat[7]
    master_all_filter = master_all[(master_all['季节'] == season) & (master_all['三级分类'] == class3) & (master_all['2017类别'] == category) & (master_all['2017新旧'] == (ind_new == '旧')) & (pd.to_datetime(master_all['日期']) >= pd.to_datetime('2016-12-31'))]
    for sku in master_all_filter['供应商款色编码'].unique():
        master_sku = master_all_filter[master_all_filter['供应商款色编码'] == sku]
        spu = master_sku['款号'].values[0]
        sycm_sku = sycm[(sycm['product_id'] == spu)]
        sale_cnt_2017 = {}
        dscnt_2017 = {}
        sale_amt_2017 = {}
        sold_out_2017 = {}
        inv_end_2017 = {}
        inv_start_2017 = {}
        for number in range(1,len(tw_list_2017)):
            line_number = number + 27
            line_number2 = number + 104
            start = tw_list_2017[number - 1]
            end = tw_list_2017[number]
            sale_cnt_2017[start] = np.sum(master_sku['成交件数'][(pd.to_datetime(master_sku['日期']) >= start) & (pd.to_datetime(master_sku['日期']) < end)])
            dscnt_2017[start] = np.average(master_sku['折扣率'][(pd.to_datetime(master_sku['日期']) >= start) & (pd.to_datetime(master_sku['日期']) < end)])*100
            if str(dscnt_2017[start]) == 'nan':
                dscnt_2017[start] = ""
            sale_amt_2017[start] = int(np.sum(master_sku['成交金额'][(pd.to_datetime(master_sku['日期']) >= start) & (pd.to_datetime(master_sku['日期']) < end)]))
            #累计售罄率 = 累计销量 / (累计销量 + 期末库存)
            sale_total = np.sum(master_sku['成交件数'][(pd.to_datetime(master_sku['日期']) >= pd.to_datetime('2017-01-01')) & (pd.to_datetime(master_sku['日期']) <= end)])
            inv_total = 0
            inv_total_2 = 0
            for pid in master_sku['商品编号'].unique():
                inv_pid = inv[inv['商品编号'] == pid].sort_values('日期')
                if sum(pd.to_datetime(inv_pid['日期']) <= end) > 0:
                    inv_total += inv_pid[pd.to_datetime(inv_pid['日期']) <= end]['库存'].values[-1]
                    if sum(pd.to_datetime(inv_pid['日期']) <= start) > 0:
                        inv_total_2 += inv_pid[pd.to_datetime(inv_pid['日期']) <= start]['库存'].values[-1]
            inv_end_2017[start] = inv_total
            inv_start_2017[start] = inv_total_2
            if (sale_total + inv_total) != 0:
                sold_out_2017[start] = sale_total / (sale_total + inv_total)
            else:
                sold_out_2017[start] = 0
            
        addcartitemcnt = []
        itemuv = []
        avgstaytime = []
        payrate = []
        today = pd.to_datetime('2017-12-31')        
        for i in range(12,0,-1):
            tw_start = today + pd.Timedelta('-'+str(i*7)+' days')
            tw_end = today + pd.Timedelta('-'+str(i*7-7)+' days')
            master_sku_tw = master_sku[(pd.to_datetime(master_sku['日期']) > pd.to_datetime(tw_start)) & (pd.to_datetime(master_sku['日期']) <= pd.to_datetime(tw_end))]
            sycm_filter_tw = sycm_sku[(pd.to_datetime(sycm_sku['select_date_begin']) > pd.to_datetime(tw_start)) & (pd.to_datetime(sycm_sku['select_date_begin']) <= pd.to_datetime(tw_end))]
            addcartitemcnt.append(np.sum(sycm_filter_tw['addCartItemCnt']))
            itemuv.append(np.sum(sycm_filter_tw['itemUv']))
            avgstaytime.append(round(np.average(sycm_filter_tw['avgStayTime']),1))
            if len(sycm_filter_tw) != 0:
                payrate.append(round(sum(master_sku_tw['成交件数']) /sum(sycm_filter_tw['itemUv'])*100,2))
            else:
                payrate.append(0)

        web_tw_dict = {'addcartitemcnt':addcartitemcnt, 'itemuv':itemuv, 'avgstaytime':avgstaytime, 'payrate':payrate}
        web_tw = pd.DataFrame(web_tw_dict)
        sht.range('A'+str(line)).value = sku
        sht.range('B'+str(line)).value = cat
        sht.range('A'+str(line+1),'J'+str(line+1)).value = ['2017周销量','2017期末累计售罄率','本期期末库存','本期期初库存','2017平均折扣','2017周销售额','访客数','成交转化率','平均停留时间','加购物车数']
        for i in range(0,52):
            sht.range('A'+str(line+2+i)).value = list(sale_cnt_2017.values())[i]
            sht.range('B'+str(line+2+i)).value = list(sold_out_2017.values())[i]
            sht.range('C'+str(line+2+i)).value = list(inv_end_2017.values())[i]
            sht.range('D'+str(line+2+i)).value = list(inv_start_2017.values())[i]
            sht.range('E'+str(line+2+i)).value = list(dscnt_2017.values())[i]
            sht.range('F'+str(line+2+i)).value = list(sale_amt_2017.values())[i]
        for i in range(0,7):
            sht.range('G'+str(line+2+i)).value = web_tw['itemuv'].fillna(0)[i]
            sht.range('H'+str(line+2+i)).value = str(web_tw['payrate'].fillna(0)[i]) + '%'
            sht.range('I'+str(line+2+i)).value = web_tw['avgstaytime'].fillna(0)[i]
            sht.range('J'+str(line+2+i)).value = web_tw['addcartitemcnt'].fillna(0)[i]
        line += 55