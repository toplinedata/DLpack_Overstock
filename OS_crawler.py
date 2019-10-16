# -*- coding: utf-8 -*-

import requests
from bs4 import BeautifulSoup
import numpy as np
import pandas as pd
import os
import datetime


Todate = datetime.datetime.strftime(datetime.date.today(), '%Y%m%d')

#local
#os.chdir('C:\\Users\\User\\Desktop\\Website Ranking\\')
#0047
os.chdir('C:\\Users\\raymond.hung\\OS_crawler\\')

if 'Desktop\\Website Ranking' in os.getcwd():
    driver_path = 'C:\\Users\\User\\Anaconda3\\chrome\\chromedriver.exe'
    work_dir = 'C:\\Users\\User\\Desktop\\Website Ranking\\'
    Download_dir = 'C:\\Users\\User\\Desktop\\Website Ranking\\'
    ROOT_DIR = 'C:\\Users\\User\\Desktop\\Website Ranking\\'
else:      
    driver_path = 'C:\\Users\\raymond.hung\\chrome\\chromedriver.exe'
    work_dir = 'C:\\Users\\raymond.hung\\Documents\\Automate_Script\\DLpack_Overstock\\'
    Download_dir = 'C:\\Users\\raymond.hung\\OS_crawler\\'
    ROOT_DIR = 'C:\\Users\\raymond.hung\\OS_crawler\\'
os.chdir(Download_dir)


#ROOT_DIR = 'C:\\Users\\raymond.hung\\OS_crawler\\'
STORE_DIR = os.path.join(ROOT_DIR, Todate)

if not os.path.exists(STORE_DIR):
    os.mkdir(STORE_DIR)

np_nan = np.empty([150, 21])
np_nan[:,:] = None

# Set the crawler index
cwl_f_type = ['Living-Room', 'Bedroom', 'Dining-Room', 'Kitchen', 'Patio', 'Home-Office', 'Kids']
cwl_i_num = 150
cwl_ind = 0

hd_os = {'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.102 Safari/537.36',
      'referrer': 'https://www.overstock.com/Home-Garden/Furniture/32/dept.html'}

#Furniture page
url = 'https://www.overstock.com/Home-Garden/Furniture/32/dept.html'

res = requests.get(url,headers=hd_os)
soup = BeautifulSoup(res.text, 'html.parser')

####type1
##find side Furniture Categories and create list
#side_list = soup.find(attrs={'class':'nl1-leftnav-ul'})
#sbg_bts = side_list.find_all('li', attrs={'class':'nl1-leftnav-li'})
#sbg_links = list(map(lambda x: x.find('a').get('href'), sbg_bts))
#
#cwl_links = []
#for ckw in cwl_f_type:
#    for sbg_link in sbg_links:
#        if ckw in sbg_link:
#            atlink_end = sbg_link.find('cat.html?')
#            cwl_link = 'https:{}cat.html?format=fusion&count={}'.format(sbg_link[:atlink_end], cwl_i_num)
#            cwl_links.append(cwl_link)

####type2
link_csv=pd.read_csv(ROOT_DIR+'OS Categories and subcategories.csv')
link_csv.dropna(inplace=True)
link_csv.reset_index(inplace=True)
cwl_f_type=link_csv['Topline SubCategory']
cwl_links=link_csv['Link']+'?format=fusion&count=150'
#cwl_link='https://www.overstock.com/Home-Garden/Dressers/2017/subcat.html?format=fusion&count=150' 
   


for cwl_link in cwl_links:
    colname = ['product_name', 'short_name', 'id', 'sku',
               'urls', 'image', 'favorite',
               'labeled_price', 'Lowest_price',  'Highest_price',
               'current_status', 'discounted_type', 'valueMessaging', 'clearance',
               'discounted_percent', 'discounted_price', 'onSaleExpiration',
               'reviews_count', 'rating', 'reviews_url', 
               'fastDelivery']
    
    Df_PdDetail = pd.DataFrame(np_nan, columns=colname, index=range(1, cwl_i_num+1))
    res = requests.get(cwl_link, headers=hd_os)
#    for i in range(cwl_i_num):
    for i in range(len(res.json()['products'])):
        pd_info = res.json()['products'][i]
        Df_PdDetail.loc[i+1, 'product_name'] = pd_info['name']
        Df_PdDetail.loc[i+1, 'short_name'] = pd_info['shortName']
        Df_PdDetail.loc[i+1, 'id'] = pd_info['id']
        Df_PdDetail.loc[i+1, 'sku'] = pd_info['sku']
        
        Df_PdDetail.loc[i+1, 'urls'] = pd_info['urls']['productPage']
        Df_PdDetail.loc[i+1, 'image'] = pd_info['urls']['image']
        Df_PdDetail.loc[i+1, 'favorite'] = pd_info['productFavCount']
        
        if pd_info['pricing']['comparison']:
            Df_PdDetail.loc[i+1, 'labeled_price'] = pd_info['pricing']['comparison']['price']
        Df_PdDetail.loc[i+1, 'Lowest_price'] = pd_info['pricing']['current']['priceBreakdown']['price']
        if pd_info['pricing']['current']['toPriceBreakdown']:
            Df_PdDetail.loc[i+1, 'Highest_price'] = pd_info['pricing']['current']['toPriceBreakdown']['price']
        
        Df_PdDetail.loc[i+1, 'current_status'] = pd_info['pricing']['current']['type']
        if pd_info['pricing']['discounted']:
            Df_PdDetail.loc[i+1, 'discounted_type'] = pd_info['pricing']['discounted']['type']
        Df_PdDetail.loc[i+1, 'valueMessaging'] = pd_info['valueMessaging']
        Df_PdDetail.loc[i+1, 'clearance'] = pd_info['clearance']
        if pd_info['pricing']['discounted']:
            Df_PdDetail.loc[i+1, 'discounted_percent'] = pd_info['pricing']['discounted']['percent']
        if pd_info['pricing']['discounted']:
            Df_PdDetail.loc[i+1, 'discounted_price'] = pd_info['pricing']['discounted']['price']
        Df_PdDetail.loc[i+1, 'onSaleExpiration'] = pd_info['onSaleExpiration']
        
        if pd_info['reviews']:
            Df_PdDetail.loc[i+1, 'reviews_count'] = pd_info['reviews']['count']
            Df_PdDetail.loc[i+1, 'rating'] = pd_info['reviews']['rating']
            Df_PdDetail.loc[i+1, 'reviews_url'] = pd_info['reviews']['url']
        
        Df_PdDetail.loc[i+1, 'fastDelivery'] = pd_info['fastDelivery']

#    catepage = cwl_f_type[cwl_ind].replace('-','')
    catepage = cwl_f_type[cwl_ind].replace('\n',',')
    catepage = cwl_f_type[cwl_ind].replace('/',',')
    Df_PdDetail.to_csv(os.path.join(STORE_DIR, catepage)+'.csv', index_label='rank')
    cwl_ind += 1
    
#Df_PdDetail.to_csv('test.csv', index_label='rank')
#res.json()
#with open("C:\\Users\\User\\Desktop\\Website Ranking\\test.json","wb")as f:
#    f.write(res.content)
#    f.close()
#print("成功寫入檔案")