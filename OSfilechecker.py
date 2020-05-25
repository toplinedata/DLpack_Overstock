# -*- coding: utf-8 -*-
"""
Created on Thu Nov 28 02:54:57 2019

@author: Ossang
"""
import os
import time
from datetime import datetime

try:    
    #local
    #set up download script folder path
    script_dir = 'C:\\Users\\User\\Desktop\\0047Automate_Script\\DLpack_Overstock\\'
    
    #Folder of OS daily download script
    OS_ProductInfo_dir = 'C:\\Users\\User\\Desktop\\0047Automate_Script\\DLpack_Overstock\\'
    OS_ProductQnA_dir = 'C:\\Users\\User\\Desktop\\0047Automate_Script\\DLpack_Overstock\\'
    OS_Replenishment_dir  = 'C:\\Users\\User\\Desktop\\0047Automate_Script\\DLpack_Overstock\\'
    os.chdir(script_dir)
except: 
    #0047
    #set up download script folder path
    script_dir = 'C:\\Users\\raymond.hung\\Documents\\Automate_Script\\DLpack_Overstock\\'
    
    #Folder of OS daily download script
    OS_ProductInfo_dir = 'N:\\E Commerce\\Public Share\\Dot Com - Overstock\\Product Information\\'
    OS_ProductQnA_dir = 'N:\\E Commerce\\Public Share\\Dot Com - Overstock\\Q&A\\'
    OS_Replenishment_dir  = 'N:\\E Commerce\\Public Share\\Dot Com - Overstock\\Replenishment Report\\'
    os.chdir(script_dir)
    
# Today's date_label for OS daily download script
date_label = time.strftime('%Y%m%d')

# Check today's download file for OS daily download script
for i in range(3):
    count = 0
    file_list=list(set(os.listdir(OS_ProductInfo_dir)+os.listdir(OS_ProductQnA_dir) \
                    +os.listdir(OS_Replenishment_dir)))

    for file in file_list:
        if 'Product Infomation '+date_label in file:
            print("Product Infomation OK")
            count+=1
        elif 'Product Page Q & A '+date_label in file:
            print("Product Page Q&A OK")
            count+=1       
        elif 'Replenishment detail '+date_label in file:
            print("Replenishment detail OK")
            count+=1

    if count < 3:
        os.system("python "+script_dir+"DL_OS_ProductInfo.py")
        os.system("python "+script_dir+"DL_OS_ProductQA.py")
        os.system("python "+script_dir+"DL_OS_Replenishment_detail.py")
    else:
        break

