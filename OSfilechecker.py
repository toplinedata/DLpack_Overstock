# -*- coding: utf-8 -*-
"""
Created on Thu Nov 28 02:54:57 2019

@author: Ossang
"""
import os
import time
import subprocess

try:    
    #local
    #set up download script folder path
    script_dir = 'C:\\Users\\User\\Desktop\\0047Automate_Script\\DLpack_Overstock\\'
    os.chdir(script_dir)
    
    #Folder of OS daily download script
    OS_ProductInfo_dir = 'C:\\Users\\User\\Desktop\\0047Automate_Script\\DLpack_Overstock\\'
    OS_ProductQnA_dir = 'C:\\Users\\User\\Desktop\\0047Automate_Script\\DLpack_Overstock\\'
    OS_Replenishment_dir  = 'C:\\Users\\User\\Desktop\\0047Automate_Script\\DLpack_Overstock\\'
    OS_ReturnReport_dir  = 'C:\\Users\\User\\Desktop\\0047Automate_Script\\DLpack_Overstock\\'
    
except: 
    #0047
    #set up download script folder path
    script_dir = 'C:\\Users\\raymond.hung\\Documents\\Automate_Script\\DLpack_Overstock\\'
    os.chdir(script_dir)
    
    #Folder of OS daily download script
    OS_ProductInfo_dir = 'N:\\E Commerce\\Public Share\\Dot Com - Overstock\\Product Information\\'
    OS_ProductQnA_dir = 'N:\\E Commerce\\Public Share\\Dot Com - Overstock\\Q&A\\'
    OS_Replenishment_dir  = 'N:\\E Commerce\\Public Share\\Dot Com - Overstock\\Replenishment Report\\'
    OS_ReturnReport_dir  = 'N:\\E Commerce\\Public Share\\Dot Com - Overstock\\return details\\'
    
# Today's date_label for OS daily download script
date_label = time.strftime('%Y%m%d')
checker01 = 0
checker02 = 0
checker03 = 0
checker04 = 0

# Check today's download file for OS daily download script
while 1==1:
    file_list=list(set(os.listdir(OS_ProductInfo_dir)+os.listdir(OS_ProductQnA_dir) \
                      +os.listdir(OS_Replenishment_dir)+os.listdir(OS_ReturnReport_dir)))

    for file in file_list:
        if 'Product Infomation '+date_label in file:
            print("Product Infomation OK")
            checker01 = 1
        elif 'Product Page Q & A '+date_label in file:
            print("Product Page Q&A OK")
            checker02 = 1
        elif 'Replenishment detail '+date_label in file:
            print("Replenishment detail OK")
            checker03 = 1
        elif 'Returns Dashboard '+date_label in file:
            print("Return Report detail OK")
            checker04 = 1

    if checker01 != 1:
        subprocess.call("python "+script_dir+"DL_OS_ProductInfo.py")
    elif checker02 != 1:
        subprocess.call("python "+script_dir+"DL_OS_ProductQA.py")
    elif checker03 != 1:
        subprocess.call("python "+script_dir+"DL_OS_Replenishment_detail.py")
    elif checker04 != 1:
        subprocess.call("python "+script_dir+"DL_OS_ReturnReport.py")
    else:
        break
