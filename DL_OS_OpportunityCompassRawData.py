# -*- coding: utf-8 -*-
"""
Created on Thu Oct 17 10:39:00 2019

@author: Ossang Ou
"""

import os
import shutil
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

# Today's date
date_label = time.strftime('%Y%m%d')

try:
    #local
    os.chdir('C:\\Users\\User\\Desktop\\0047Automate_Script\\DLpack_Overstock\\')
except:
    #0047
    os.chdir('N:\\E Commerce\\Public Share\\Dot Com - Overstock\\Q&A\\temp_DL_file\\')

#storage Directory
if 'Desktop\\0047Automate_Script' in os.getcwd():
    driver_path = 'C:\\Users\\User\\Anaconda3\\chrome\\chromedriver.exe'
    work_dir = 'C:\\Users\\User\\Desktop\\0047Automate_Script\\DLpack_Overstock\\'
    temp_dir = 'C:\\Users\\User\\Desktop\\0047Automate_Script\\DLpack_Overstock\\temp_DL_file\\'
    Download_dir = 'C:\\Users\\User\\Desktop\\0047Automate_Script\\DLpack_Overstock\\'
else:      
    driver_path = 'C:\\Users\\raymond.hung\\chrome\\chromedriver.exe'
    work_dir = 'C:\\Users\\raymond.hung\\Documents\\Automate_Script\\DLpack_Overstock\\'
    temp_dir = 'N:\E Commerce\Public Share\Dot Com - Overstock\Opportunity Compass\opportunity compass raw data\\'
    Download_dir = 'N:\E Commerce\Public Share\Dot Com - Overstock\Opportunity Compass\opportunity compass raw data\\'

os.chdir(temp_dir)

# Account and Password
login_info = pd.read_csv(work_dir+ 'Account & Password.csv',index_col=0)
username = login_info.loc['Account', 'CONTENT']
password = login_info.loc['Password', 'CONTENT']

# Chrome driver setting
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
prefs = {'download.default_directory' : temp_dir}
options.add_experimental_option("prefs", prefs)
driver = webdriver.Chrome(driver_path,chrome_options=options)

# Supplier Oasis website turn into login page
Supplier_Oasis = 'https://www.supplieroasis.com/Pages/default.aspx'
driver.get(Supplier_Oasis)
driver.find_element_by_link_text('Sign In').click()

LoadingChecker = (By.ID, 'ContentPlaceHolder1_SubmitButton')
WebDriverWait(driver, 120).until(EC.presence_of_element_located(LoadingChecker))
# Input username and password and login
driver.find_element_by_id('ContentPlaceHolder1_UsernameTextBox').send_keys(username)
driver.find_element_by_id('ContentPlaceHolder1_PasswordTextBox').send_keys(password)
driver.find_element_by_id('ContentPlaceHolder1_SubmitButton').click()
time.sleep(10)


# Turn to Opportunity Compass Page
LoadingChecker = (By.PARTIAL_LINK_TEXT, 'Opportunity Compass')
WebDriverWait(driver, 120).until(EC.presence_of_element_located(LoadingChecker))
driver.find_element_by_partial_link_text('Opportunity Compass').click()
time.sleep(3)

# Download Opportunity Compass Report
LoadingChecker = (By.CSS_SELECTOR, '.btn-primary')
WebDriverWait(driver, 120).until(EC.presence_of_element_located(LoadingChecker))
driver.find_element_by_css_selector('.btn-primary').click()
time.sleep(60)

driver.quit()

shutil.move('Opportunity_Compass_Snapshot.xlsx', Download_dir+'Opportunity_Compass_Snapshot_'+date_label+'.xlsx')