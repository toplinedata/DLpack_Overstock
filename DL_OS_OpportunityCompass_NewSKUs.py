# -*- coding: utf-8 -*-
"""
Created on Fri Oct 30 14:00:06 2020

@author: User
"""

import os
import zipfile
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

# Today's date
date_label = time.strftime('%Y_%m_%d')

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
    temp_dir = 'N:\\E Commerce\\Public Share\\Dot Com - Overstock\\Opportunity Compass\\opportunity compass raw data\\'
    Download_dir = 'N:\\E Commerce\\Public Share\\Dot Com - Overstock\\Opportunity Compass\\opportunity compass raw data\\New SKUs\\'

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

# Wait for click Sign In button
for i in range(5):
    try:
        LoadingChecker = (By.LINK_TEXT, 'Sign In')
        WebDriverWait(driver, 120).until(EC.presence_of_element_located(LoadingChecker))
        driver.find_element_by_link_text('Sign In').click()
        break
    except:
        driver.refresh()
        if i == 2:
            driver.quit()
            driver = webdriver.Chrome(driver_path,chrome_options=options)
            Supplier_Oasis = 'https://www.supplieroasis.com/Pages/default.aspx'
            driver.get(Supplier_Oasis)

for i in range(5):
    try:
        LoadingChecker = (By.ID, 'ContentPlaceHolder1_SubmitButton')
        WebDriverWait(driver, 120).until(EC.presence_of_element_located(LoadingChecker))
        # Input username and password and login
        driver.find_element_by_id('ContentPlaceHolder1_UsernameTextBox').send_keys(username)
        driver.find_element_by_id('ContentPlaceHolder1_PasswordTextBox').send_keys(password)
        driver.find_element_by_id('ContentPlaceHolder1_SubmitButton').click()
        time.sleep(5)
        break
    except:
        driver.refresh()

# Turn to Opportunity Compass page
for i in range(5):
    try:
        LoadingChecker = (By.PARTIAL_LINK_TEXT, 'Opportunity Compass')
        WebDriverWait(driver, 30).until(EC.presence_of_element_located(LoadingChecker))
        driver.get(driver.find_element_by_partial_link_text('Opportunity Compass').get_attribute('href'))
        time.sleep(10)
        break
    except:
        driver.refresh()

# Click NEW SKUS Tab
for i in range(5):
    try:
        LoadingChecker = (By.ID, 'NEW_SKUS')
        WebDriverWait(driver, 30).until(EC.presence_of_element_located(LoadingChecker))
        driver.find_element_by_id('NEW_SKUS').click()
        time.sleep(10)
        break
    except:
        driver.refresh()

# Export Report
if driver.find_element_by_css_selector('#exportButton').text == "Export All New SKUs":
    driver.find_element_by_css_selector('#exportButton').click()

time.sleep(30)

driver.quit()

fz = zipfile.ZipFile('Opportunity-Compass-NEW_SKUS-'+date_label+'.zip', 'r')
for file in fz.namelist():
    fz.extract(file, Download_dir)
    fz.close()

if os.path.exists(Download_dir+file):
    os.remove('Opportunity-Compass-NEW_SKUS-'+date_label+'.zip')
