# -*- coding: utf-8 -*-
"""
Created on Wed Jul 25 02:47:26 2018

@author: Chieh-Hsu Yang
"""

import os
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
    os.chdir('N:\\E Commerce\\Public Share\\Dot Com - Overstock\\')

#storage Directory
if 'Desktop\\0047Automate_Script' in os.getcwd():
    driver_path = 'C:\\Users\\User\\Anaconda3\\chrome\\chromedriver.exe'
    work_dir = 'C:\\Users\\User\\Desktop\\0047Automate_Script\\DLpack_Overstock\\'
    Download_dir = 'C:\\Users\\User\\Desktop\\0047Automate_Script\\DLpack_Overstock\\'
else:      
    driver_path = 'C:\\Users\\raymond.hung\\chrome\\chromedriver.exe'
    work_dir = 'C:\\Users\\raymond.hung\\Documents\\Automate_Script\\DLpack_Overstock\\'
    Download_dir = 'N:\\E Commerce\\Public Share\\Dot Com - Overstock\\Page Rank\\'
os.chdir(Download_dir)

# Account and Password
login_info = pd.read_csv(work_dir+ 'Account & Password.csv',index_col=0)
username = login_info.loc['Account', 'CONTENT']
password = login_info.loc['Password', 'CONTENT']

# Chrome driver setting
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
prefs = {'download.default_directory' : Download_dir}
options.add_experimental_option("prefs", prefs)
driver = webdriver.Chrome(driver_path,chrome_options=options)

# Supplier Oasis website turn into login page
Supplier_Oasis = 'https://www.supplieroasis.com/Pages/default.aspx'
driver.get(Supplier_Oasis)
driver.find_element_by_link_text('Sign In').click()

# Input username and password and login
LoadingChecker = (By.ID, 'ContentPlaceHolder1_SubmitButton')
WebDriverWait(driver, 120).until(EC.presence_of_element_located(LoadingChecker))
driver.find_element_by_id('ContentPlaceHolder1_UsernameTextBox').send_keys(username)
driver.find_element_by_id('ContentPlaceHolder1_PasswordTextBox').send_keys(password)
driver.find_element_by_id('ContentPlaceHolder1_SubmitButton').click()#error
time.sleep(5)

# Turn to Compass Accrual page
driver.get('https://edge.supplieroasis.com/improve/#/')
#download Compass Accrual report
driver.get(driver.find_element_by_xpath('//*[@id="header-export-btn"]').get_attribute('href'))
time.sleep(60)

#close web page
driver.quit()    

