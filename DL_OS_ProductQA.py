# -*- coding: utf-8 -*-
"""
Created on Wed Jul 25 01:20:58 2018

@author: Chieh-Hsu Yang
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
    temp_dir = 'N:\\E Commerce\\Public Share\\Dot Com - Overstock\\Q&A\\temp_DL_file\\'
    Download_dir = 'N:\\E Commerce\\Public Share\\Dot Com - Overstock\\Q&A\\'

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

for i in range(5):
    try:
        # Turn to Report page
        driver.get('https://edge.supplieroasis.com/partner-reports')
        
        LoadingChecker = (By.PARTIAL_LINK_TEXT, 'Product Dashboard')
        WebDriverWait(driver, 60).until(EC.presence_of_element_located(LoadingChecker))
        # Scroll to bottum and get the Product Dashboard href
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        driver.get(driver.find_element_by_partial_link_text('Product Dashboard').get_attribute('href'))
        time.sleep(30)
        break

    except:
        driver.refresh()
        
for i in range(5):
    try:
        # Click Product Questions & Answers
        LoadingChecker = (By.XPATH, '//*[@k="W858"]')
        WebDriverWait(driver, 120).until(EC.presence_of_element_located(LoadingChecker))
        driver.find_element_by_xpath('//*[@k="W858"]').click()
        time.sleep(10)
        break
    except:
        driver.refresh()
        time.sleep(10)

# Shift to Q&A tab
QA_page = driver.window_handles[-1]
driver.switch_to_window(QA_page)

# Click dropdown menus then download excel file
try: 
    LoadingChecker = (By.CSS_SELECTOR, '.mstrmojo-HBox-cell.mstrmojo-ToolBar-cell')
    WebDriverWait(driver, 120).until(EC.element_to_be_clickable(LoadingChecker))
    driver.find_elements_by_css_selector('.mstrmojo-HBox-cell.mstrmojo-ToolBar-cell')[0].click()
    time.sleep(5)
    driver.find_elements_by_css_selector('.mstrmojo-CMI-text')[1].click()
except:
    LoadingChecker = (By.CLASS_NAME, 'tbDown')
    WebDriverWait(driver, 120).until(EC.element_to_be_clickable(LoadingChecker))
    driver.find_elements_by_class_name('tbDown')[0].click()
    time.sleep(5)
    driver.find_elements_by_class_name('cmd4')[0].click()
    
time.sleep(30)

driver.quit()

shutil.move('Product Page Q & A.xlsx', Download_dir+'Product Page Q & A '+date_label+'.xlsx')
