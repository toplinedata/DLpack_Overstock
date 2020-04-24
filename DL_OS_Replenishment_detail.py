# -*- coding: utf-8 -*-
"""
Created on Wed Jul 25 02:47:26 2018

@author: Chieh-Hsu Yang
"""

import os
import time
import shutil
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
 os.chdir('N:\\E Commerce\\Public Share\\Dot Com - Overstock\\Replenishment Report\\')

#storage Directory
if 'Desktop\\0047Automate_Script' in os.getcwd():
    driver_path = 'C:\\Users\\User\\Anaconda3\\chrome\\chromedriver.exe'
    work_dir = 'C:\\Users\\User\\Desktop\\0047Automate_Script\\DLpack_Overstock\\'
    Download_dir = 'C:\\Users\\User\\Desktop\\0047Automate_Script\\DLpack_Overstock\\'
else:      
    driver_path = 'C:\\Users\\raymond.hung\\chrome\\chromedriver.exe'
    work_dir = 'C:\\Users\\raymond.hung\\Documents\\Automate_Script\\DLpack_Overstock\\'
    Download_dir = 'N:\\E Commerce\\Public Share\\Dot Com - Overstock\\Replenishment Report\\'
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
LoadingChecker = (By.LINK_TEXT, 'Sign In')
WebDriverWait(driver, 120).until(EC.presence_of_element_located(LoadingChecker))
driver.find_element_by_link_text('Sign In').click()

# Input username and password and login
for _ in range(3):
    try:
        LoadingChecker = (By.ID, 'ContentPlaceHolder1_SubmitButton')
        WebDriverWait(driver, 120).until(EC.presence_of_element_located(LoadingChecker))
        driver.find_element_by_id('ContentPlaceHolder1_UsernameTextBox').send_keys(username)
        driver.find_element_by_id('ContentPlaceHolder1_PasswordTextBox').send_keys(password)
        driver.find_element_by_id('ContentPlaceHolder1_SubmitButton').click()
        time.sleep(5)
    except:
        driver.refresh()

# Turn to Report page
driver.get('https://edge.supplieroasis.com/reporting')

# check download existed or not
if os.path.exists('Replenishment Dashboard.xlsx'):
    os.remove('Replenishment Dashboard.xlsx')

# Scroll to bottum and get the Inventory Dashboard href
for _ in range(3):
    try:
        LoadingChecker = (By.PARTIAL_LINK_TEXT, 'REPLENISHMENT DASHBOARD')
        WebDriverWait(driver, 120).until(EC.presence_of_element_located(LoadingChecker))
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        # Click Product Replenishment
        driver.get(driver.find_element_by_partial_link_text('REPLENISHMENT DASHBOARD').get_attribute('href'))
    except:
        driver.refresh()
    
# Click dropdown menus and download excel file
try:
    LoadingChecker = (By.CSS_SELECTOR, '.mstrmojo-HBox-cell.mstrmojo-ToolBar-cell')
    WebDriverWait(driver, 30).until(EC.element_to_be_clickable(LoadingChecker))
    driver.find_elements_by_css_selector('.mstrmojo-HBox-cell.mstrmojo-ToolBar-cell')[0].click()
    LoadingChecker = (By.CSS_SELECTOR, '.mstrmojo-CMI-text')
    WebDriverWait(driver, 30).until(EC.element_to_be_clickable(LoadingChecker))
    driver.find_elements_by_css_selector('.mstrmojo-CMI-text')[1].click()

# Click Product Replenishment
#LoadingChecker = (By.XPATH, '//*[@id="mstr116"]/div/table/tbody/tr/td[3]/input')
#WebDriverWait(driver, 900).until(EC.element_to_be_clickable(LoadingChecker))
#time.sleep(5)
#driver.find_element_by_xpath('//*[@id="mstr116"]/div/table/tbody/tr/td[3]/input').click()

# Find inner widow element and scroll to most right side
#LoadingChecker = (By.XPATH, '//*[@id="mstr202"]')
#WebDriverWait(driver, 900).until(EC.element_to_be_clickable(LoadingChecker))
#inner = driver.find_element_by_xpath('//*[@id="mstr80_scrollboxNode"]')
#driver.execute_script('arguments[0].scrollLeft = arguments[0].scrollWidth', inner)
#time.sleep(1)

except:
    LoadingChecker = (By.CLASS_NAME, 'tbDown')
    WebDriverWait(driver, 30).until(EC.element_to_be_clickable(LoadingChecker))
    driver.find_elements_by_class_name('tbDown')[0].click()
    LoadingChecker = (By.CLASS_NAME, 'cmd4')
    WebDriverWait(driver, 30).until(EC.element_to_be_clickable(LoadingChecker))
    driver.find_elements_by_class_name('cmd4')[0].click()
time.sleep(60)

driver.quit()

shutil.move('Replenishment Dashboard.xlsx', 'Replenishment detail '+date_label+'.xlsx')
