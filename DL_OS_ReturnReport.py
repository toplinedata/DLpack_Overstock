# -*- coding: utf-8 -*-
"""
Created on Tue Jul 24 21:53:04 2018

@author: Chieh-Hsu Yang
"""

import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
# from selenium.common.exceptions import TimeoutException

# Today's date
date_label = time.strftime('%Y%m%d')
try:
    #local
    os.chdir('C:\\Users\\User\\Desktop\\0047Automate_Script\\DLpack_Overstock\\')
except:
    #0047
    os.chdir('N:\\E Commerce\\Public Share\\Dot Com - Overstock\\return details\\')

#storage Directory
if 'Desktop\\0047Automate_Script' in os.getcwd():
    driver_path = 'C:\\Users\\User\\Anaconda3\\chrome\\chromedriver.exe'
    work_dir = 'C:\\Users\\User\\Desktop\\0047Automate_Script\\DLpack_Overstock\\'
    Download_dir = 'C:\\Users\\User\\Desktop\\0047Automate_Script\\DLpack_Overstock\\'
else:      
    driver_path = 'C:\\Users\\raymond.hung\\chrome\\chromedriver.exe'
    work_dir = 'C:\\Users\\raymond.hung\\Documents\\Automate_Script\\DLpack_Overstock\\'
    Download_dir = 'N:\\E Commerce\\Public Share\\Dot Com - Overstock\\return details\\'


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

LoadingChecker = (By.ID, 'ContentPlaceHolder1_SubmitButton')
WebDriverWait(driver, 30).until(EC.presence_of_element_located(LoadingChecker))
# Input username and password and login
driver.find_element_by_id('ContentPlaceHolder1_UsernameTextBox').send_keys(username)
driver.find_element_by_id('ContentPlaceHolder1_PasswordTextBox').send_keys(password)
driver.find_element_by_id('ContentPlaceHolder1_SubmitButton').click()
time.sleep(5)

# Turn to Report page
driver.get('https://edge.supplieroasis.com/reporting')


LoadingChecker = (By.PARTIAL_LINK_TEXT, 'RETURNS DASHBOARD')
WebDriverWait(driver, 60).until(EC.presence_of_element_located(LoadingChecker))

# Scroll to bottum and get the Product Dashboard href
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
driver.get(driver.find_element_by_partial_link_text('RETURNS DASHBOARD').get_attribute('href'))
#time.sleep(30)

for i in range(3):
    #Turn to "Details by Order" sheet
    try:
        LoadingChecker = (By.XPATH, '//*[@id="mstr119"]/div/table/tbody/tr/td[5]/input')
        WebDriverWait(driver,60).until(EC.element_to_be_clickable(LoadingChecker))
        time.sleep(10)
        driver.find_element_by_xpath('//*[@id="mstr119"]/div/table/tbody/tr/td[5]/input').click()
    except:
        LoadingChecker = (By.CSS_SELECTOR, '#mstr122 > div > table > tbody > tr > td:nth-child(5) > input')
        WebDriverWait(driver, 60).until(EC.element_to_be_clickable(LoadingChecker))
        time.sleep(10)
        driver.find_element_by_css_selector('#mstr122 > div > table > tbody > tr > td:nth-child(5) > input').click()

#    time.sleep(100)
    

    # Click dropdown menus and download excel file    
#    try:
    if os.path.exists(Download_dir+'Returns Dashboard.xlsx'):
        os.remove(Download_dir+'Returns Dashboard.xlsx')
        
    try:
        try:
    #            driver.find_element_by_xpath('//*[@id="mstr217"]/div').click()
            LoadingChecker = (By.CSS_SELECTOR, 'table.mstrmojo-ToolBar > tbody > tr > td > table > tbody > tr > td:nth-child(1) > div')
            WebDriverWait(driver, 100).until(EC.presence_of_element_located(LoadingChecker))
            driver.find_element_by_css_selector('table.mstrmojo-ToolBar > tbody > tr > td > table > tbody > tr > td:nth-child(1) > div').click()
            time.sleep(3)
    #            driver.find_element_by_xpath('//*[@id="mstr244"]/td[2]').click()
            LoadingChecker = (By.CSS_SELECTOR, 'body > div.mstrmojo-ContextMenu > div.mstrmojo-CM-itemsContainer > table > tbody:nth-child(2) > tr > td.mstrmojo-CMI-text')
            WebDriverWait(driver, 100).until(EC.presence_of_element_located(LoadingChecker))
            driver.find_element_by_css_selector('body > div.mstrmojo-ContextMenu > div.mstrmojo-CM-itemsContainer > table > tbody:nth-child(2) > tr > td.mstrmojo-CMI-text').click()
            time.sleep(3)
        except:
            driver.find_element_by_xpath('//*[@id="mstr213"]/div').click()
            time.sleep(3)
            driver.find_element_by_xpath('#mstr241 > td.mstrmojo-CMI-text').click()
            time.sleep(3)
    
        time.sleep(100)
        os.rename('Returns Dashboard.xlsx', 'Returns Dashboard '+date_label+'.xlsx')
        break
    
#    except FileExistsError:
#        print('FileExistsError')
#        break
    except Exception as e:
        print(e)
        if os.path.exists(Download_dir+'Returns Dashboard '+date_label+'.xlsx'):
            os.remove(Download_dir+'Returns Dashboard '+date_label+'.xlsx')
        driver.refresh()
        
driver.quit()
