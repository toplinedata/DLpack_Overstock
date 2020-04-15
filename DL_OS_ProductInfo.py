# -*- coding: utf-8 -*-
"""
Created on Tue Jul 24 21:53:04 2018

@author: Chieh-Hsu Yang
"""

import os
import time
import shutil
import pandas as pd
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException

# Today's date
date_label = time.strftime('%Y%m%d')
try:
    #local
    os.chdir('C:\\Users\\User\\Desktop\\0047Automate_Script\\DLpack_Overstock\\')
except:
    #0047
    os.chdir('N:\\E Commerce\\Public Share\\Dot Com - Overstock\\Product Information\\')

#storage Directory
if 'Desktop\\0047Automate_Script' in os.getcwd():
    driver_path = 'C:\\Users\\User\\Anaconda3\\chrome\\chromedriver.exe'
    work_dir = 'C:\\Users\\User\\Desktop\\0047Automate_Script\\DLpack_Overstock\\'
    Download_dir = 'C:\\Users\\User\\Desktop\\0047Automate_Script\\DLpack_Overstock\\'
else:      
    driver_path = 'C:\\Users\\raymond.hung\\chrome\\chromedriver.exe'
    work_dir = 'C:\\Users\\raymond.hung\\Documents\\Automate_Script\\DLpack_Overstock\\'
    Download_dir = 'N:\\E Commerce\\Public Share\\Dot Com - Overstock\\Product Information\\'


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
    
for i in range(3):
    try:
        # Supplier Oasis website turn into login page
        Supplier_Oasis = 'https://www.supplieroasis.com/Pages/default.aspx'
        driver.get(Supplier_Oasis)
        
        LoadingChecker = (By.LINK_TEXT, 'Sign In')
        WebDriverWait(driver,60).until(EC.element_to_be_clickable(LoadingChecker))
        driver.find_element_by_link_text('Sign In').click()
        break
    except:
        driver.refresh()

for i in range(3):
    try:
        LoadingChecker = (By.ID, 'ContentPlaceHolder1_SubmitButton')
        WebDriverWait(driver, 30).until(EC.presence_of_element_located(LoadingChecker))
        # Input username and password and login
        driver.find_element_by_id('ContentPlaceHolder1_UsernameTextBox').send_keys(username)
        driver.find_element_by_id('ContentPlaceHolder1_PasswordTextBox').send_keys(password)
        driver.find_element_by_id('ContentPlaceHolder1_SubmitButton').click()
        time.sleep(5)
        break
    except:
        driver.refresh()
        
for j in range(3):
    for i in range(3):
        try:
            # Turn to Report page
            driver.get('https://edge.supplieroasis.com/reporting')
            ## Find Side Menu element and use execute java script move
            #sidemenu = driver.find_element_by_xpath('//*[@id="sofs-header"]/menu/div/div[3]/left-menu/ul')
            #driver.execute_script('arguments[0].scrollTop = arguments[0].scrollHeight', sidemenu)
            ## Click Report and 
            #driver.find_element_by_xpath('//*[@menu-item="REPORTS"]').click()
            
            LoadingChecker = (By.PARTIAL_LINK_TEXT, 'PRODUCT DASHBOARD')
            WebDriverWait(driver, 60).until(EC.presence_of_element_located(LoadingChecker))
            
            # Scroll to bottum and get the Product Dashboard href
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            driver.get(driver.find_element_by_partial_link_text('PRODUCT DASHBOARD').get_attribute('href'))
            time.sleep(30)
            break
        except:
            driver.refresh()
    
    for i in range(3):
        try:
            #Turn to "Product Information" sheet
            css = 'div.mstrmojo-DocLayoutViewer-layout > div > div > div:nth-child(3) > div > div:nth-child(3) > div.content > div > div > table > tbody > tr > td:nth-child(2) > input'
            LoadingChecker = (By.CSS_SELECTOR, css)
            WebDriverWait(driver,60).until(EC.element_to_be_clickable(LoadingChecker))
            driver.find_element_by_css_selector(css).click()
            break
        except:
            driver.refresh()
    
    time.sleep(10)
        
    # Click dropdown menus and download excel file
    for i in range(3):
        try:
            css = 'div.mstrmojo-DocLayoutViewer-layout > div > div > div:nth-child(3) > div > div:nth-child(1) > div > div:nth-child(2) > div > div:nth-child(1) > div:nth-child(2) > div:nth-child(1) > table > tbody > tr > td.mstrmojo-portlet-slot-toolbar.grid-menu > table > tbody > tr > td > table > tbody > tr > td:nth-child(1) > div'
            LoadingChecker = (By.CSS_SELECTOR, css)
            WebDriverWait(driver,60).until(EC.element_to_be_clickable(LoadingChecker))
            driver.find_element_by_css_selector(css).click()
            time.sleep(3)
            css = 'table.mstrmojo-itemwrap-table > tbody:nth-child(2) > tr > td.mstrmojo-CMI-text'
            LoadingChecker = (By.CSS_SELECTOR, css)
            WebDriverWait(driver,60).until(EC.element_to_be_clickable(LoadingChecker))
            driver.find_element_by_css_selector(css).click()
            break
        except:
            driver.refresh()
    
    time.sleep(40)
    
    if Path(Download_dir + 'Product Dashboard.xlsx').exists():
        break
    
shutil.move('Product Dashboard.xlsx', 'Product Infomation '+date_label+'.xlsx')
        
driver.quit()
