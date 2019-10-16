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
#time.sleep(30)

for i in range(3):
    #Turn to "Product Information" sheet
    try:
        LoadingChecker = (By.XPATH, '//*[@id="mstr90"]/div/table/tbody/tr/td[2]/input')#maybe change89 90
        WebDriverWait(driver,60).until(EC.element_to_be_clickable(LoadingChecker))
        time.sleep(10)
        driver.find_element_by_css_selector('#mstr90 > div > table > tbody > tr > td:nth-child(2) > input').click()
    except:
        LoadingChecker = (By.XPATH, '//*[@id="mstr91"]/div/table/tbody/tr/td[2]/input')#maybe change89 90
        WebDriverWait(driver, 60).until(EC.element_to_be_clickable(LoadingChecker))
        time.sleep(10)
        driver.find_element_by_css_selector('#mstr91 > div > table > tbody > tr > td:nth-child(2) > input').click()

    time.sleep(10)
    
#    try:
#        #check dropdown menus ready
#        LoadingChecker = (By.XPATH, '//*[@id="mstr119"]')#maybe change119 120
#        WebDriverWait(driver, 30).until(EC.presence_of_element_located(LoadingChecker))
#        # Find inner widow element and scroll to most right side
#        inner = driver.find_element_by_xpath('//*[@id="mstr80_scrollboxNode"]')
#        driver.execute_script('arguments[0].scrollLeft = arguments[0].scrollWidth', inner)
#        time.sleep(10)
#    except TimeoutException:
#        print("timeout")
#        driver.refresh()
#        continue

    # Click dropdown menus and download excel file
    try:
        if driver.find_element_by_xpath('//*[@id="mstr119"]').get_attribute("onclick"):
            driver.find_element_by_xpath('//*[@id="mstr119"]').click()
            time.sleep(3)
            driver.find_element_by_xpath('//*[@class="mstrmojo-ContextMenuItem cmd4"]').click()
            time.sleep(3)
            
        elif driver.find_element_by_xpath('//*[@id="mstr120"]').get_attribute("onclick"):
            driver.find_element_by_xpath('//*[@id="mstr120"]').click()
            time.sleep(3)
            driver.find_element_by_xpath('//*[@class="mstrmojo-ContextMenuItem cmd4"]').click()
            time.sleep(3)

        time.sleep(20)
        os.rename('Product Dashboard.xlsx', 'Product Infomation '+date_label+'.xlsx')
        break
    except FileExistsError:
        print('FileExistsError')
        break
    except:
        print('fail to download excel file')
        driver.refresh()
        
driver.quit()