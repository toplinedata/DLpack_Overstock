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
driver.find_element_by_link_text('Sign In').click()

LoadingChecker = (By.ID, 'ContentPlaceHolder1_SubmitButton')
WebDriverWait(driver, 120).until(EC.presence_of_element_located(LoadingChecker))
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
WebDriverWait(driver, 120).until(EC.presence_of_element_located(LoadingChecker))
# Scroll to bottum and get the Product Dashboard href
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
driver.get(driver.find_element_by_partial_link_text('PRODUCT DASHBOARD').get_attribute('href'))

LoadingChecker = (By.XPATH, '//*[@k="W858"]')
WebDriverWait(driver, 120).until(EC.presence_of_element_located(LoadingChecker))
#Find inner widow element and scroll to most right side
#inner = driver.find_element_by_xpath('//*[@id="mstr80_scrollboxNode"]')
#driver.execute_script('arguments[0].scrollLeft = arguments[0].scrollWidth', inner)
#time.sleep(3)

# Click Product Questions & Answers
driver.find_element_by_xpath('//*[@k="W858"]').click()
time.sleep(10)

# Shift to Q&A tab
QA_page = driver.window_handles[-1]
driver.switch_to_window(QA_page)


# Click dropdown menus and download excel file
#try:
#    driver.find_element_by_xpath('//*[@id="mstr136"]/div').click()
#    time.sleep(3)
#    driver.find_element_by_xpath('//*[@class="mstrmojo-ContextMenuItem cmd4"]').click()
#except:
#    driver.find_element_by_xpath('//*[@id="mstr135"]/div').click()
#    time.sleep(3)
#    driver.find_element_by_xpath('//*[@class="mstrmojo-ContextMenuItem cmd4"]').click()
#    
try: #select and click
    LoadingChecker = (By.CSS_SELECTOR, '.mstrmojo-HBox-cell.mstrmojo-ToolBar-cell')
    WebDriverWait(driver, 300).until(EC.element_to_be_clickable(LoadingChecker))
    driver.find_elements_by_css_selector('.mstrmojo-HBox-cell.mstrmojo-ToolBar-cell')[0].click()
    driver.find_elements_by_css_selector('.mstrmojo-CMI-text')[1].click()
# Click dropdown menus and download excel file
except:
    LoadingChecker = (By.CLASS_NAME, 'tbDown')
    WebDriverWait(driver, 300).until(EC.element_to_be_clickable(LoadingChecker))
    driver.find_elements_by_class_name('tbDown')[0].click()
    driver.find_elements_by_class_name('cmd4')[0].click()
    
time.sleep(30)

driver.quit()

shutil.move('Product Page Q & A.xlsx', Download_dir+'Product Page Q & A '+date_label+'.xlsx')