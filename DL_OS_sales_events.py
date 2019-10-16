import csv
import time
import datetime
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
import os
import shutil

st_time = str(input('Download start time?(MM/DD/YYYY) From:'))
end_time = input('Download end time?(MM/DD/YYYY) To:')
driver_path = 'C:\\Users\\User\\Anaconda3\\chrome\\chromedriver.exe'
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
driver = webdriver.Chrome(driver_path,chrome_options=options)
# Overstock supplier website
OS_oasis = 'https://www.supplieroasis.com'
driver.get(OS_oasis)
time.sleep(1)

# Click login bottom, and turn to login in page
driver.find_element_by_link_text('Sign In').click()
time.sleep(1)
# Input username and password and login
username = driver.find_element_by_id('ContentPlaceHolder1_UsernameTextBox').send_keys('homelegance1')
password = driver.find_element_by_id('ContentPlaceHolder1_PasswordTextBox').send_keys('Overstock123!@#')
driver.find_element_by_id('ContentPlaceHolder1_SubmitButton').click()
time.sleep(1)

# Turn to promotions page
driver.find_element_by_link_text('Promotions').click()
time.sleep(5)


# Enter the start and the end of download time period
# Turn iframe to frame
driver.find_element_by_link_text('Sales Manager').click()
time.sleep(5)

promote_iframe = driver.find_element_by_id('reportFrame')
driver.switch_to.frame(promote_iframe)
time.sleep(15)
#select status 'Past'
s1 = Select(driver.find_element_by_xpath('/html/body/div[3]/div/div/div/div[3]/form/select[2]'))
s1.select_by_visible_text('Past')

# Input date
download_start = driver.find_element_by_id('start')
download_start.clear()
download_start.send_keys(st_time)
download_end = driver.find_element_by_id('end')
download_end.clear()
download_end.send_keys(end_time)
time.sleep(15)
# Click 'Search' bottom
driver.find_element_by_css_selector('body > div:nth-child(3) > div > div > div > div:nth-child(3) > form > button:nth-child(8)').click()
time.sleep(30)

# Set index to detect the end page of promote events
page_count = 1
event_in_page_max = driver.find_element_by_id('sales-events').get_attribute('outerHTML').count('<tr') - 1
event_in_page = event_in_page_max

while event_in_page == event_in_page_max:
    print('Crawling page' +str(page_count))
    # Access table source code
    table_element_code = driver.find_element_by_id('sales-events').get_attribute('outerHTML')
    # Count event in page
    event_in_page = table_element_code.count('<tr') - 1
    
    
    for i in range(1,event_in_page+1):
        event_name = driver.find_element_by_xpath('//*[@id="sales-events"]/tbody/tr['+str(i)+']/td[2]').text[11:]
        event_name = event_name.replace('!','').replace(',','').replace('.','').replace('?','').replace(':','').replace('(','').replace(')','').replace('\"','').replace('*','')
        event_start = driver.find_element_by_xpath('//*[@id="sales-events"]/tbody/tr['+str(i)+']/td[3]').text
        event_start_t = datetime.datetime.strptime(event_start, '%m/%d/%Y').strftime('%Y%m%d')
        event_start_c = datetime.datetime.strptime(event_start, '%m/%d/%Y').strftime('%Y/%m/%d')
        event_end = driver.find_element_by_xpath('//*[@id="sales-events"]/tbody/tr['+str(i)+']/td[4]').text
        event_end_t = datetime.datetime.strptime(event_end, '%m/%d/%Y').strftime('%Y%m%d')
        event_end_c = datetime.datetime.strptime(event_end, '%m/%d/%Y').strftime('%Y/%m/%d')
        event_filename = 'C:\\Users\\User\\Desktop\\0047Automate_Script\\DLpack_Overstock\\test\\' +event_start_t+ '_' + event_end_t+ '_' +event_name+ '.csv'
        
        
    
        if not os.path.isfile(event_filename):
            # Click and turn to the event page
            driver.find_element_by_xpath('//*[@id="sales-events"]/tbody/tr['+str(i)+']').click()                
            time.sleep(1)                
            #Click the download bottom
            try:
                driver.find_element_by_link_text('Download CSV').click()
                time.sleep(20)
                os.renames('C:\\Users\\User\\Downloads\\data.csv', event_filename)
                time.sleep(5)
               
            except NoSuchElementException:
                print(str(page_count)+ '-' +str(i)+ ', Skip')
                
            
            #Click the back bottom and turn back to the event list
            driver.find_element_by_link_text('Back').click()
            time.sleep(20)
                    
            if page_count > 1:
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                next_bottom = driver.find_element_by_link_text(str(page_count))
                next_bottom.click()
                time.sleep(3)
                driver.execute_script("window.scrollTo(0, 0);")
                        
            if i >= 4 and i < event_in_page:
                driver.execute_script('window.scrollBy(0,' +str(70*(i-4+1))+ ')')
            elif i < 4:
                continue
            elif i == event_in_page:
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            else:
                break
            
        else:
            print(str(page_count)+ '-' +str(i)+ ', ' +event_name+ '_' +event_start_t+ '_' +event_end_t+ '.csv is exist')
            driver.execute_script('window.scrollBy(0, 65)')

    # Next page
    if event_in_page == event_in_page_max:
        page_count += 1
        next_bottom = driver.find_element_by_link_text(str(page_count))
        next_bottom.click()
        time.sleep(5)
        driver.execute_script("window.scrollTo(0, 0);")

driver.quit()
