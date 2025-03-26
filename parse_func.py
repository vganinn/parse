#!pip install selenium
#!pip install webdriver_manager

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException
from selenium.webdriver.support.ui import WebDriverWait
import time
import pandas as pd
from datetime import datetime
import os

def parse_data(banks_table: str, form_cbr: str, start_date: str, end_date: str, step: int, notify: bool = False):
    
    start_time = time.time()
    
    output_df = pd.read_excel(banks_table)
    output_df['ogrn'] = output_df['ogrn'].astype(str)
    
    banks = pd.read_excel(banks_table)
    banks['ogrn'] = banks['ogrn'].astype(str)
    
    output_file = 'BANK_DATA_' + start_date + '_' + end_date + '.xlsx'
    
    if os.path.exists(output_file):
            os.remove(output_file)
    
    search_query = 'https://cbr.ru/finorg/'
    service = Service(executable_path='C:/chromedriver/chromedriver.exe')
    driver = webdriver.Chrome(service=service)
    driver.get(search_query)
    
    if step == 1:
        years = int(end_date.split('-')[0]) - int(start_date.split('-')[0])
        months = int(end_date.split('-')[1]) - int(start_date.split('-')[1])
        mon_len = years * 12 + months + 1
        
    elif step == 3:
        years = int(end_date[:4]) - int(start_date[:4])
        months = int(end_date[4:]) - int(start_date[4:])
        mon_len = int(years * 5 + months / 3)
    
    if form_cbr == '123':
        xpth = '//*[@id="content"]/div/div/div/div[3]/div[2]/table/tbody/tr[2]/td[3]'
        
    elif form_cbr == '802':
        xpth = '//*[@id="content"]/div/div/div/div[4]/div[2]/table/tbody/tr[39]/td[4]/nobr'
        
    elif form_cbr == '803':
        xpth = '//*[@id="content"]/div/div/div/div[4]/div[2]/table/tbody/tr[39]/td[4]/nobr'
        
    
    dt = start_date
    
    for mon in range(mon_len):
        
        own_funds = []
        
        if notify == True:
            print()
            print(f'Собираю данные за {dt}:')
        
        for i in range(banks.shape[0]):
            
            try:
                search_window  =  driver.find_element(By.ID,  'SearchPrase' )
                search_window.send_keys(banks.ogrn[i] + Keys.RETURN)
        
                time.sleep(1)
    
                bank = banks.csname[i]
                bank_name = driver.find_element(By.LINK_TEXT, bank)
        
                driver.find_element(By.ID,  'SearchPrase' ).clear()
    
            except NoSuchElementException:
                
                own_funds.append(None)
                if notify == True:
                    print(f'Банк с названием {bank} не найден! Перехожу к следующему')
                driver.find_element(By.ID,  'SearchPrase' ).clear()
            
            else:
                bank_name.click()
                window_handles = driver.window_handles
                driver.switch_to.window(window_handles[-1])
        
                regnum = driver.find_element(By.XPATH, "//*[@id='content']/div/div/div/div[2]/div[2]/div[9]/div[2]").text
        
                reports = driver.find_element(By.LINK_TEXT, "Раскрываемая отчетность")
                reports.click()
        
                time.sleep(1)
            
                link = "https://cbr.ru/banking_sector/credit/coinfo/f" + form_cbr + "?regnum=" + str(regnum) + '&dt=' + dt
            
                if notify == True:
                    print(f"{i+1})", bank, link)
    
                try:
                    driver.get(link)
        
                except NoSuchElementException:
                    if notify == True:
                        print(f'Возникли проблемы с регистрационным номером банка {bank}, пропускаю')
            
                    own_funds.append(None)
            
                    driver.close()
                    driver.switch_to.window(driver.window_handles[0])
                
                else:
                    window_handles = driver.window_handles
                    driver.switch_to.window(window_handles[-1])
        
                    time.sleep(1)
            
                    try:
                        funds = driver.find_element(By.XPATH, xpth)
                    
                    except:
                        if notify == True:
                            print(f'Возникли проблемы с регистрационным номером банка {bank}, пропускаю')
                        own_funds.append(None)
                        
                    else:
                        time.sleep(2)
                        
                        try:
                            own_funds.append(int(funds.text.replace(' ', '')))
                        
                        except ValueError:
                            own_funds.append(None)
                            if notify == True:
                                print('Не могу преобразовать текст в число, пропускаю')
                    
                    driver.close()
                    driver.switch_to.window(driver.window_handles[0])

        output_df[dt] = own_funds
        
        if os.path.exists(output_file):
            os.remove(output_file)
        
        output_df.to_excel(output_file, index=False)
                    
        if step == 1:
            
            if  int(dt.split('-')[1]) < 12:
                
                if int(dt.split('-')[1]) < 9:
                    dt = dt.split('-')[0] + '-0' + str(int(dt.split('-')[1]) + 1) + '-' + dt.split('-')[2]
                    
                else:
                    dt = dt.split('-')[0] + '-'+ str(int(dt.split('-')[1]) + 1) + '-' + dt.split('-')[2]
            
            elif int(dt.split('-')[1]) == 12:
                dt = str(int(dt.split('-')[0])+1) + '-01-' + dt.split('-')[2]
        
        elif step == 3:
            
            if dt[4:] == '03':
                dt = dt[:4] + '06'
            
            elif dt[4:] == '06':
                dt = dt[:4] + '09'
            
            elif dt[4:] == '09':
                 dt = dt[:4] + '12'
            
            elif dt[4:] == '12':
                 dt = dt[:4] + '03'
                    
    driver.quit()
    
    end_time = time.time()
    elapsed_time = end_time - start_time
    
    if notify == True:
        print()
    print('Прошло времени:', round(elapsed_time / 60, 2), 'мин. Данные находятся в файле:', output_file)
