# -*- coding: utf-8 -*-
"""
Created on Mon Feb 21 12:04:37 2022

@author: Paul Ferdinand Popp, Assistant, CAD
"""

import getRegisterScreenshot_Helpers as helpers
import logging
import os
import pyautogui
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from pathlib import Path
import pandas as pd
import time
from datetime import datetime


isProductRun = False
if isProductRun == True:
    tTEST = ""
    Testumgebung = ""
else:
    tTEST = "_TEST"
    Testumgebung = " Testumgebung"


# For given Company name, location, orderNo and idxAddress from InputSheet.xlsx, searches the Handelsregister website for a matching entry. 
# If entry found, save a screenshot of the site and create db entry
def findHandelsregister(name, location, country, orderNo, idxAddress, timeStamp):
    try:
        name = str(name)
        name = name.replace(str(", " + location), "")
        nameSearch = name.replace("-", " ")
        if len(nameSearch.split()) > 5: nameSearch = " ".join(nameSearch.split(" ", 5)[:5])
        foundHandelsregister = False
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        driver = webdriver.Chrome(executable_path = os.getcwd() + '//' + 'chromedriver.exe', options=options)
        driver.get("https://www.handelsregister.de/rp_web/erweitertesuche.xhtml")
        assert "Registerportal" in driver.title
        
        
        elem = driver.find_element_by_id("form:schlagwoerter")
        elem.clear()
        elem.send_keys(nameSearch)
        elem.send_keys(Keys.PAGE_DOWN)
        
        time.sleep(0.1)
        driver.find_element(by=By.XPATH,value='//*[@id="form:schlagwortOptionen"]/div[1]/div/div/div[2]').click()
        time.sleep(0.5)
        
        driver.find_element(by=By.XPATH,value='//*[@id="form:ergebnisseProSeite_label"]').click()
        time.sleep(0.5)
        
        driver.find_element_by_id("form:ergebnisseProSeite_3").click()
        time.sleep(2)

        driver.find_element(by=By.XPATH,value='//*[@id="form:btnSuche"]/span').click()

        
        
        # wait until next page is loaded
        try:
            element_present = EC.presence_of_element_located((By.XPATH, "//*[@class='ui-paginator-current']"))
            WebDriverWait(driver, 10).until(element_present)
        except TimeoutException:
            print ("Timed out waiting for page to load")
            
        
        assert "No results found." not in driver.page_source
        
        
        # count search results
        countResultsText = driver.find_element(by=By.CLASS_NAME,value='ui-paginator-current').text
        countResults = int(countResultsText[int(countResultsText.find('of') + 3):].replace(' records', '')) # Handle 0 results
        if countResults == 0:
            driver.close()
            # comment for test
            return foundHandelsregister
        time.sleep(2)
        
        dataResults = []
        
        for i in range(countResults):
            table_id = driver.find_element_by_xpath("//*[contains(@id, 'ergebnissForm:selectedSuchErgebnisFormTable:" + str(i) + "')]")
            # rows = table_id.find_elements_by_tag_name("tr")
            rows = table_id.find_elements(by=By.TAG_NAME, value='tr')
            j = 0 
            for row in rows:
                j = j + 1
                # result = row.find_elements_by_tag_name("td")#note: index start from 0, 1 is col 2
                result = row.find_elements(by=By.TAG_NAME, value='td')  # note: index start from 0, 1 is col 2
                if j == 2:
                    listResult = [result[0].text, result[1].text, result[2].text, result[3].text]
                    dataResults.append(listResult)
        dfResults = pd.DataFrame(dataResults)
           
        
        # Find best match
        idxBestResult = None
        for i in range(len(dfResults)):
            if (name.upper() in dfResults.iloc[i, 0].upper() and dfResults.iloc[i, 1].upper() == location.upper()):
                logging.info('    Match (0): ' + str(dfResults.iloc[i, 0]) + '-' + str(dfResults.iloc[i, 1]))
                foundHandelsregister = True
                idxBestResult = i
                break
            elif (dfResults.iloc[i, 0].upper() == name.upper()):
                logging.info('   Match (1): ' + str(dfResults.iloc[i, 0]))
                foundHandelsregister = True
                idxBestResult = i
        
        # get Address and Screenshot
        if idxBestResult is not None:            
            elementsScrollTo = driver.find_elements_by_xpath("//*[starts-with(@id, 'ergebnissForm:selectedSuchErgebnisFormTable:" + str(min(idxBestResult + 1, len(dfResults) - 1)) + ":j_idt'" + " )]")
            elementsUT = driver.find_elements_by_xpath("//*[starts-with(@id, 'ergebnissForm:selectedSuchErgebnisFormTable:" + str(idxBestResult) + ":j_idt'" + " )]")
                    
            for i in range(len(elementsUT)):
                
                if '4:popupLink' in elementsUT[i].get_attribute("id"):
                    # RB: solve the problem that the page cannot scroll to the element
                    # actions = ActionChains(driver)
                    # actions.move_to_element(elementsScrollTo[i]).perform()
                    
                    driver.execute_script("arguments[0].scrollIntoView();", elementsScrollTo[i])
                    elementsUT[i].click()
                    break
            body = driver.find_element_by_css_selector('body')
            body.send_keys(Keys.PAGE_DOWN)
            
            # table_id = driver.find_element_by_id('ut_formInfobox:j_idt121')
            # rows = table_id.find_elements_by_tag_name("tr")
            
            ############################################### RB START: find_elemennt_by_id -> alternative methods ###############################################
            
            # original method
            
            # table_id = driver.find_element_by_id('ut_formInfobox:j_idt122') #KA Quick Fix
            # rows = table_id.find_elements(by=By.TAG_NAME, value='tr')            
            
            # dataAddress = []
            # for row in rows:
            #     # result = row.find_elements_by_tag_name("td") #note: index start from 0, 1 is col 2
            #     result = row.find_elements(by=By.TAG_NAME, value='td')  # note: index start from 0, 1 is col 2
            #     listAddress = [result[0].text, result[1].text]
            #     dataAddress.append(listAddress)
            # dfAddress = pd.DataFrame(dataAddress)
            # compName = dfAddress.iloc[-3, 1]
            # street = dfAddress.iloc[-2, 1]
            # try:
            #     zipCode = dfAddress.iloc[-1, 1].split(' ', 1)[0]
            #     location = dfAddress.iloc[-1, 1].split(' ', 1)[1]
            # except Exception:
            #     logging.error("zipcode/location not found OrderNo - " + orderNo + " - " + name + " - " + location)
            #     zipCode = 'N/A'
            #     location = dfAddress.iloc[-1, 1]            
            

            #KA & RB changed method from static to dynamic tags
            # Alt method 2
            try:
                tb_text = ''
                elements = driver.find_elements_by_tag_name('tbody')
                for element in elements:
                    #print(element.text) 
                    tb_text += element.text 
                tb_text_cut = tb_text.split("Address (subject to correction):",1)[1]
                list_tb_text_cut = list(filter(bool, tb_text_cut.splitlines())) # convert text to list
                list_tb_text_cut = [j.strip() for j in list_tb_text_cut] # eliminate the white spaces
                compName = list_tb_text_cut[0]
                street = list_tb_text_cut[1]
                try:
                    zipCode = list_tb_text_cut[2].split(' ', 1)[0]
                    location = list_tb_text_cut[2].split(' ', 1)[1]
                except Exception:
                    logging.error("zipcode/location not found OrderNo - " + orderNo + " - " + name + " - " + location)
                    zipCode = 'N/A'
                    location = list_tb_text_cut[2]   
                
            # Alt method 3
            except:
                print ("error in method 2")
                rows = driver.find_elements_by_tag_name('tr')
                dataAddress = []
                listAddress = []
                for row in rows[1:]:
                    result = row.find_elements_by_tag_name("td")  # note: index start from 0, 1 is col 2
                    listAddress = [result[0].text, result[1].text]
                    dataAddress.append(listAddress)
                dfAddress = pd.DataFrame(dataAddress)
                compName = dfAddress.iloc[-3, 1]
                street = dfAddress.iloc[-2, 1]
                try:
                    zipCode = dfAddress.iloc[-1, 1].split(' ', 1)[0]
                    location = dfAddress.iloc[-1, 1].split(' ', 1)[1]
                except Exception:
                    logging.error("zipcode/location not found OrderNo - " + orderNo + " - " + name + " - " + location)
                    zipCode = 'N/A'
                    location = dfAddress.iloc[-1, 1]  
                

            ############################################### RB END: find_elemennt_by_id -> alternative methods ###############################################                
            

            saveDir = helpers.getSaveDir(name)
            fileName = helpers.getFilename(name) + '_' + timeStamp + '.png'
            savePath = saveDir + fileName 
            helpers.createSQLAddressEntry(compName, street, zipCode, location, country, orderNo, idxAddress, saveDir, fileName, timeStamp, 'Handelsregister')
            myScreenshot = pyautogui.screenshot()
            myScreenshot.save(savePath)
        else:
            logging.info(' no match')
        
        driver.close()
        if foundHandelsregister: print('Handelsregister found.')
        return foundHandelsregister
        
    except Exception as e:
        logging.error('Error in findHandelsregister: '+ str(e)+ 'Best match couldnt be found in column-list' )
        driver.close()
        return foundHandelsregister

def findDBHoovers(name,location, country, orderNo, idxAddress, timeStamp):
    try:
        
        name = str(name)
        name = name.replace(str(", " + location), "")
        foundDBHoovers = False
        # search for name and location and save results in dfResults
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        driver = webdriver.Chrome(executable_path = os.getcwd() + '//' + 'chromedriver.exe', options=options)
        driver.get("https://app.dnbhoovers.com/")
        # "login"
        username = driver.find_element_by_name("username")
        username.clear()
        username.send_keys('Bob.b.Tix@de.ey.com')
        driver.find_element_by_class_name("continue-btn").click()
        while 'Desktop' not in driver.title:
            time.sleep(1)
        time.sleep(3)
               
       # search                                                             # For pythonenv version 3.7 and higher
        # search = driver.find_element(by=By.CLASS_NAME,value='ant-input')
        # search.clear()
        # search.send_keys(name)
        # search.send_keys(Keys.RETURN)
        # while 'Search & Build a List' not in driver.title:
        #     time.sleep(1)
        # time.sleep(10)
        # if "We're sorry, we couldn't find any results" in driver.page_source:
        #     logging.info(' no match')
        #     driver.close()
        #     return False
        # driver.find_element(by=By.CLASS_NAME,value='selected-view-icon').click()
        # time.sleep(1)
        # viewSelection = driver.find_elements(by=By.CLASS_NAME,value='view-icon')
        # time.sleep(1)
        # viewSelection[1].click()
        # time.sleep(1)
        # driver.find_element(by=By.CLASS_NAME,value='table-label').click()
        # while 'Tradestyle' not in driver.page_source:
        #     time.sleep(1)
        # time.sleep(1)
        
        # rows = driver.find_element(by=By.TAG_NAME,value="tbody").find_elements(by=By.TAG_NAME,value="tr")
        # dataResults = []
        # for row in rows:   
        #     result = row.find_element(by=By.TAG_NAME,value="td") #note: index start from 0, 1 is col 2
        #     listResult = [result[1].text, result[4].text, result[7].text, result[9].text, result[10].text]
        #     dataResults.append(listResult)
        # dfResults = pd.DataFrame(dataResults).iloc[1: , :].reset_index()
        # dfResults.columns = ['Index', 'Company', 'Street', 'Location', 'ZipCode', 'Country']
        
        # search
        search = driver.find_element_by_class_name("ant-input")     # For python embedded version 3.6 and below
        search.clear()
        search.send_keys(name)
        search.send_keys(Keys.RETURN)
        while 'Search & Build a List' not in driver.title:
            time.sleep(1)
        time.sleep(10)
        if "We're sorry, we couldn't find any results" in driver.page_source:
            logging.info(' no match')
            driver.close()
            return False
        driver.find_element_by_class_name('selected-view-icon').click()
        time.sleep(1)
        viewSelection = driver.find_elements_by_class_name('view-icon')
        time.sleep(1)
        viewSelection[1].click()
        time.sleep(1)
        driver.find_element_by_class_name('table-label').click()
        while 'Tradestyle' not in driver.page_source:
            time.sleep(1)
        time.sleep(1)
        
         ######################### old method ###########################################################
        # table_id = driver.find_element(by=By.CLASS_NAME,value='ant-table-tbody')
        # rows = table_id.find_elements_by_tag_name("tr") # get all of the rows in the table
        # rows = driver.find_element(by=By.TAG_NAME,value='tbody').find_elements(by=By.TAG_NAME,value='tr') # get all of the rows in the table
        # dataResults = []
        # for row in rows:   
        #     result = row.find_element(by=By.TAG_NAME,value='td') #note: index start from 0, 1 is col 2
        #     listResult = [result[1].text, result[4].text, result[7].text, result[9].text, result[10].text]
        #     dataResults.append(listResult)
        # dfResults = pd.DataFrame(dataResults).iloc[1: , :].reset_index()
        # dfResults.columns = ['Index', 'Company', 'Street', 'Location', 'ZipCode', 'Country']
        
        ######################### New method ###########################################################
        
        # RB: change find method
        # table_id = driver.find_element_by_class_name('ant-table-tbody')
        # rows = table_id.find_elements_by_tag_name("tr") # get all of the rows in the table
        rows = driver.find_element_by_tag_name("tbody").find_elements_by_tag_name("tr") # get all of the rows in the table
        dataResults = []
        for row in rows:   
            result = row.find_elements_by_tag_name("td") #note: index start from 0, 1 is col 2
            listResult = [result[1].text, result[4].text, result[7].text, result[9].text, result[10].text]
            dataResults.append(listResult)
        dfResults = pd.DataFrame(dataResults).iloc[1: , :].reset_index()
        dfResults.columns = ['Index', 'Company', 'Street', 'Location', 'ZipCode', 'Country']
       
        # Find best result
        try: 
            dfMatches = dfResults.loc[(dfResults['Company'].apply(lambda x: name.upper() in x.upper())) & \
                                  (dfResults['Location'].apply(lambda x: location.upper() in x.upper()))]
            if len(dfMatches) == 0:
                dfMatches = dfResults.loc[(dfResults['Company'].apply(lambda x: name.upper() in x.upper()))]
            lendfMatches = len(dfMatches)
        except Exception:
            lendfMatches = 0
            
        if lendfMatches > 0:
            foundDBHoovers = True
            dfMatches = dfMatches.iloc[0, :]
            logging.info('    Match (2): ' + str(dfMatches['Company']) + '-' + str(dfMatches['Location']))
            saveDir = helpers.getSaveDir(name)
            fileName = helpers.getFilename(name) + '_' + timeStamp + '.png'
            savePath = saveDir + fileName 
            #linkMatch = driver.find_element_by_xpath("//*[@id='main']/section/main/section/main/main/section/main/div/div[3]/div[2]/div/div/div/div/div/div[2]/table/tbody/tr[" +\
            #                                         str(dfMatches['Index'] + 1) +"]/td[2]/div/span/a").get_attribute('href')
            
            driver.find_element_by_class_name('selected-view-icon').click()
            time.sleep(2)
            driver.find_element_by_xpath("//*[@id='main']/section/main/section/main/main/section/main/div/div[3]/div[1]/div/div[2]/div[2]/div/div[2]/div[3]/div/ul/li[1]/div").click()
            viewSelection = driver.find_elements_by_class_name('view-icon')
            time.sleep(1)
            viewSelection[0].click()
            time.sleep(7)
           
            if len(dfResults) == 1:
                linkMatch = driver.find_element_by_xpath("//*[@id='main']/section/main/section/main/main/section/main/div/div[3]/div[2]/div/ul/div/li/div[1]/div[3]/div/div[2]/a").get_attribute('href')
            else:
                linkMatch = driver.find_element_by_xpath("//*[@id='main']/section/main/section/main/main/section/main/div/div[3]/div[2]/div/ul/div[" +\
                                                         str(dfMatches['Index']) +"]/li/div[1]/div[3]/div/div[2]/a").get_attribute('href')
            
            time.sleep(1)
            driver.get(linkMatch)
            time.sleep(2)
            helpers.createSQLAddressEntry(dfMatches['Company'], dfMatches['Street'], dfMatches['ZipCode'], dfMatches['Location'], country, orderNo, idxAddress, saveDir, fileName, timeStamp, 'D&B Hoovers')
            myScreenshot = pyautogui.screenshot()
            myScreenshot.save(savePath)
        else:
            logging.info(' no match')
            
        driver.close()
        if foundDBHoovers: print('DB Hoovers found.')
        return foundDBHoovers
    
    except Exception as e:
        logging.error('Error in findDBHoovers: '+ str(e))
        driver.close()
        driver.quit()        
        return foundDBHoovers
    
                

        




