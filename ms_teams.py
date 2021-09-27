#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Jun 17 21:13:18 2021

@author: Arpit
"""
import selenium
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from validate_email import validate_email
import time
import os

def add_team_func(ids,pas,name,fileName,col_name):

    URL='https://teams.microsoft.com'
    options = webdriver.ChromeOptions()
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    driver= webdriver.Chrome(executable_path='chromedriver.exe',options=options)
    driver.maximize_window()
    driver.get(URL)

    try: 
        WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.ID,'i0116')))   
        login_id=driver.find_element_by_id('i0116')
        login_id.send_keys(ids)      #### Mail Id
        submit1=driver.find_element_by_id('idSIButton9')
        submit1.click()
        
        password=driver.find_element_by_id('i0118')
        password.send_keys(pas)                 ##### Password
        time.sleep(4)
        submit2=driver.find_element_by_id('idSIButton9') # Submitting the password
        submit2.click()
        
        submit3=driver.find_element_by_id('idSIButton9') # To make login 
        submit3.click()

        try:
            driver.find_element_by_id('passwordError')#  Might throw error due to incorrect password or mail ID
            msg='Invalid Credentials'
            driver.close()
            return msg
        except Exception as e:
            WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.CSS_SELECTOR,'.ts-btn.ts-btn-fluent.ts-btn-fluent-secondary.ts-btn-fluent-with-icon.join-team-button')))
            create_team=driver.find_element_by_css_selector('.ts-btn.ts-btn-fluent.ts-btn-fluent-secondary.ts-btn-fluent-with-icon.join-team-button') # select create or join button
            create_team.click()
    
    except Exception as e:
        print(e)
        msg='Invalid Credentials'
        driver.close()
        return msg
    
    WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.CSS_SELECTOR,'.ts-btn.ts-btn-fluent.ts-btn-fluent-with-icon.ts-btn-fluent-primary')))
    create=driver.find_element_by_css_selector('.ts-btn.ts-btn-fluent.ts-btn-fluent-with-icon.ts-btn-fluent-primary')
    driver.execute_script("arguments[0].click();",create) # This works for hidden buttons as it holds the object until this statement is executed
    
    WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.ID,'create-teamname')))
    team_name=driver.find_element_by_id('create-teamname')
    team_name.send_keys(name)      ## Name of the Team

    submit_div=driver.find_elements_by_class_name('ts-modal-button-div')
    btn=submit_div[1].find_element_by_css_selector('.ts-btn.ts-btn-fluent.ts-btn-fluent-primary')
    time.sleep(2)
    btn.click()
   
    ######### Adding the members
    import pandas as pd
    pd.options.mode.chained_assignment = None  # default='warn'

    paths='file/'+fileName
    if fileName.split('.')[1]=='xlsx':
        df=pd.read_excel(paths,usecols=[col_name])
    elif fileName.split('.')[1]=='csv':
        df=pd.read_csv(paths,usecols=[col_name])
    
    # Fetching values that are only valid email ID
    new_df=df[df[col_name].apply(lambda x: validate_email(str(x)))]
    new_df.reset_index(inplace=True,drop=True)

    # Fetching values that are not added
    remvoed_values= df[df[col_name].isin(new_df[col_name])==False].dropna()
    new_df.loc[len(new_df)]=''
    
    rv_path= 'file/'+fileName+'_rv.xlsx'
    remvoed_values.to_excel(rv_path,index=False)

    
    WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.CLASS_NAME,'ts-people-picker')))
    member_name=driver.find_element_by_class_name('ts-people-picker')
    driver.execute_script("window.open('');")
    driver.switch_to.window(driver.window_handles[1])
    driver.close()
    driver.switch_to.window(driver.window_handles[0])
    time.sleep(1)
    memberName=member_name.find_element_by_css_selector('.ts-search-input.ng-pristine.ng-valid.ng-empty.ng-touched')
    
    for i in range(len(new_df[col_name])):
        if i==3:
            try:
                length=len(driver.find_element_by_class_name('recipients-list').find_elements_by_class_name('ts-selected-contact'))
                if length<2:
                    msg='Something went wrong with MS Teams. Not Able to add Members'
                    driver.close()
                    return msg
                    
            except:
                msg='Something went wrong with MS Teams. Not Able to add Members'
                driver.close()
                return msg
        memberName.send_keys(new_df[col_name][i])
        time.sleep(5)
        memberName.send_keys(Keys.ARROW_DOWN)
        memberName.send_keys(Keys.ENTER)
        memberName.clear()
    
    add_div=driver.find_element_by_class_name('ts-add-btn-div')
    add_button=add_div.find_element_by_css_selector('.ts-btn.ts-btn-fluent.ts-btn-fluent-primary')
    add_button.click()
    
    while True:
        footer_div=driver.find_element_by_class_name('ts-modal-dialog-footer-buttons')
        foot_btn=footer_div.find_element_by_class_name('ts-modal-button-div').find_element_by_css_selector('.ts-btn.ts-btn-fluent.ts-btn-fluent-secondary')
        if foot_btn.text=='Close':
            break
        time.sleep(2)
    
    time.sleep(4)
    foot_btn.click()
    
    return ''
    


            
