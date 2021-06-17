import selenium
from selenium import webdriver
from selenium.webdriver.support.select import Select 
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from bs4 import BeautifulSoup
import requests
import time
import easygui


def add_team_func(id,pas):
    URL='https://login.microsoftonline.com/common/oauth2/authorize?response_type=id_token&client_id=5e3ce6c0-2b1f-4285-8d4b-75ee78787346&redirect_uri=https%3A%2F%2Fteams.microsoft.com%2Fgo&state=a3e39318-0960-4927-8c0f-2a57c52d0394&client-request-id=65d68f2e-5199-412f-970c-a78c93e0001a&x-client-SKU=Js&x-client-Ver=1.0.9&nonce=f1a7c949-f673-4e5d-961b-a7e2d244f523&domain_hint=&sso_reload=true'
    page= requests.get(URL)
    bs=BeautifulSoup(page.content,'html.parser')
    pretty_page=bs.prettify()
    
    driver = webdriver.Chrome(executable_path='./chromedriver')   ## Chomedriver Directory
    driver.maximize_window()
    driver.get(URL)
    action = ActionChains(driver)
    
    login_id=driver.find_element_by_id('i0116')
    login_id.send_keys(id)      #### Mail Id
    submit1=driver.find_element_by_id('idSIButton9')
    submit1.click()
    
    
    
    password=driver.find_element_by_id('i0118')
    password.send_keys(pas)                 ##### Password
    time.sleep(4)
    submit2=driver.find_element_by_id('idSIButton9') # Submitting the password
    submit2.click()
    
    submit3=driver.find_element_by_id('idSIButton9') # To make login 
    submit3.click()

    WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.CSS_SELECTOR,'.ts-btn.ts-btn-fluent.ts-btn-fluent-secondary.ts-btn-fluent-with-icon.join-team-button')))
    create_team=driver.find_element_by_css_selector('.ts-btn.ts-btn-fluent.ts-btn-fluent-secondary.ts-btn-fluent-with-icon.join-team-button') # select create or join button
    create_team.click()
    
    WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.CSS_SELECTOR,'.ts-btn.ts-btn-fluent.ts-btn-fluent-with-icon.ts-btn-fluent-primary')))
    create=driver.find_element_by_css_selector('.ts-btn.ts-btn-fluent.ts-btn-fluent-with-icon.ts-btn-fluent-primary')
    driver.execute_script("arguments[0].click();",create) # This works for hidden buttons as it holds the object until this statement is executed
    
    WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.ID,'create-teamname')))
    team_name=driver.find_element_by_id('create-teamname')
    team_name.send_keys('my team')      ## Name of the Team
    
    submit_div=driver.find_elements_by_class_name('ts-modal-button-div')
    btn=submit_div[1].find_element_by_css_selector('.ts-btn.ts-btn-fluent.ts-btn-fluent-primary')
    time.sleep(2)
    btn.click()
    
    cont=easygui.boolbox(msg="Are you sure you want to continue",title="Add Members",choices=('[Y]es','[N]o'))
    
    if cont:
        ######### Adding the members
        import pandas as pd
        df=pd.read_excel('mail_id.xlsx')
        df.loc[len(df)]=''
        
        time.sleep(4)
        driver.switch_to.window(driver.current_window_handle) 
        WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.CSS_SELECTOR,'.ts-search-input.ng-pristine.ng-valid.ng-empty.ng-touched')))
        member_name=driver.find_element_by_css_selector('.ts-search-input.ng-pristine.ng-valid.ng-empty.ng-touched')
        for i in range(len(df['id'])):
            member_name.send_keys(df['id'][i])
            time.sleep(3)
            member_name.send_keys(Keys.ARROW_DOWN)
            member_name.send_keys(Keys.ENTER)
            member_name.clear()
        
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
    
    return driver
    
        
ids=easygui.enterbox(msg="Enter your Mail ID:",title="Mail ID",strip=True)
pas=easygui.passwordbox(msg="Enter your Password:",title="Password")
if ids==None or pas==None or ids=='' or pas=='':
    pass
else:
    try:
        driver=add_team_func(id=ids,pas=pas)
    except:
        rerun=easygui.ccbox(msg='Something wetn wrong!!!\n Do you want to continue re-run of the code')
        if(rerun==True):
            driver=add_team_func(id=ids,pas=pas)
        else:
            pass

# WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.CLASS_NAME,'ts-modal-button-div')))



            
