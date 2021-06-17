
import selenium
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import easygui


def add_team_func(ids,pas,name):
    URL='https://login.microsoftonline.com/common/oauth2/authorize?response_type=id_token&client_id=5e3ce6c0-2b1f-4285-8d4b-75ee78787346&redirect_uri=https%3A%2F%2Fteams.microsoft.com%2Fgo&state=a3e39318-0960-4927-8c0f-2a57c52d0394&client-request-id=65d68f2e-5199-412f-970c-a78c93e0001a&x-client-SKU=Js&x-client-Ver=1.0.9&nonce=f1a7c949-f673-4e5d-961b-a7e2d244f523&domain_hint=&sso_reload=true'
    
    driver = webdriver.Chrome(executable_path='./chromedriver')   ## Chomedriver Directory
    driver.maximize_window()
    driver.get(URL)
    
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

    WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.CSS_SELECTOR,'.ts-btn.ts-btn-fluent.ts-btn-fluent-secondary.ts-btn-fluent-with-icon.join-team-button')))
    create_team=driver.find_element_by_css_selector('.ts-btn.ts-btn-fluent.ts-btn-fluent-secondary.ts-btn-fluent-with-icon.join-team-button') # select create or join button
    create_team.click()
    
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
    df=pd.read_excel('mail_id.xlsx')
    df.loc[len(df)]=''
    
    
    WebDriverWait(driver, 50).until(EC.element_to_be_clickable((By.CLASS_NAME,'ts-people-picker')))
    member_name=driver.find_element_by_class_name('ts-people-picker')
    print('Found First')
    driver.execute_script("window.open('');")
    driver.switch_to.window(driver.window_handles[1])
    driver.close()
    driver.switch_to.window(driver.window_handles[0])
    time.sleep(1)
    memberName=member_name.find_element_by_css_selector('.ts-search-input.ng-pristine.ng-valid.ng-empty.ng-touched')
 
    for i in range(len(df['id'])):
        memberName.send_keys(df['id'][i])
        time.sleep(3)
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
    
    return driver
    
ids=easygui.enterbox(msg="Enter your Mail ID:",title="Mail ID",strip=True)
pas=easygui.passwordbox(msg="Enter your Password:",title="Password")
name=easygui.enterbox(msg="Enter the name of the team:",title="Name of the team",strip=True)
if (ids is None or pas is None or name is None) or (ids=='' or pas=='' or name==''):
    pass
else:
    try:
        driver=add_team_func(ids=ids,pas=pas,name=name)
    except Exception as e:
        print(e)
        rerun=easygui.ccbox(msg='Something went wrong!!!\n Do you want to continue re-run of the code')
        if(rerun==True):
            driver=add_team_func(ids=ids,pas=pas,name=name)
        else:
            pass



            
