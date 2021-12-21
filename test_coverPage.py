from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.by import By

from time import sleep
import pandas as pd
import configparser
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains


cisco_mail = 'wlongsid@cisco.com'
cisco_pass = '@Rt025813603'
# cisco_mail = 'npanichc@cisco.com'
# cisco_pass = 'white2%%%%QW'



driver=webdriver.Chrome()

driver.get("https://www-plmprd.cisco.com/Agile/")



WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//input[@id="userInput"]'))).send_keys(cisco_mail)

WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//input[@name="login-button"]'))).click()

WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//input[@id="passwordInput"]'))).send_keys(cisco_pass)

WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//input[@id="login-button"]'))).click()

count_render_2fa = 0

while (driver.title != "Two-Factor Authentication"):
    sleep(1)
    count_render_2fa+=1
    print("Wait for two-factor authentication render:",count_render_2fa)


push_login_btn = WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//button[@type="submit"]')))

push_login_btn_text = push_login_btn.text

print("Found button:",push_login_btn_text)

if push_login_btn_text == "Send Me a Push":

    push_login_btn.click()
    

else:
    print("Not found \"Send Me a Push\" button")

        

two_fa_url=driver.current_url

count_duo_pass = 0

while(two_fa_url==driver.current_url):
    sleep(1)
    count_duo_pass+=1
    print("Wait for count_duo_pass:",count_duo_pass)
    if count_duo_pass == 30:
        print("!!! LOGIN TIMEOUT !!!")
        break






print("Login to QIS is success!!!")
sleep(1)
# driver.get("https://www-plmprd.cisco.com/Agile/")

driver.get("https://www-plmprd.cisco.com/Agile/")

WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//input[@name="QUICKSEARCH_STRING"]'))).send_keys("FA-0460606")
WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//a[@id="top_simpleSearch"]'))).click()
fa_case_status = WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//h2[@style="color:Blue;"]'))).text
print(fa_case_status)


html = driver.page_source

with open('html.txt', 'w', encoding='utf-8-sig') as file:
    file.write((html))



WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//span[@id="MSG_Editspan"]'))).click()

WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//span[@id="show_floater_R1_4859_0_displayspan"]'))).click()

html = driver.page_source

with open('html2.txt', 'w', encoding='utf-8-sig') as file:
    file.write((html))



WebDriverWait(driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//input[@id="floater_search_text_R1_4859_0_display"]'))).send_keys("wlongsid")

WebDriverWait(driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//span[@id="searchspan"]'))).click()

print(WebDriverWait(driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//td[@class=" yui-dt-string yui-dt-first"]'))).text)



element = WebDriverWait(driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//td[@class=" yui-dt-string yui-dt-first"]')))

# create action chain object
action = ActionChains(driver)

# double click the item
action.double_click(on_element = element)

# perform the operation
action.perform()

sleep(5)

html = driver.page_source
with open('html3.txt', 'w', encoding='utf-8-sig') as file:
    file.write((html))

sleep(5)

# WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//span[@id="MSG_Cancelspan"]'))).click()

box = WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//textarea[@name="R1_2017_7"]'))).text

print(box.__len__())