from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.by import By

from time import sleep
import pandas as pd
import configparser
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains


import pywinauto


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

WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//input[@name="QUICKSEARCH_STRING"]'))).send_keys("FA-0505813")
WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//a[@id="top_simpleSearch"]'))).click()
fa_case_status = WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//h2[@style="color:Blue;"]'))).text
print(fa_case_status)

WebDriverWait(driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//div[@id="tabsDiv"]//li')))[-2].click()






countIcon = WebDriverWait(driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//table[@class="GMSection"]//tr[@class="GMDataRow"]')))

print(countIcon)
print("count tr",countIcon.__len__())




for i in range(3):

    WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//a[@id="MSG_AddAttachment_10"]'))).click()



    WebDriverWait(driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//span[@class="fileinput-span lf-btn-span"]'))).click()

    html = driver.page_source
    with open('html2.txt', 'w', encoding='utf-8-sig') as file:
        file.write((html))


    # handles = driver.window_handles
    # print("handle_windows",handles.__len__())

    # for handle in handles:

    #     driver.switch_to.window(handle)
    #     window_title = driver.title
    #     print(window_title)



    print(pywinauto.findwindows.find_windows())
    notFoundBrowseGUI = True
    while(notFoundBrowseGUI):
        try:
            app = pywinauto.Application().connect(title="Open", class_name="#32770")

            print(app)

            dlg = app.window(title="Open", class_name="#32770")

            print(dlg)

            notFoundBrowseGUI = False



        except:

            print("Wait for Browse GUI")

    # app.dlg.print_control_identifiers()

    resultInspectPath = "C:\\Users\\wisitl\\Downloads\\"

    addedFiles = resultInspectPath
    for i in ["FIW220204S7_22.JPG", "FIW220204S7_11.JPG", "FIW220204S7_33.JPG" ]:
        addedFiles+="\""+i+"\""


    print(addedFiles)
    add = app.dlg.Edit.type_keys(addedFiles+"{ENTER}")

    print(add)




    #tst wait untill uploaded


    html = driver.page_source
    with open('html3.txt', 'w', encoding='utf-8-sig') as file:
        file.write((html))

    WebDriverWait(driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//div[@class="upload-buttons-span"]//span[@id="uploadFilesUMspan"]'))).click()

    # sleep(5)

    # WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//div[@class="upload-buttons-span"]//span[@id="uploadFilesUMspan"]'))).click()


    downloading =True

    while(downloading):

        try:
            download_value = WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//div[@class="progress"]'))).get_attribute("aria-valuenow")
            print(download_value)
            downloading = False

        except:
            print("Not clikc upload")
            WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//div[@class="upload-buttons-span"]//span[@id="uploadFilesUMspan"]'))).click()




