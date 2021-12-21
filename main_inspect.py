from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.by import By
from time import sleep
import pandas as pd
import configparser
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import os
import re
import sys
import pywinauto
import datetime

class INSPECTION_MACRO:

    
    def __init__(self):

        #config.init file
        self.my_config_parser = configparser.ConfigParser()
        self.my_config_parser.read('config.ini')
        self.config = { 

        'inspect_email': self.my_config_parser.get('INSPECTION_LOGIN','email'),
        'inspect_password': self.my_config_parser.get('INSPECTION_LOGIN','password'),
        'inspect_path': self.my_config_parser.get('INSPECTION_LOGIN','inspect_path'),
        'inherit_path': self.my_config_parser.get('INSPECTION_LOGIN','inherit_path'),

        }

        #config pd with excel
        self.all_fa_sn_qms_sheet = {}
        self.all_fa_case_result_sheet = {}
        
        # self.raw_sn = ""
        
        #config read all files name 
        self.all_fa_case_inspect_files = {}

        #debug
        self.can_not_update_state = {}

        return None



    def read_excel_data(self):

        self.df = pd.read_excel('Inspection_Template.xlsx',["qms","result"])
        # self.index = self.df.index
        # self.number_of_rows = self.index.__len__()

        self.df_qms_sheet = self.df["qms"]

        self.df_result_sheet = self.df["result"]


        for i in range(self.df_qms_sheet.__len__()):

            fa_case = self.df_qms_sheet["Case Number"][i]

            if fa_case not in self.all_fa_sn_qms_sheet:

                self.all_fa_sn_qms_sheet[fa_case] = [self.df_qms_sheet["Serial Number"][i]]

            else:

                self.all_fa_sn_qms_sheet[fa_case].append(self.df_qms_sheet["Serial Number"][i])

                        
        for i in range(self.df_result_sheet.__len__()):

            self.all_fa_case_result_sheet[self.df["result"]["Case Number"][i]] = {  "Case Owner" : self.df["result"]["Case Owner"][i],
                                                                                    "PF" : self.df["result"]["PF"][i],
                                                                                    "Part Inspection Summary" : self.df["result"]["Part Inspection Summary"][i],
                                                                                    "Is Potential Counterfeit" : self.df["result"]["Is Potential Counterfeit"][i],
                                                                                    "Visual Inspection Code" : self.df["result"]["Visual Inspection Code"][i],
                                                                                    "Case Status" : self.df["result"]["Case Status"][i],
                                                                                    "Fault Duplication Test Plan" : self.df["result"]["Fault Duplication Test Plan"][i] }

            # self.raw_sn+=self.df["Part Inspection Summary"][i]

        # self.all_sn = re.findall("[A-Za-z]{3}[0-9]{4}[0-9A-Za-z]{4}", self.raw_sn)

        return True


    def scan_files(self):

        self.dir_list = os.listdir(self.config["inspect_path"])

        return self.dir_list


    def map_inspect_files(self):

        for fa_case in self.all_fa_sn_qms_sheet:

            self.all_fa_case_inspect_files[fa_case] = []

            for sn in self.all_fa_sn_qms_sheet[fa_case]:

                for file in self.dir_list:

                    if sn in file:

                        self.all_fa_case_inspect_files[fa_case].append(file)

        return self.all_fa_case_inspect_files




    #Control web

    def login(self):

        self.driver=webdriver.Chrome()

        self.driver.get("https://www-plmprd.cisco.com/Agile/")


        WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//input[@id="userInput"]'))).send_keys(self.config["inspect_email"])

        WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//input[@name="login-button"]'))).click()

        WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//input[@id="passwordInput"]'))).send_keys(self.config["inspect_password"])

        WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//input[@id="login-button"]'))).click()

        count_render_2fa = 0

        while (self.driver.title != "Two-Factor Authentication"):
            sleep(1)
            count_render_2fa+=1
            print("Wait for two-factor authentication render:",count_render_2fa)

        print("Fix to TOKEN")

        push_login_btn = WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//button[@type="submit"]')))

        push_login_btn_text = push_login_btn.text

        print("Found button:",push_login_btn_text)

        # if push_login_btn_text == "Send Me a Push":

        #     push_login_btn.click()
            

        # else:
        #     print("Not found \"Send Me a Push\" button")

                

        two_fa_url=self.driver.current_url

        count_duo_pass = 0

        while(two_fa_url==self.driver.current_url):
            sleep(1)
            count_duo_pass+=1
            print("Wait for count_duo_pass:",count_duo_pass)
            if count_duo_pass == 30:
                print("!!! LOGIN TIMEOUT !!!")
                self.driver.quit()
                sys.exit()
                


        print("Login to QIS is success!!!")
        sleep(1)
        # self.driver.get("https://www-plmprd.cisco.com/Agile/")

        self.driver.get("https://www-plmprd.cisco.com/Agile/")

        self.main_page = self.driver.current_window_handle

        print("Main Page:",self.main_page)


        handles = self.driver.window_handles

        for handle in handles:
            sleep(1)
            self.driver.switch_to.window(handle)
            if self.main_page != self.driver.current_window_handle:
                self.driver.close()

        self.driver.switch_to.window(self.main_page)

        return True


    def window_handles(self,fa_case):
        sleep(1)

        not_reset_audit = True
        while(not_reset_audit):

            try :   
                WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//span[@id="MSG_NextStatusspan"]'))).click()
                sleep(1)
                WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//a[@id="ewfinish"]'))).click()
                sleep(1)
                not_reset_audit = False
                
            except:

                reset_handle = False

                handles = self.driver.window_handles
                size = len(handles)
                print("No. Of window handles:",size,handles)

                for handle in handles:
                    sleep(1)
                    self.driver.switch_to.window(handle)
                    window_title = self.driver.title
                    if window_title == 'Application Error':
                        print("Application Error Window:",window_title,handle)
                        self.driver.close()
                        reset_handle = True

                if reset_handle:
                    
                    self.driver.switch_to.window(self.main_page)
                    sleep(1)
                    WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//a[@id="top_simpleSearch"]'))).click()
                    sleep(1)
                    


        #window handle

        found_change_status_window = False
        count_open_change_status_window = 0
        while(not found_change_status_window):

            reset_handle = False
            sleep(1)  #9-Dec-2021  add delay
            handles = self.driver.window_handles
            size = len(handles)
            print("No. Of window handles:",size,handles)

            for handle in handles:
                sleep(1)
                self.driver.switch_to.window(handle)
                window_title = self.driver.title
                if window_title == 'Change Status':
                    print("Change Status Window:",window_title,handle)
                    found_change_status_window = True
                    break
                elif window_title == 'Application Error':
                    print("Application Error Window:",window_title,handle)
                    sleep(1) #9-Dec-2021  add delay
                    self.driver.close()
                    sleep(1)  #9-Dec-2021  add delay
                    reset_handle = True


            if reset_handle:
                
                self.driver.switch_to.window(self.main_page)
                sleep(1)
                WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//span[@id="MSG_NextStatusspan"]'))).click()  #9-Dec-2021  visible to clickable
                sleep(1)
                WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="ewfinish"]'))).click()    #9-Dec-2021
                sleep(1)


            
            count_open_change_status_window+=3
            if count_open_change_status_window > 10:
                print("Can't open change_status Window:",count_open_change_status_window,"second")
                self.can_not_update_state[fa_case] = "Can't Open Change Status Window"
                break


        if count_open_change_status_window < 10:

            WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//a[@id="save"]'))).click()                

            # WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//span[@id="savespan"]'))).click()

            count_close_change_status_window = 0
            while (len(self.driver.window_handles) > 1):
                sleep(1)
                count_close_change_status_window+=1
                print("Wait for Close Change Status Window:",count_close_change_status_window,"second")
                if count_close_change_status_window > 10:
                    print("Can't Close Change Status Window:",count_close_change_status_window)
                    self.can_not_update_state[fa_case] = "Can't Close Change Status Window"
                    sleep(1)  #9-Dec-2021  add delay
                    self.driver.close()
                    sleep(1)  #9-Dec-2021  add delay
                    break       
        sleep(1)
        
        self.driver.switch_to.window(self.main_page)        


    def auto_cover_page_awaiting_assignment(self,fa_case,status):

        try:

            if status == 'testable':

                

                WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//span[@id="MSG_Editspan"]'))).click()

                WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//span[@id="show_floater_R1_4859_0_displayspan"]'))).click()

                WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//input[@id="floater_search_text_R1_4859_0_display"]'))).send_keys(self.all_fa_case_result_sheet[fa_case]['Case Owner'])

                WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//span[@id="searchspan"]'))).click()

                print(WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//td[@class=" yui-dt-string yui-dt-first"]'))).text)

                elm_select = WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//td[@class=" yui-dt-string yui-dt-first"]')))

                ActionChains(self.driver).double_click(on_element = elm_select).perform()

                # # create action chain object
                # action = ActionChains(self.driver)

                # # double click the item
                # action.double_click(on_element = elm_select)

                # # perform the operation
                # action.perform()

                sleep(1)

                #Part Inspection Summary
                WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//textarea[@name="R1_2017_7"]'))).send_keys(self.all_fa_case_result_sheet[fa_case]['Part Inspection Summary'])
            
                sleep(1)
                # Case Review Summary
                WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//textarea[@name="R1_2018_7"]'))).send_keys("-")
                sleep(1)
                # Part History Investigation Summary
                WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//textarea[@name="R1_2019_7"]'))).send_keys("-")
                sleep(1)
                # Is Potential Counterfeit   0 = Null, 1 = No, 2 = Yes            
                WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//select[@name="R1_2020_7"]//option')))[1].click()
                sleep(1)
                # Passed Visual Inspection
                WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//select[@name="R1_2021_7"]//option')))[2].click()

                sleep(1)

                WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="close_floater_R1_4859_0_display"]'))).click() 
                sleep(1)

                WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="MSG_Save"]'))).click()   


                self.window_handles(fa_case)


                WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//span[@id="MSG_Editspan"]'))).click()


                #Fault Duplication Test Plan
                WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//textarea[@name="R1_1332_7"]'))).send_keys(self.all_fa_case_result_sheet[fa_case]['Fault Duplication Test Plan'])
                sleep(1)



                WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="MSG_Save"]'))).click()   

                # sleep(3)

                # WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//span[@id="MSG_Cancelspan"]'))).click()      
                

            elif status == 'untestable':

                WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//span[@id="MSG_Editspan"]'))).click()

                WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//span[@id="show_floater_R1_4859_0_displayspan"]'))).click()

                WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//input[@id="floater_search_text_R1_4859_0_display"]'))).send_keys(self.all_fa_case_result_sheet[fa_case]['Case Owner'])

                WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//span[@id="searchspan"]'))).click()

                print(WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//td[@class=" yui-dt-string yui-dt-first"]'))).text)

                elm_select = WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//td[@class=" yui-dt-string yui-dt-first"]')))

                ActionChains(self.driver).double_click(on_element = elm_select).perform()

                # # create action chain object
                # action = ActionChains(self.driver)

                # # double click the item
                # action.double_click(on_element = elm_select)

                # # perform the operation
                # action.perform()

                sleep(1)
                #Part Inspection Summary
                WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//textarea[@name="R1_2017_7"]'))).send_keys(self.all_fa_case_result_sheet[fa_case]['Part Inspection Summary'])
                sleep(1)
                # Case Review Summary
                WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//textarea[@name="R1_2018_7"]'))).send_keys("-")
                sleep(1)
                # Part History Investigation Summary
                WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//textarea[@name="R1_2019_7"]'))).send_keys("-")
                sleep(1)

                # Is Potential Counterfeit   0 = Null, 1 = No, 2 = Yes       
                if self.all_fa_case_result_sheet[fa_case]['Is Potential Counterfeit'].lower() == "yes":

                    WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//select[@name="R1_2020_7"]//option')))[2].click()

                elif self.all_fa_case_result_sheet[fa_case]['Is Potential Counterfeit'].lower() == "no":

                    WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//select[@name="R1_2020_7"]//option')))[1].click()

                else:
                    WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//select[@name="R1_2020_7"]//option')))[0].click()

                sleep(1)

                # Passed Visual Inspection
                if "pass" in self.all_fa_case_result_sheet[fa_case]['Visual Inspection Code'].lower():

                    WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//select[@name="R1_2021_7"]//option')))[2].click()

                elif "fail" in self.all_fa_case_result_sheet[fa_case]['Visual Inspection Code'].lower():

                    WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//select[@name="R1_2021_7"]//option')))[1].click()

                else:
                    WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//select[@name="R1_2021_7"]//option')))[0].click()


                sleep(1)

                WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="close_floater_R1_4859_0_display"]'))).click() 
                sleep(1)

                WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="MSG_Save"]'))).click()      

                # sleep(3)

                # WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//span[@id="MSG_Cancelspan"]'))).click()   

        except:

            self.can_not_update_state[fa_case] = "Can't update in \"Awaiting Assignment State\""


    def auto_cover_page_inspection_review(self,fa_case,status):

        try :

            if status == 'testable':

                                
                WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//span[@id="MSG_Editspan"]'))).click()

                WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//span[@id="show_floater_R1_4859_0_displayspan"]'))).click()

                WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//input[@id="floater_search_text_R1_4859_0_display"]'))).send_keys(self.all_fa_case_result_sheet[fa_case]['Case Owner'])

                WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//span[@id="searchspan"]'))).click()

                print(WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//td[@class=" yui-dt-string yui-dt-first"]'))).text)

                elm_select = WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//td[@class=" yui-dt-string yui-dt-first"]')))

                ActionChains(self.driver).double_click(on_element = elm_select).perform()

                # # create action chain object
                # action = ActionChains(self.driver)

                # # double click the item
                # action.double_click(on_element = elm_select)

                # # perform the operation
                # action.perform()

                sleep(1)

                #Part Inspection Summary
                old_summary = WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//textarea[@name="R1_2017_7"]'))).text
                

                if old_summary.__len__() > 0:

                    WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//textarea[@name="R1_2017_7"]'))).send_keys("\n\n"+self.all_fa_case_result_sheet[fa_case]['Part Inspection Summary'])

                    sleep(1)       



                else:

                    WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//textarea[@name="R1_2017_7"]'))).send_keys(self.all_fa_case_result_sheet[fa_case]['Part Inspection Summary'])
                
                    sleep(1)
                    # Case Review Summary
                    WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//textarea[@name="R1_2018_7"]'))).send_keys("-")
                    sleep(1)
                    # Part History Investigation Summary
                    WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//textarea[@name="R1_2019_7"]'))).send_keys("-")
                    sleep(1)


                # Is Potential Counterfeit   0 = Null, 1 = No, 2 = Yes            
                WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//select[@name="R1_2020_7"]//option')))[1].click()
                sleep(1)
                # Passed Visual Inspection
                WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//select[@name="R1_2021_7"]//option')))[2].click()

                sleep(1)


                WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="close_floater_R1_4859_0_display"]'))).click() 

                WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="MSG_Save"]'))).click() 

                self.window_handles(fa_case)
                sleep(1)

                WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//span[@id="MSG_Editspan"]'))).click()
                sleep(1)

            
                WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//textarea[@name="R1_1332_7"]'))).send_keys(self.all_fa_case_result_sheet[fa_case]['Fault Duplication Test Plan'])
                
                sleep(1)

                WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="MSG_Save"]'))).click()      

                sleep(1)


                #Fault Duplication Test Plan
                # WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//textarea[@name="R1_1332_7"]'))).send_keys(self.all_fa_case_result_sheet[fa_case]['Fault Duplication Test Plan'])
                # sleep(1)


                # WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="close_floater_R1_4859_0_display"]'))).click() 
                # sleep(1)

                # WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="MSG_Save"]'))).click()      

                # sleep(3)

                # WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//span[@id="MSG_Cancelspan"]'))).click()   

                

            elif status == 'untestable':

                WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//span[@id="MSG_Editspan"]'))).click()

                WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//span[@id="show_floater_R1_4859_0_displayspan"]'))).click()

                WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//input[@id="floater_search_text_R1_4859_0_display"]'))).send_keys(self.all_fa_case_result_sheet[fa_case]['Case Owner'])

                WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//span[@id="searchspan"]'))).click()

                print(WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//td[@class=" yui-dt-string yui-dt-first"]'))).text)

                elm_select = WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//td[@class=" yui-dt-string yui-dt-first"]')))

                ActionChains(self.driver).double_click(on_element = elm_select).perform()

                # # create action chain object
                # action = ActionChains(self.driver)

                # # double click the item
                # action.double_click(on_element = elm_select)

                # # perform the operation
                # action.perform()

                old_summary = WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//textarea[@name="R1_2017_7"]'))).text
                

                if old_summary.__len__() > 0:

                    #Part Inspection Summary
                    WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//textarea[@name="R1_2017_7"]'))).send_keys("\n\n"+self.all_fa_case_result_sheet[fa_case]['Part Inspection Summary'])
                    sleep(1)

                else :

                    #Part Inspection Summary
                    WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//textarea[@name="R1_2017_7"]'))).send_keys(self.all_fa_case_result_sheet[fa_case]['Part Inspection Summary'])
                    sleep(1)



                # Is Potential Counterfeit   0 = Null, 1 = No, 2 = Yes       
                if self.all_fa_case_result_sheet[fa_case]['Is Potential Counterfeit'].lower() == "yes":

                    WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//select[@name="R1_2020_7"]//option')))[2].click()

                elif self.all_fa_case_result_sheet[fa_case]['Is Potential Counterfeit'].lower() == "no":

                    WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//select[@name="R1_2020_7"]//option')))[1].click()

                else:
                    WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//select[@name="R1_2020_7"]//option')))[0].click()

                sleep(1)

                # Passed Visual Inspection
                if "pass" in self.all_fa_case_result_sheet[fa_case]['Visual Inspection Code'].lower():

                    WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//select[@name="R1_2021_7"]//option')))[2].click()

                elif "fail" in self.all_fa_case_result_sheet[fa_case]['Visual Inspection Code'].lower():

                    WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//select[@name="R1_2021_7"]//option')))[1].click()

                else:
                    WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//select[@name="R1_2021_7"]//option')))[0].click()


                sleep(1)
                
                WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="close_floater_R1_4859_0_display"]'))).click() 
                sleep(1)

                WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="MSG_Save"]'))).click() 

                # sleep(3)

                # WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//span[@id="MSG_Cancelspan"]'))).click()   

        except:

            self.can_not_update_state[fa_case] = "Can't update in \"Inspection Review State\""     

                

    def auto_cover_page_fault_duplication(self,fa_case,status):

        try:

            WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//span[@id="MSG_Editspan"]'))).click()

            sleep(1)

            if status == 'testable':
                

                #Part Inspection Summary
                old_summary = WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//textarea[@name="R1_2017_7"]'))).text
                

                if old_summary.__len__() > 0:

                    WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//textarea[@name="R1_2017_7"]'))).send_keys("\n\n"+self.all_fa_case_result_sheet[fa_case]['Part Inspection Summary'])

                    sleep(1)    

                    # Is Potential Counterfeit   0 = Null, 1 = No, 2 = Yes            
                    WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//select[@name="R1_2020_7"]//option')))[1].click()
                    sleep(1)
                    # Passed Visual Inspection
                    WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//select[@name="R1_2021_7"]//option')))[2].click()

                    sleep(1)   
                
                    WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//textarea[@name="R1_1332_7"]'))).send_keys("\n\n"+self.all_fa_case_result_sheet[fa_case]['Fault Duplication Test Plan'])
                    sleep(1)



                    WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="MSG_Save"]'))).click()

                    # sleep(3)

                    # WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//span[@id="MSG_Cancelspan"]'))).click()   




                else:


                    WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//span[@id="MSG_Cancelspan"]'))).click()   

                    self.can_not_update_state[fa_case] = "Fault Duplication State is \"BLANK\'"       

                

            elif status == 'untestable':

                

                old_summary = WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//textarea[@name="R1_2017_7"]'))).text
                

                if old_summary.__len__() > 0:

                    #Part Inspection Summary
                    WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//textarea[@name="R1_2017_7"]'))).send_keys("\n\n"+self.all_fa_case_result_sheet[fa_case]['Part Inspection Summary'])
                    sleep(1)


                    WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="MSG_Save"]'))).click()

                    # sleep(3)

                    # WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//span[@id="MSG_Cancelspan"]'))).click()                       

                else :

                    WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//span[@id="MSG_Cancelspan"]'))).click()   

                    self.can_not_update_state[fa_case] = "Fault Duplication State is \"BLANK\'"   

        except:

            self.can_not_update_state[fa_case] = "Can't update in \"Fault Duplication State\""


    
    def upload_attachment(self,fa_case):

        try:

            WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//a[@id="top_simpleSearch"]'))).click()

            #16-Dec-2021 check web render

            try:
            
                sleep(3)
            
                check_web_render = WebDriverWait(self.driver, 3).until(ec.visibility_of_element_located((By.XPATH, '//h4[@id="searchResultHeader"]'))).text
                
                if "loading" in check_web_render.lower() or "results" in check_web_render.lower():
                
                    web_not_render = True
                
                    while (web_not_render):

                        try:
                    
                            print("Wait For Web Rendering")
                        
                            WebDriverWait(self.driver, 3).until(ec.visibility_of_element_located((By.XPATH, '//a[@id="top_simpleSearch"]'))).click()

                            sleep(3)
                            
                            check_web_render = WebDriverWait(self.driver, 3).until(ec.visibility_of_element_located((By.XPATH, '//h4[@id="searchResultHeader"]'))).text
                            
                            if "loading" not in check_web_render.lower() and "resutls" not in check_web_render.lower():
                            
                                web_not_render = False

                        except:

                            print("While check web not render is ERROR")

                            break
        
            except:
            
                print("Not found web loading status")         


            WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//div[@id="tabsDiv"]//li')))[-2].click()

            WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//a[@id="MSG_AddAttachment_10"]'))).click()


            #14-Dec-2021  loop upload cause text box is limit text


            
            add_files_list = []

            count_add_file = 0
            

            add_files = self.config["inspect_path"]

            for file in self.all_fa_case_inspect_files[fa_case]:

                

                if count_add_file == 6 :

                    add_files_list.append(add_files)

                    add_files = self.config["inspect_path"]

                    count_add_file=0

                add_files+="\""+ file +"\""

                count_add_file+=1


            if count_add_file < 6 :

                add_files_list.append(add_files)


            for addFile in add_files_list:

                WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//span[@class="fileinput-span lf-btn-span"]'))).click()

                not_found_browse_window = True

                while(not_found_browse_window):

                    try:

                        app = pywinauto.Application().connect(title="Open")

                        not_found_browse_window = False


                    except:

                        print("Wait for Browse Window")
                        sleep(1)
                        
                dlg = app.window(title="Open")

            
                dlg.Edit.type_keys(addFile+"{ENTER}")
                sleep(1)



                

            WebDriverWait(self.driver, 5).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="uploadFilesUM"]'))).click()

            downloading =True

            while(downloading):

                try:
                    upload_box = WebDriverWait(self.driver, 5).until(ec.visibility_of_element_located((By.XPATH, '//a[@id="lfuploadpalette_window_close"]'))).click()
                    print(upload_box)
                    downloading = False

                except:
                    print("Can not click \"upload\"")
                    WebDriverWait(self.driver, 5).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="uploadFilesUM"]'))).click()



        except:        

            self.can_not_update_state[fa_case] = "Can't update in \"Attach File State\" "





                

    def auto_update(self,fa_case):

        WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//input[@name="QUICKSEARCH_STRING"]'))).send_keys(Keys.CONTROL+'a',Keys.BACKSPACE) #clear 
        WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//input[@name="QUICKSEARCH_STRING"]'))).send_keys(fa_case)
        WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//a[@id="top_simpleSearch"]'))).click()
        
        if str(self.all_fa_case_result_sheet[fa_case]["Fault Duplication Test Plan"]).lower() != "nan":
            
            test_status = 'testable'
            
        else:
                        
            test_status = 'untestable'
            
        try :
            
            #check web render
            
            try:
            
                sleep(3)
            
                check_web_render = WebDriverWait(self.driver, 3).until(ec.visibility_of_element_located((By.XPATH, '//h4[@id="searchResultHeader"]'))).text
                
                if "loading" in check_web_render.lower() or "results" in check_web_render.lower():
                
                    web_not_render = True
                
                    while (web_not_render):

                        try:
                    
                            print("Wait For Web Rendering")
                        
                            WebDriverWait(self.driver, 3).until(ec.visibility_of_element_located((By.XPATH, '//a[@id="top_simpleSearch"]'))).click()

                            sleep(3)
                            
                            check_web_render = WebDriverWait(self.driver, 3).until(ec.visibility_of_element_located((By.XPATH, '//h4[@id="searchResultHeader"]'))).text
                            
                            if "loading" not in check_web_render.lower() and "resutls" not in check_web_render.lower():
                            
                                web_not_render = False

                        except:

                            print("While check web not render is ERROR")

                            break
        
            except:
            
                print("Not found web loading status")
                
                
            
            fa_case_status = WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//h2[@style="color:Blue;"]'))).text
            print(fa_case_status)

        except:

            self.can_not_update_state[fa_case] = "Can't update in \"FA-CASE'S PRIORITY\""

            print("Skip >>>",fa_case)

            return False

        if fa_case_status == "Awaiting Assignment" :

            print("Process >>> Awaiting Assignment")

            self.window_handles(fa_case)

            self.auto_cover_page_awaiting_assignment(fa_case,test_status)


        # old_case may be counterfiet or 5y
        elif fa_case_status == "Inspection & Review":

            print("Process >>> Inspection & Review")

            self.auto_cover_page_inspection_review(fa_case,test_status)


        elif fa_case_status == "Fault Duplication":

            print("Process >>> Fault Duplication")
            
            self.auto_cover_page_fault_duplication(fa_case,test_status)


        else:

            self.can_not_update_state[fa_case] = "Case Status is \"INCORRECT\""


        if self.all_fa_case_inspect_files[fa_case].__len__() > 0:

            self.upload_attachment(fa_case)



    def main(self):

        print('*** Start...  >>>  read_excel_data() ***')
        self.read_excel_data()
        print('*** Start...  >>>  scan_files() ***')
        self.scan_files()
        print('*** Start...  >>>  map_inspect_files() ***')
        print(self.map_inspect_files())
        print('*** Start...  >>>  login() ***')
        self.login()
        print('*** Start...  >>>  auto_update(fa_case) ***')
        for fa_case in self.all_fa_case_result_sheet:

            self.auto_update(fa_case)

        print('*** Finish...  >>>  Result ***')

        print(self.can_not_update_state)

        log = datetime.datetime.now()

        log = str(log)+ '\n'


        for i,j in self.can_not_update_state.items():

            log+=i+' : '+ j + '\n'

        with open('log.txt', 'a') as file:
            file.write(log)



        
if __name__ == '__main__':

    INSPECTION_MACRO().main()

    while(input("Please enter \"e\" to Exit !")).lower()[0] != 'e':

        pass



