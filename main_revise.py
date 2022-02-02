import logging
import datetime
import configparser
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.by import By
from time import sleep
import pandas as pd
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import os
import sys
import pywinauto
import pyautogui




class INSPECTION_MACRO:

    def __init__(self):

        # create logger
        self.logger = logging.getLogger(__name__)
        self.logger.setLevel(logging.DEBUG)

        # create console handler
        ch = logging.StreamHandler()

        #create file handler 
        date = str(datetime.datetime.now().strftime('%d-%b-%Y %H_%M_%S %p'))
        fh = logging.FileHandler('debug\\{}.log'.format(date),encoding='utf-8')

        # create formatter
        formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s',datefmt='%d/%b/%Y %I:%M:%S %p')

        # add formatter to ch
        ch.setFormatter(formatter)

        #add formatter to fh
        fh.setFormatter(formatter)

        # add ch to logger
        self.logger.addHandler(ch)

        #add fh to logger
        self.logger.addHandler(fh)


        #config.init file
        self.my_config_parser = configparser.ConfigParser()

        self.my_config_parser.read('config\\config.ini')

        self.config = { 

        'inspect_email': self.my_config_parser.get('INSPECTION_LOGIN','email'),
        'inspect_password': self.my_config_parser.get('INSPECTION_LOGIN','password'),
        'inspect_path': self.my_config_parser.get('INSPECTION_LOGIN','inspect_path'),

        }

        #config pd with excel
        self.all_fa_sn_qms_sheet = {}
        self.all_fa_case_result_sheet = {}
        
        # self.raw_sn = ""
        
        #config read all files name 
        self.all_fa_case_inspect_files = {}

        #debug
        self.can_not_update_state = {}



        en = pyautogui.password(text='Please Enter Your \"EN\"', title='Step 1 : EN', mask='☠')

        self.logger.info('EN : '+en)

        if en == '509357':

            self.pass_code = pyautogui.prompt(text='Please Enter \"Pass Code\"', title='Step 2 : Pass Code')

        else:

            pyautogui.alert(text='Incorrect EN !!!',title='Error!',button='Exit')

            sys.exit()


        return None



    def read_excel_data(self):

        self.df = pd.read_excel('Inspection_Template.xlsx',["qms","result"])

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

            self.logger.info("Wait for two-factor authentication render:"+str(count_render_2fa))


        push_login_btn = WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//button[@type="submit"]')))

        push_login_btn_text = push_login_btn.text

        if push_login_btn_text == "Send Me a Push":

            clickable = False

            while(not clickable):

                try:

                    WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//button[@type="submit"]'))).click()

                    clickable = True

                except:

                    self.logger.warning("Can't Clickable'")
                
        else:
            self.logger.info("Not found \"Send Me a Push\" button")


        two_fa_url=self.driver.current_url

        count_duo_pass = 0

        while(two_fa_url==self.driver.current_url):

            sleep(1)

            count_duo_pass+=1

            self.logger.info("Wait for count_duo_pass:"+str(count_duo_pass))

            if count_duo_pass == 30:

                self.logger.warning("!!! LOGIN TIMEOUT !!!")

                self.driver.quit()

                sys.exit()

        self.logger.info("Login to QIS is success!!!")

        sleep(1)

        self.driver.get("https://www-plmprd.cisco.com/Agile/")

        self.main_page = self.driver.current_window_handle

        self.logger.debug("Main Page:"+self.main_page)

        handles = self.driver.window_handles

        for handle in handles:

            sleep(1)

            self.driver.switch_to.window(handle)

            if self.main_page != self.driver.current_window_handle:

                self.driver.close()

        self.driver.switch_to.window(self.main_page)

        self.driver.maximize_window()

        return True


    def window_handles(self,fa_case):

        not_reset_audit = True

        while(not_reset_audit):

            try :   

                WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//span[@id="MSG_NextStatusspan"]'))).click()

                WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//a[@id="ewfinish"]'))).click()

                not_reset_audit = False
                
            except:

                reset_handle = False

                handles = self.driver.window_handles

                self.logger.debug("No. Of window handles1: "+str(len(handles))+", "+handles)

                for handle in handles:

                    self.driver.switch_to.window(handle)

                    window_title = self.driver.title

                    if window_title == 'Application Error':

                        self.logger.error("Application Error Window: "+window_title+" , " +handle)

                        reset_handle = True

                        self.driver.close()

                if reset_handle:
                    
                    self.driver.switch_to.window(self.main_page)

                    WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//a[@id="top_simpleSearch"]'))).click()


    def auto_update(self,fa_case):

        WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//input[@name="QUICKSEARCH_STRING"]'))).send_keys(Keys.CONTROL+'a',Keys.BACKSPACE) #clear 

        WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//input[@name="QUICKSEARCH_STRING"]'))).send_keys(fa_case)

        WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//a[@id="top_simpleSearch"]'))).click()
        
        if str(self.all_fa_case_result_sheet[fa_case]["Fault Duplication Test Plan"]).lower() != "nan":
            
            test_status = 'testable'
            
        else:
                        
            test_status = 'untestable'
        

            
        #check web render
        while True :

            try:

                if WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//div[@class="column_one layout"]//h2'))).text == fa_case:

                    break

            except:

                if fa_case in WebDriverWait(self.driver, 3).until(ec.visibility_of_element_located((By.XPATH, '//h4[@id="searchResultHeader"]'))).text:

                    rows = WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//tr[@class="GMDataRow"]')))

                    row = rows[-1]

                    self.logger.debug(row)

                    entries = row.find_elements(By.TAG_NAME,'td')

                    fa_link = entries[3]

                    self.logger.info("FA Link >>> "+fa_link.text)

                    if fa_case == fa_link.text.strip():

                        fa_link.click()          

                    else:

                        raise Exception("FA Link is Missmatch")                    

                else:

                    self.logger.critical(fa_case+ " : Page is not rendering !!!")

                    WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//a[@id="top_simpleSearch"]'))).click()

        
        fa_case_status = WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//h2[@style="color:Blue;"]'))).text

        self.logger.info(fa_case_status)


        if fa_case_status == "Awaiting Assignment" :

            self.logger.debug(fa_case+ ' : '+"Process >>> Awaiting Assignment")

            self.window_handles(fa_case)

            self.auto_cover_page_awaiting_assignment(fa_case,test_status)

        # old_case may be counterfiet or 5y
        elif fa_case_status == "Inspection & Review":

            self.logger.debug(fa_case+ ' : '+"Process >>> Inspection & Review")

            self.auto_cover_page_inspection_review(fa_case,test_status)

        elif fa_case_status == "Fault Duplication":

            self.logger.debug(fa_case+ ' : '+"Process >>> Fault Duplication")
            
            self.auto_cover_page_fault_duplication(fa_case,test_status)

        else:

            self.can_not_update_state[fa_case] = "Case Status is \"INCORRECT\""

            self.logger.critical(fa_case+ ' : '+"is not update due to Incorrect Status")


        if self.all_fa_case_inspect_files[fa_case].__len__() > 0:

            self.upload_attachment(fa_case)


    def auto_cover_page_awaiting_assignment(self,fa_case,status):

        try:

            if status == 'testable':
                
                WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//span[@id="MSG_Editspan"]'))).click()

                WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//span[@id="show_floater_R1_4859_0_displayspan"]'))).click()

                WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//input[@id="floater_search_text_R1_4859_0_display"]'))).send_keys(self.all_fa_case_result_sheet[fa_case]['Case Owner'])

                WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//span[@id="searchspan"]'))).click()

                elm_select = WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//td[@class=" yui-dt-string yui-dt-first"]')))

                ActionChains(self.driver).double_click(on_element = elm_select).perform()

                #Part Inspection Summary
                WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//textarea[@name="R1_2017_7"]'))).send_keys(self.all_fa_case_result_sheet[fa_case]['Part Inspection Summary'])

                # Case Review Summary
                WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//textarea[@name="R1_2018_7"]'))).send_keys("-")

                # Part History Investigation Summary
                WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//textarea[@name="R1_2019_7"]'))).send_keys("-")

                # Is Potential Counterfeit   0 = Null, 1 = No, 2 = Yes            
                WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//select[@name="R1_2020_7"]//option')))[1].click()

                # Passed Visual Inspection
                WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//select[@name="R1_2021_7"]//option')))[2].click()

                WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="close_floater_R1_4859_0_display"]'))).click() 

                WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="MSG_Save"]'))).click()   

                self.window_handles(fa_case)

                WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//span[@id="MSG_Editspan"]'))).click()

                #Fault Duplication Test Plan
                WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//textarea[@name="R1_1332_7"]'))).send_keys(self.all_fa_case_result_sheet[fa_case]['Fault Duplication Test Plan'])
                

            elif status == 'untestable':

                WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//span[@id="MSG_Editspan"]'))).click()

                WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//span[@id="show_floater_R1_4859_0_displayspan"]'))).click()

                WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//input[@id="floater_search_text_R1_4859_0_display"]'))).send_keys(self.all_fa_case_result_sheet[fa_case]['Case Owner'])

                WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//span[@id="searchspan"]'))).click()

                elm_select = WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//td[@class=" yui-dt-string yui-dt-first"]')))

                ActionChains(self.driver).double_click(on_element = elm_select).perform()

                #Part Inspection Summary
                WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//textarea[@name="R1_2017_7"]'))).send_keys(self.all_fa_case_result_sheet[fa_case]['Part Inspection Summary'])

                # Case Review Summary
                WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//textarea[@name="R1_2018_7"]'))).send_keys("-")

                # Part History Investigation Summary
                WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//textarea[@name="R1_2019_7"]'))).send_keys("-")

                # Is Potential Counterfeit   0 = Null, 1 = No, 2 = Yes       
                if self.all_fa_case_result_sheet[fa_case]['Is Potential Counterfeit'].lower() == "yes":

                    WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//select[@name="R1_2020_7"]//option')))[2].click()

                elif self.all_fa_case_result_sheet[fa_case]['Is Potential Counterfeit'].lower() == "no":

                    WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//select[@name="R1_2020_7"]//option')))[1].click()

                else:

                    WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//select[@name="R1_2020_7"]//option')))[0].click()

                # Passed Visual Inspection
                if "pass" in self.all_fa_case_result_sheet[fa_case]['Visual Inspection Code'].lower():

                    WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//select[@name="R1_2021_7"]//option')))[2].click()

                elif "fail" in self.all_fa_case_result_sheet[fa_case]['Visual Inspection Code'].lower():

                    WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//select[@name="R1_2021_7"]//option')))[1].click()

                else:

                    WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//select[@name="R1_2021_7"]//option')))[0].click()

                WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="close_floater_R1_4859_0_display"]'))).click() 

            WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="MSG_Save"]'))).click()      


        except:

            self.can_not_update_state[fa_case] = "Can't update in \"Awaiting Assignment State\""

            self.logger.critical(fa_case+" : is not complete in Awaiting Assignment Process")

            # handle with alert
            try:
                WebDriverWait(self.driver, 5).until(ec.alert_is_present())

                alert = self.driver.switch_to.alert

                alert.accept()

                self.logger.debug(fa_case+" >>> alert accepted")

            except:

                self.logger.debug(fa_case+" >>> no alert")

            #handle with refresh browser
                try:

                    if "refresh" in WebDriverWait(self.driver, 5).until(ec.visibility_of_element_located((By.XPATH, '//p[@id="dms_msg"]'))).text:

                        WebDriverWait(self.driver, 5).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="MSG_Cancel"]'))).click() 

                        self.logger.critical(fa_case+" >>> Refresh Browser")

                except:

                    self.logger.critical(fa_case+" >>> Not Refresh Browser")


    def auto_cover_page_inspection_review(self,fa_case,status):

        try :

            if status == 'testable':

                WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//span[@id="MSG_Editspan"]'))).click()

                WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//span[@id="show_floater_R1_4859_0_displayspan"]'))).click()

                WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//input[@id="floater_search_text_R1_4859_0_display"]'))).send_keys(self.all_fa_case_result_sheet[fa_case]['Case Owner'])

                WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//span[@id="searchspan"]'))).click()

                elm_select = WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//td[@class=" yui-dt-string yui-dt-first"]')))

                ActionChains(self.driver).double_click(on_element = elm_select).perform()

                #Part Inspection Summary
                old_summary = WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//textarea[@name="R1_2017_7"]'))).text

                if old_summary.__len__() > 0:

                    WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//textarea[@name="R1_2017_7"]'))).send_keys("\n\n"+self.all_fa_case_result_sheet[fa_case]['Part Inspection Summary'])

                else:

                    WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//textarea[@name="R1_2017_7"]'))).send_keys(self.all_fa_case_result_sheet[fa_case]['Part Inspection Summary'])

                    # Case Review Summary
                    WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//textarea[@name="R1_2018_7"]'))).send_keys("-")

                    # Part History Investigation Summary
                    WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//textarea[@name="R1_2019_7"]'))).send_keys("-")

                # Is Potential Counterfeit   0 = Null, 1 = No, 2 = Yes            
                WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//select[@name="R1_2020_7"]//option')))[1].click()

                # Passed Visual Inspection
                WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//select[@name="R1_2021_7"]//option')))[2].click()

                WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="close_floater_R1_4859_0_display"]'))).click() 

                WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="MSG_Save"]'))).click() 

                self.window_handles(fa_case)

                WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//span[@id="MSG_Editspan"]'))).click()

                WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//textarea[@name="R1_1332_7"]'))).send_keys(self.all_fa_case_result_sheet[fa_case]['Fault Duplication Test Plan']) 


            elif status == 'untestable':

                WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//span[@id="MSG_Editspan"]'))).click()

                WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//span[@id="show_floater_R1_4859_0_displayspan"]'))).click()

                WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//input[@id="floater_search_text_R1_4859_0_display"]'))).send_keys(self.all_fa_case_result_sheet[fa_case]['Case Owner'])

                WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//span[@id="searchspan"]'))).click()

                elm_select = WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//td[@class=" yui-dt-string yui-dt-first"]')))

                ActionChains(self.driver).double_click(on_element = elm_select).perform()

                old_summary = WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//textarea[@name="R1_2017_7"]'))).text
                
                if old_summary.__len__() > 0:

                    #Part Inspection Summary
                    WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//textarea[@name="R1_2017_7"]'))).send_keys("\n\n"+self.all_fa_case_result_sheet[fa_case]['Part Inspection Summary'])

                else :

                    #Part Inspection Summary
                    WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//textarea[@name="R1_2017_7"]'))).send_keys(self.all_fa_case_result_sheet[fa_case]['Part Inspection Summary'])

                # Is Potential Counterfeit   0 = Null, 1 = No, 2 = Yes       
                if self.all_fa_case_result_sheet[fa_case]['Is Potential Counterfeit'].lower() == "yes":

                    WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//select[@name="R1_2020_7"]//option')))[2].click()

                elif self.all_fa_case_result_sheet[fa_case]['Is Potential Counterfeit'].lower() == "no":

                    WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//select[@name="R1_2020_7"]//option')))[1].click()

                else:

                    WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//select[@name="R1_2020_7"]//option')))[0].click()

                # Passed Visual Inspection
                if "pass" in self.all_fa_case_result_sheet[fa_case]['Visual Inspection Code'].lower():

                    WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//select[@name="R1_2021_7"]//option')))[2].click()

                elif "fail" in self.all_fa_case_result_sheet[fa_case]['Visual Inspection Code'].lower():

                    WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//select[@name="R1_2021_7"]//option')))[1].click()

                else:

                    WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//select[@name="R1_2021_7"]//option')))[0].click()

                WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="close_floater_R1_4859_0_display"]'))).click() 


            WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="MSG_Save"]'))).click() 



        except:

            self.can_not_update_state[fa_case] = "Can't update in \"Inspection Review State\""

            self.logger.critical(fa_case+" : is not complete in Inspection Review Process")

            # handle with alert

            try:
                WebDriverWait(self.driver, 5).until(ec.alert_is_present())

                alert = self.driver.switch_to.alert

                alert.accept()

                self.logger.debug(fa_case+" >>> alert accepted")

            except:

                self.logger.debug(fa_case+" >>> no alert")

            #handle with refresh browser
                try:

                    if "refresh" in WebDriverWait(self.driver, 5).until(ec.visibility_of_element_located((By.XPATH, '//p[@id="dms_msg"]'))).text:

                        WebDriverWait(self.driver, 5).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="MSG_Cancel"]'))).click() 

                        self.logger.critical(fa_case+" >>> Refresh Browser")

                except:

                    self.logger.critical(fa_case+" >>> Not Refresh Browser")



    def auto_cover_page_fault_duplication(self,fa_case,status):

        try:

            WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//span[@id="MSG_Editspan"]'))).click()

            if status == 'testable':

                #Part Inspection Summary
                old_summary = WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//textarea[@name="R1_2017_7"]'))).text
                
                if old_summary.__len__() > 0:

                    WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//textarea[@name="R1_2017_7"]'))).send_keys("\n\n"+self.all_fa_case_result_sheet[fa_case]['Part Inspection Summary'])

                    # Is Potential Counterfeit   0 = Null, 1 = No, 2 = Yes            
                    WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//select[@name="R1_2020_7"]//option')))[1].click()

                    # Passed Visual Inspection
                    WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//select[@name="R1_2021_7"]//option')))[2].click()
                
                    WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//textarea[@name="R1_1332_7"]'))).send_keys("\n\n"+self.all_fa_case_result_sheet[fa_case]['Fault Duplication Test Plan'])

                    WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="MSG_Save"]'))).click()

                else:

                    WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//span[@id="MSG_Cancelspan"]'))).click()   

                    self.logger.critical("Fault Duplication State is \"BLANK\'")

                    self.can_not_update_state[fa_case] = "Fault Duplication State is \"BLANK\'"


            elif status == 'untestable':

                old_summary = WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//textarea[@name="R1_2017_7"]'))).text
                
                if old_summary.__len__() > 0:

                    #Part Inspection Summary
                    WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//textarea[@name="R1_2017_7"]'))).send_keys("\n\n"+self.all_fa_case_result_sheet[fa_case]['Part Inspection Summary'])

                    WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="MSG_Save"]'))).click()                     

                else :

                    WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//span[@id="MSG_Cancelspan"]'))).click()  

                    self.can_not_update_state[fa_case] = "Fault Duplication State is \"BLANK\'"   

                    self.logger.critical("Fault Duplication State is \"BLANK\'")

        except:

            self.can_not_update_state[fa_case] = "Can't update in \"Fault Duplication State\""

            self.logger.critical(fa_case+" : is not complete in Fault Duplication Process")

            # handle with alert
            try:
                WebDriverWait(self.driver, 5).until(ec.alert_is_present())

                alert = self.driver.switch_to.alert

                alert.accept()

                self.logger.debug(fa_case+" >>> alert accepted")

            except:

                self.logger.debug(fa_case+" >>> no alert")

            #handle with refresh browser
                try:

                    if "refresh" in WebDriverWait(self.driver, 5).until(ec.visibility_of_element_located((By.XPATH, '//p[@id="dms_msg"]'))).text:

                        WebDriverWait(self.driver, 5).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="MSG_Cancel"]'))).click() 

                        self.logger.critical(fa_case+" >>> Refresh Browser")

                except:

                    self.logger.critical(fa_case+" >>> Not Refresh Browser")


    def upload_attachment(self,fa_case):

        try:

            WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//span[@id="top_refreshspan"]'))).click()

            WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//div[@id="tabsDiv"]//li')))[-2].click()

            WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//a[@id="MSG_AddAttachment_10"]'))).click()

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


            if count_add_file <= 6 :

                add_files_list.append(add_files)


            for addFile in add_files_list:

                WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//span[@class="fileinput-span lf-btn-span"]'))).click()

                not_found_browse_window = True

                while(not_found_browse_window):

                    try:

                        app = pywinauto.Application().connect(title="Open")

                        not_found_browse_window = False

                    except:

                        self.logger.warning("Wait for Browse Window")

                        sleep(1)
                        
                dlg = app.window(title="Open")

                dlg.Edit.type_keys(addFile+"{ENTER}")


            WebDriverWait(self.driver, 5).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="uploadFilesUM"]'))).click()

            downloading =True

            while(downloading):

                try:

                    close_upload_box = WebDriverWait(self.driver, 5).until(ec.visibility_of_element_located((By.XPATH, '//a[@id="lfuploadpalette_window_close"]'))).click()

                    self.logger.debug(close_upload_box)

                    downloading = False

                except:

                    self.logger.critical("Can not click \"upload\"")

                    WebDriverWait(self.driver, 5).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="uploadFilesUM"]'))).click()



        except:        

            self.can_not_update_state[fa_case] = "Can't update in \"Attachment State\""

            self.logger.critical(fa_case+" : is not complete in Attached Process")

            # handle with alert
            try:
                WebDriverWait(self.driver, 5).until(ec.alert_is_present())

                alert = self.driver.switch_to.alert

                alert.accept()

                self.logger.debug(fa_case+" >>> alert accepted")

            except:

                self.logger.debug(fa_case+" >>> no alert")

            #handle with refresh browser
                try:

                    if "refresh" in WebDriverWait(self.driver, 5).until(ec.visibility_of_element_located((By.XPATH, '//p[@id="dms_msg"]'))).text:

                        WebDriverWait(self.driver, 5).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="MSG_Cancel"]'))).click() 

                        self.logger.critical(fa_case+" >>> Refresh Browser")

                except:

                    self.logger.critical(fa_case+" >>> Not Refresh Browser")


    def main(self):

        self.logger.info('*** Start...  >>>  read_excel_data() ***')

        self.read_excel_data()

        self.logger.info('*** Start...  >>>  scan_files() ***')

        self.scan_files()

        self.logger.info('*** Start...  >>>  map_inspect_files() ***')

        self.logger.info(self.map_inspect_files())

        self.logger.info('*** Start...  >>>  login() ***')

        self.login()

        self.logger.info('*** Start...  >>>  auto_update(fa_case) ***')

        for fa_case in self.all_fa_case_result_sheet:

            self.auto_update(fa_case)

        self.logger.info('*** Finish...  >>>  Result ***')

        self.logger.critical(self.can_not_update_state)

        log = datetime.datetime.now()

        log = str(log)+ '\n'

        for i,j in self.can_not_update_state.items():

            log+=i+' : '+ j + '\n'

        with open('log.txt', 'a') as file:

            file.write(log)






if __name__ == '__main__':

    run = INSPECTION_MACRO()

    run.main()

    if len(run.can_not_update_state) > 0:

        result = run.log
    
    else:

        result = 'Complete All FA Cases !!!'

    while True:

        en = pyautogui.password(text=result, title='Result', mask='☠') 

        if en == '509357':

            break

        run.logger.critical('Incorrect EN >>> '+en)
    
    run.logger.info('End of program >>> '+en)