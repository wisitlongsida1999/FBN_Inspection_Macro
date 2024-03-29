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
import traceback



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

        self.df = pd.read_excel('Inspection_Template.xlsm',["qms","result"])

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

                    WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//button[@id="passcode"]'))).click()

                    clickable = True

                    WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//input[@name="passcode"]'))).send_keys(self.pass_code)

                    WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//button[@id="passcode"]'))).click()

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

        with open('log.txt', 'a') as file:

            date = str(datetime.datetime.now())+ '\n'

            file.write(date)

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

                self.logger.debug("No. Of window handles1: "+str(len(handles))+", "+str(handles))

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


        #window handle
        found_change_status_window = False
        count_open_change_status_window = 0
        while(not found_change_status_window):

            reset_handle = False
            sleep(1)#9-Dec-2021  add delay
            handles = self.driver.window_handles
            size = len(handles)
            self.logger.debug("No. Of window handles2: "+str(size)+' >>>  '+str(handles))

            for handle in handles:
                
                self.driver.switch_to.window(handle)
                window_title = self.driver.title
                if window_title == 'Change Status':
                    self.logger.debug("Change Status Window: "+window_title+' >>> '+str(handles))
                    found_change_status_window = True
                    break
                elif window_title == 'Application Error':
                    self.logger.error("Application Error Window: "+window_title+' >>> '+str(handles))

                    sleep(1)#9-Dec-2021  add delay
                    reset_handle = True
                    sleep(1)#9-Dec-2021  add delay
                    self.driver.close()

            if reset_handle:
                
                self.driver.switch_to.window(self.main_page)
                
                WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//span[@id="MSG_NextStatusspan"]'))).click()  #9-Dec-2021  visible to clickable
                
                WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="ewfinish"]'))).click()    #9-Dec-2021
                
            count_open_change_status_window+=3
            
            if count_open_change_status_window > 10:
            
                self.logger.critical("Can't open change_status Window: "+str(count_open_change_status_window)+" second")
                self.can_not_update_state[fa_case] = "Can't Open Change Status Window"
                
                self.log(fa_case, "Can't Open Change Status Window")

                break


        if count_open_change_status_window < 10:

            WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//a[@id="save"]'))).click()                

            count_close_change_status_window = 0
            while (len(self.driver.window_handles) > 1):
                sleep(1)
                count_close_change_status_window+=1
                self.logger.warning("Wait for Close Change Status Window: "+ str(count_close_change_status_window)+" second")
                if count_close_change_status_window > 10:
                    self.logger.critical("Can't Close Change Status Window: "+str(count_close_change_status_window))
                    self.can_not_update_state[fa_case] = "Can't Close Change Status Window"
                    self.log(fa_case, "Can't Close Change Status Window")
                    sleep(1)  #9-Dec-2021  add delay
                    self.driver.close()
                    sleep(1)  #9-Dec-2021  add delay
                    break       
        sleep(1)
        
        self.driver.switch_to.window(self.main_page)      
        
        

    def auto_update(self,fa_case):

        WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//input[@name="QUICKSEARCH_STRING"]'))).send_keys(Keys.CONTROL+'a',Keys.BACKSPACE) #clear 
        sleep(1)
        WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//input[@name="QUICKSEARCH_STRING"]'))).send_keys(fa_case)
        sleep(1)       
        WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//a[@id="top_simpleSearch"]'))).click()
        sleep(1) 
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
                        sleep(1) 

                    else:

                        raise Exception("FA Link is Missmatch")                    

                else:

                    self.logger.critical(fa_case+ " : Page is not rendering !!!")

                    WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//a[@id="top_simpleSearch"]'))).click()
                    sleep(1) 
        
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
            
            self.log(fa_case, "Case Status is \"INCORRECT\"")

            self.logger.critical(fa_case+ ' : '+"is not update due to Incorrect Status")


        if self.all_fa_case_inspect_files[fa_case].__len__() > 0:

            self.upload_attachment(fa_case)


    def auto_cover_page_awaiting_assignment(self,fa_case,status):

        try:

            WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//span[@id="MSG_Editspan"]'))).click()
            sleep(1)
            WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//span[@id="show_floater_R1_4859_0_displayspan"]'))).click()
            sleep(1)
            WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//input[@id="floater_search_text_R1_4859_0_display"]'))).send_keys(self.all_fa_case_result_sheet[fa_case]['Case Owner'])
            sleep(1)
            WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//span[@id="searchspan"]'))).click()
            sleep(1)
            elm_select = WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//td[@class=" yui-dt-string yui-dt-first"]')))

            ActionChains(self.driver).double_click(on_element = elm_select).perform()
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

            if status == 'testable':    
                # Is Potential Counterfeit   0 = Null, 1 = No, 2 = Yes            
                WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//select[@name="R1_2020_7"]//option')))[1].click()
                sleep(1)
                # Passed Visual Inspection
                WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//select[@name="R1_2021_7"]//option')))[2].click()
                sleep(1)
                WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="close_floater_R1_4859_0_display"]'))).click() 
                sleep(1)
                WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="MSG_Save"]'))).click()   
                sleep(3)
                self.window_handles(fa_case)
                sleep(1)
                WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//span[@id="MSG_Editspan"]'))).click()
                sleep(1)
                #Fault Duplication Test Plan
                WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//textarea[@name="R1_1332_7"]'))).send_keys(self.all_fa_case_result_sheet[fa_case]['Fault Duplication Test Plan'])
                sleep(1)

            elif status == 'untestable':

                # Is Potential Counterfeit   0 = Null, 1 = No, 2 = Yes       
                if self.all_fa_case_result_sheet[fa_case]['Is Potential Counterfeit'].lower() == "yes":

                    WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//select[@name="R1_2020_7"]//option')))[2].click()
                    sleep(1)
                elif self.all_fa_case_result_sheet[fa_case]['Is Potential Counterfeit'].lower() == "no":

                    WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//select[@name="R1_2020_7"]//option')))[1].click()
                    sleep(1)
                else:

                    WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//select[@name="R1_2020_7"]//option')))[0].click()
                    sleep(1)
                # Passed Visual Inspection
                if "pass" in self.all_fa_case_result_sheet[fa_case]['Visual Inspection Code'].lower():

                    WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//select[@name="R1_2021_7"]//option')))[2].click()
                    sleep(1)
                elif "fail" in self.all_fa_case_result_sheet[fa_case]['Visual Inspection Code'].lower():

                    WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//select[@name="R1_2021_7"]//option')))[1].click()
                    sleep(1)
                else:

                    WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//select[@name="R1_2021_7"]//option')))[0].click()
                    sleep(1)
                WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="close_floater_R1_4859_0_display"]'))).click() 
                sleep(1)
            WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="MSG_Save"]'))).click()      
            sleep(3)

        except:

            self.can_not_update_state[fa_case] = "Can't update in \"Awaiting Assignment State\""

            self.log(fa_case, "Can't update in \"Awaiting Assignment State\"")

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
            WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//span[@id="MSG_Editspan"]'))).click()
            sleep(1)    
            WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//span[@id="show_floater_R1_4859_0_displayspan"]'))).click()
            sleep(1)
            WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//input[@id="floater_search_text_R1_4859_0_display"]'))).send_keys(self.all_fa_case_result_sheet[fa_case]['Case Owner'])
            sleep(1)
            WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//span[@id="searchspan"]'))).click()
            sleep(1)
            elm_select = WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//td[@class=" yui-dt-string yui-dt-first"]')))
            ActionChains(self.driver).double_click(on_element = elm_select).perform()
            sleep(1)

            old_summary = WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//textarea[@name="R1_2017_7"]'))).text

            if status == 'testable':

                #Part Inspection Summary
                if old_summary.__len__() > 0:

                    if self.all_fa_case_result_sheet[fa_case]['Part Inspection Summary'] not in old_summary :

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
                sleep(1)
                WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="MSG_Save"]'))).click() 
                sleep(3)
                self.window_handles(fa_case)
                WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//span[@id="MSG_Editspan"]'))).click()
                sleep(1)
                WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//textarea[@name="R1_1332_7"]'))).send_keys(self.all_fa_case_result_sheet[fa_case]['Fault Duplication Test Plan']) 
                sleep(1)

            elif status == 'untestable':
                
                if old_summary.__len__() > 0:

                    if self.all_fa_case_result_sheet[fa_case]['Part Inspection Summary'] not in old_summary:

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
                    sleep(1)
                elif self.all_fa_case_result_sheet[fa_case]['Is Potential Counterfeit'].lower() == "no":

                    WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//select[@name="R1_2020_7"]//option')))[1].click()
                    sleep(1)
                else:

                    WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//select[@name="R1_2020_7"]//option')))[0].click()
                    sleep(1)
                # Passed Visual Inspection
                if "pass" in self.all_fa_case_result_sheet[fa_case]['Visual Inspection Code'].lower():

                    WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//select[@name="R1_2021_7"]//option')))[2].click()
                    sleep(1)
                elif "fail" in self.all_fa_case_result_sheet[fa_case]['Visual Inspection Code'].lower():

                    WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//select[@name="R1_2021_7"]//option')))[1].click()
                    sleep(1)
                else:

                    WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//select[@name="R1_2021_7"]//option')))[0].click()
                    sleep(1)
                WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="close_floater_R1_4859_0_display"]'))).click() 
                sleep(1)

            WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="MSG_Save"]'))).click() 
            sleep(3)


        except:

            self.can_not_update_state[fa_case] = "Can't update in \"Inspection Review State\""

            self.log(fa_case, "Can't update in \"Inspection Review State\"")

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
            sleep(1)
            old_summary = WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//textarea[@name="R1_2017_7"]'))).text
            old_test_plan = WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//textarea[@name="R1_1332_7"]'))).text

            if status == 'testable':

                #Part Inspection Summary
                if old_summary.__len__() > 0:

                    if self.all_fa_case_result_sheet[fa_case]['Part Inspection Summary'] not in old_summary :

                        WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//textarea[@name="R1_2017_7"]'))).send_keys("\n\n"+self.all_fa_case_result_sheet[fa_case]['Part Inspection Summary'])
                        sleep(1)
                    # Is Potential Counterfeit   0 = Null, 1 = No, 2 = Yes            
                    WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//select[@name="R1_2020_7"]//option')))[1].click()
                    sleep(1)
                    # Passed Visual Inspection
                    WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//select[@name="R1_2021_7"]//option')))[2].click()
                    sleep(1)
                    
                    if self.all_fa_case_result_sheet[fa_case]['Fault Duplication Test Plan'] not in old_test_plan:

                        WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//textarea[@name="R1_1332_7"]'))).send_keys("\n\n"+self.all_fa_case_result_sheet[fa_case]['Fault Duplication Test Plan'])
                        sleep(1)
                    WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="MSG_Save"]'))).click()
                    sleep(3)

                else:

                    WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//span[@id="MSG_Cancelspan"]'))).click()   
                    sleep(1)
                    self.logger.critical("Fault Duplication State is \"BLANK\'")

                    self.can_not_update_state[fa_case] = "Fault Duplication State is \"BLANK\'"

                    self.log(fa_case, "Fault Duplication State is \"BLANK\'")


            elif status == 'untestable':
                
                if old_summary.__len__() > 0:

                    if self.all_fa_case_result_sheet[fa_case]['Part Inspection Summary'] not in old_summary :
                    #Part Inspection Summary
                        WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//textarea[@name="R1_2017_7"]'))).send_keys("\n\n"+self.all_fa_case_result_sheet[fa_case]['Part Inspection Summary'])
                        sleep(1)
                    WebDriverWait(self.driver, 10).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="MSG_Save"]'))).click()                     
                    sleep(3)
                else :

                    WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//span[@id="MSG_Cancelspan"]'))).click()  
                    sleep(1)
                    self.can_not_update_state[fa_case] = "Fault Duplication State is \"BLANK\'"   

                    self.log(fa_case, "Fault Duplication State is \"BLANK\'" )

                    self.logger.critical("Fault Duplication State is \"BLANK\'")

        except:

            self.can_not_update_state[fa_case] = "Can't update in \"Fault Duplication State\""

            self.log(fa_case, "Can't update in \"Fault Duplication State\"")

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

            WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//a[@id="top_simpleSearch"]'))).click()

            sleep(1) 
            
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
                            sleep(1) 

                        else:

                            raise Exception("FA Link is Missmatch")                    

                    else:

                        self.logger.critical(fa_case+ " : Page is not rendering !!!")

                        WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//a[@id="top_simpleSearch"]'))).click()
                        sleep(1) 
            
            sleep(1)
            WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//div[@id="tabsDiv"]//li')))[-2].click()
            sleep(1)


            # check duplicated
            file_name_ls = []

            rows_exact = int(WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//strong[@id="totalCount_ATTACHMENTS_FILELIST"]'))).text)

            rows = WebDriverWait(self.driver, 10).until(ec.visibility_of_all_elements_located((By.XPATH, '//tr[@class="GMDataRow"]')))

            rows_len = len(rows)

            self.logger.info("Exact rows >>> "+str(rows_exact)+" Rows number >>> "+str(rows_len))

            if int(rows_exact)*2 != rows_len:

                self.logger.critical(fa_case + ': Rows number does not match !!!')

                self.err.update({fa_case: 'Rows number does not match'})

            row_start = int(rows_len/2)
            
            for i in range(row_start, rows_len):

                row = rows[i]

                self.logger.debug(row)

                entries = row.find_elements(By.TAG_NAME,'td')

                file_name = entries[3].text.strip()

                file_name_ls.append(file_name)

                self.logger.debug("File Name >>> "+file_name)

            self.logger.info("List of files that was uploaded for "+fa_case+' >>> '+ str(file_name_ls))

            add_files_list = []

            count_add_file = 0

            for name in file_name_ls:

                if name in self.all_fa_case_inspect_files[fa_case]:

                    self.all_fa_case_inspect_files[fa_case].remove(name)

            if self.all_fa_case_inspect_files[fa_case].__len__() > 0:

                WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//a[@id="MSG_AddAttachment_10"]'))).click()

                sleep(1)
            
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
                    sleep(1)
                    not_found_browse_window = True

                    while(not_found_browse_window):

                        try:

                            app = pywinauto.Application().connect(title="Open")

                            not_found_browse_window = False

                        except:

                            self.logger.warning("Wait for Browse Window")

                            sleep(1)
                            
                    dlg = app.window(title="Open")
                    
                    sleep(1)

                    dlg.Edit.type_keys(addFile+"{ENTER}")
                    
                    sleep(1)


                WebDriverWait(self.driver, 5).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="uploadFilesUM"]'))).click()
                sleep(1)
                downloading =True

                while(downloading):

                    try:

                        close_upload_box = WebDriverWait(self.driver, 5).until(ec.visibility_of_element_located((By.XPATH, '//a[@id="lfuploadpalette_window_close"]'))).click()
                        sleep(1)
                        self.logger.debug(close_upload_box)

                        downloading = False

                    except:

                        self.logger.critical("Can not click \"upload\"")

                        WebDriverWait(self.driver, 5).until(ec.element_to_be_clickable((By.XPATH, '//a[@id="uploadFilesUM"]'))).click()
                        sleep(1)


        except:        

            self.can_not_update_state[fa_case] = "Can't update in \"Attachment State\""

            self.log(fa_case, "Can't update in \"Attachment State\"")

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


    def log(self,fa_case,err):

        with open('log.txt', 'a') as file:

            res = fa_case+ ' : '+ err +'\n'

            file.write(res)
        
        return True


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



if __name__ == '__main__':

    

    run = INSPECTION_MACRO()


    try:

        run.main()

        if len(run.can_not_update_state) > 0:

            result = run.log
        
        else:

            result = 'Complete All FA Cases !!!'

        while True:
        
            try:

                en = pyautogui.password(text=result, title='Result', mask='☠') 

                if en == '509357':

                    break

                run.logger.critical('Incorrect EN >>> '+en)
                
            except:
            
                run.logger.critical('Incorrect EN >>> Error Result Box !')
        
        run.logger.info('End of program >>> '+en)

        run.driver.quit()

    except:

        tb = traceback.format_exc()

        run.logger.critical(tb)
