from numpy import NaN, nan
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
import datetime

import logging


class TEST:

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

        return None


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


    def auto_update(self,fa_case):

        WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//input[@name="QUICKSEARCH_STRING"]'))).send_keys(Keys.CONTROL+'a',Keys.BACKSPACE) #clear 

        WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//input[@name="QUICKSEARCH_STRING"]'))).send_keys('FA-0493190')

        WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//a[@id="top_simpleSearch"]'))).click()

        while True :

                        try:

                            if WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//div[@class="column_one layout"]//h2'))).text == fa_case:

                                break

                        except:

                            if fa_case in WebDriverWait(self.driver, 3).until(ec.visibility_of_element_located((By.XPATH, '//div[@class="view_controls"]//h4'))).text:

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

        print(True)

        WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//span[@id="top_refreshspan"]'))).click()




        # while True :

        #     try:

        #         if WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//div[@class="column_one layout"]//h2'))).text == fa_case:

        #             print(True)

        #             print(WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//div[@class="column_one layout"]//h2'))).text)

        #             break

        #     except:

        #         print(False)




run = TEST()


run.login()

run.auto_update('FA-0493190')

