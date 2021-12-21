import pywinauto

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



#config.init file
my_config_parser = configparser.ConfigParser()
my_config_parser.read('config.ini')
config = { 

'inspect_email': my_config_parser.get('INSPECTION_LOGIN','email'),
'inspect_password': my_config_parser.get('INSPECTION_LOGIN','password'),
'inspect_path': my_config_parser.get('INSPECTION_LOGIN','inspect_path'),
'inherit_path': my_config_parser.get('INSPECTION_LOGIN','inherit_path'),
}

print(pywinauto.findwindows.find_windows())

app = pywinauto.Application().connect(title="Open")

print(app)

dlg = app.window(title="Open")



# app.dlg.print_control_identifiers()

resultInspectPath = "C:\\Users\\wisitl\\Documents\\PY\\Inspection_Macro\\Inspect_Files\\"

addedFiles = config['inspect_path'] 
for i in ["FA-0435090_JFQ2141302V_RX.pdf", "FA-0433961_ACA22450035_RX.pdf" ]:
    addedFiles+="\""+i+"\""


    
dlg.Edit.type_keys(addedFiles+"{ENTER}")




