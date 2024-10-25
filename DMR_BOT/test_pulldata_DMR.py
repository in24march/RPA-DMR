import logging
import pandas as pd
from datetime import datetime, timedelta
import time

from selenium import webdriver
from selenium.webdriver.support.select import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common import exceptions
import pandas as pd
import glob

from setting_dmr import *
# from map_rac import *
import setting_dmr

def tesk_pull(date):
    file_path = r"C:\test_put_file"
    new_path = r"C:\Users\chayakor\Desktop\DMR_BOT\BC"
    assign_id = ['671000155', '671000156', '671000157']
    for id in assign_id:
        for file_name in os.listdir(file_path):
            if id in file_name:
                print(f"File found: {file_name}")
                new_file_name = f"TEMP_{id}{os.path.splitext(file_name)[1]}"
                old_file = os.path.join(file_path, file_name)
                new_file = os.path.join(new_path, new_file_name)
                
                
                os.rename(old_file, new_file)
                print(f"บันทึกไฟล์ใหม่: {new_file}")
                
    
if __name__ == '__main__':
    date = datetime(2024, 8, 7)
    tesk_pull(date)