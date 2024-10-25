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
from selenium.common.exceptions import NoAlertPresentException, TimeoutException, StaleElementReferenceException
import shutil
import time
import pandas as pd
import glob

from setting_dmr import *
# from map_rac import *
import setting_dmr

def login():
    try:
        option = webdriver.ChromeOptions()
        pref = {'download.default_directory': Data_path}
        option.add_experimental_option('prefs', pref)
        option.add_argument('ignore-certificate-errors')
        option.add_argument("--no-sandbox")
        option.add_experimental_option("detach", True)
        option.add_argument("--disable-dev-shm-usage")
        driver = webdriver.Chrome(options= option)
        driver.implicitly_wait(30)
    except Exception as e:
        print(e)
    
    login_ins = setting_dmr.Login()
    login_ins.login()
    login_ins.webdriver()
    driver.get(login_ins.dmr)
    print("Open webdriver")
    driver. maximize_window()
    driver.find_element(By.NAME, 'txtUserid').send_keys(login_ins.user)
    driver.find_element(By.NAME, 'txtPassword').send_keys(login_ins.password)
    driver.find_element(By.NAME, 'btnLogin').click()
    
    return driver



def get_data(date):
    finder_file = find_file(Master_path)
    file_data = finder_file.file_last_time()

    # อ่านข้อมูลจาก Excel
    df = pd.read_excel(file_data)
    print(df.dtypes)

    # กำหนดวันที่ปัจจุบัน
    date_today = datetime.now()
    today = pd.to_datetime(date)
    today_date = pd.to_datetime(date_today)
    as_of_day_str = str(today_date.day)
    as_of_month_str = str(today_date.month)
    as_of_year_str = str(today_date.year)
    
    # date_range = [today, today - timedelta(days=1), today - timedelta(days=2)]

    # ตรวจสอบว่า 'Assign Date' <= today <= 'Unassign Date'
    condition = (df['Assign Date'] <= today) & (df['Unassign Date'] >= today)
    # condition = (df['Assign Date'] <= today) & (df['Unassign Date'] >= date_range[-1])

    # ดึงข้อมูลที่ตรงตามเงื่อนไข
    result = df.loc[condition, ['Assign ID', 'Assign Date', 'Unassign Date']].values.tolist()

    assign_ids_by_date = {}
    id_list = []

    # วนลูปเพื่อจัดเก็บข้อมูลใน dictionary
    for assign_id, assign_date, unassign_date in result:
        assign_date_str = assign_date.strftime('%Y-%m-%d')
        unassign_date_str = unassign_date.strftime('%Y-%m-%d')

        # จัดเก็บ Assign ID ใน dictionary ตามวันที่
        if (assign_date_str, unassign_date_str) not in assign_ids_by_date:
            assign_ids_by_date[(assign_date_str, unassign_date_str)] = []

        assign_ids_by_date[(assign_date_str, unassign_date_str)].append(assign_id)

        # พิมพ์ค่าของ Assign ID, Assign Date และ Unassign Date
        print(f"Assign ID: {assign_id}, Assign Date: {assign_date_str}, Unassign Date: {unassign_date_str}")
        


    # แสดงผล
    for (assign_date_str, unassign_date_str), assign_ids in assign_ids_by_date.items():
        assign_date = pd.to_datetime(assign_date_str)
        unassign_date = pd.to_datetime(unassign_date_str)

        # แยกวัน เดือน ปี
        assign_day = assign_date.day
        assign_day_str = str(assign_day)
        assign_month = assign_date.month
        assign_month_str = str(assign_month)
        assign_year = assign_date.year
        assign_year_str = str(assign_year)

        unassign_day = unassign_date.day
        unassign_day_str = str(unassign_day)
        unassign_month = unassign_date.month
        unassign_month_str = str(unassign_month)
        unassign_year = unassign_date.year
        unassign_year_str = str(unassign_year)
        
        #login
        driver = login()
        time.sleep(2)
        
        # select type data
        select_type(driver)

        time.sleep(2)
        #select Assign Date
        print("กำลังเลือกวันที่...")
        select_date(driver, assign_month_str, assign_year_str, assign_day_str, "ContentPlaceHolder1_img1")
        print("เลือกวันที่เสร็จสิ้น")


        time.sleep(2)
        # Select Unassign Date
        print("กำลังเลือกวันที่สิ้นสุด...")
        select_date(driver, unassign_month_str, unassign_year_str, unassign_day_str, "ContentPlaceHolder1_img2")
        print("เลือกวันที่สิ้นสุดเสร็จสิ้น")
        
        time.sleep(1)
        type_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_btnAssignID"))
        )
        type_button.click()
        time.sleep(2)
        
        print("Assign IDs:")
        for assign_id in assign_ids:
            print(f"- {assign_id}")
            time.sleep(2)
            select_checkbox_by_id(driver, assign_id)
        ok_button = driver.find_element(By.ID, "ContentPlaceHolder1_btnOk")
        ok_button.click()
        select_date(driver, as_of_month_str, as_of_year_str, as_of_day_str, "ContentPlaceHolder1_img3")
        time.sleep(5)
        select_date(driver, as_of_month_str, as_of_year_str, as_of_day_str, "ContentPlaceHolder1_imgRunDate")
        save_button = driver.find_element(By.ID, "ContentPlaceHolder1_btnSave")
        save_button.click()
        for i in range(30):
            try:
                alert = driver.switch_to.alert
                alert_text = alert.text
                alert.accept()
                time.sleep(10)
                x = alert.text
                print(x)
                id_number = x.split(":")[-1].strip()
                print(id_number)
                id_list.append(id_number)
                if alert_text != 'Do you want to set up criteria report?':
                    raise
                break
            except NoAlertPresentException:
                print("ไม่มี alert ปรากฏ")
        driver.quit()
        
    print(id_list)
    
    file_path = r"C:\test_put_file"
    new_path = r"C:\Users\chayakor\Desktop\DMR_BOT\BC"
    assign_id = ['671000155', '671000156', '671000157']
    
    missing_files = assign_id.copy()

    while missing_files:
        for id in missing_files:
            found = False
            for file_name in os.listdir(file_path):
                if id in file_name:
                    print(f"File found: {file_name}")
                    
                    # เปลี่ยนชื่อไฟล์ใหม่
                    new_file_name = f"TEMP_{id}{os.path.splitext(file_name)[1]}"
                    
                    # Path เดิมและใหม่ของไฟล์
                    old_file = os.path.join(file_path, file_name)
                    new_file = os.path.join(new_path, new_file_name)
                    
                    # copy จาก path หลักและเปลี่ยนชื่อไฟล์
                    shutil.copy(old_file, new_file)  # คัดลอกไฟล์แทนการย้าย
                    print(f"Copied file to: {new_file}")
                    
                    # ถ้าเจอไฟล์แล้วออกจากลูป while
                    missing_files.remove(id)
                    found = True
                    break
                
                if found:
                    break  # ออกจากลูปเพื่อตรวจสอบ ID ถัดไป
                
        if missing_files:  # ถ้ายังมี ID ที่ไม่เจอไฟล์
            print("Waiting for missing files...")
            time.sleep(5)

def select_type(driver):
    menu_report = driver.find_element(By.XPATH, "//a[contains(text(), 'Report')]")
    actions = ActionChains(driver)
    actions.move_to_element(menu_report).perform()

    # Step 2: เลื่อนเมาส์ไปที่เมนูย่อย "Residential"
    menu_residential = driver.find_element(By.XPATH, "//a[contains(text(), 'Residential')]")
    actions.move_to_element(menu_residential).perform()
    #คลิก
    driver.find_element(By.XPATH, "//a[contains(text(), 'DMRTH004 - Summary by Call Status')]").click()
    time.sleep(1)

    label = driver.find_element(By.XPATH, "//label[text() = 'Set Batch Criteria']")
    input_radio = label.find_element(By.XPATH, "preceding-sibling::input")
    
    if not input_radio.is_selected():
        input_radio.click()
    
    # เลือก type PRE-DD
    time.sleep(2)
    type_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_btnAssignType"))
        )
    type_button.click()
    time.sleep(2)
    row_type = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//table[@id='ContentPlaceHolder1_gvResult7']//tr[td/span[text()='PRE-DD']]"))
    )

    checkbox = row_type.find_element(By.XPATH, ".//input[@type='checkbox']")
    driver.execute_script("arguments[0].click();", checkbox)
    ok_button = driver.find_element(By.ID, "ContentPlaceHolder1_btnOk")
    ok_button.click()
    
def select_date(driver, month, year, day, calendar_button_id):
    
    attempts = 5
    while attempts > 0:
        try:
            # รอให้ปุ่ม Calendar พร้อมคลิกได้
            calendar_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, calendar_button_id))
            )
            calendar_button.click()  # คลิกที่ปุ่ม Calendar
            break  # หากคลิกสำเร็จ ให้ออกจากลูป
        except StaleElementReferenceException:
            attempts -= 1  # ลดจำนวน retry ลง
            if attempts == 0:
                raise  # ถ้าหมด retry แล้ว ยัง error อยู่ ให้ raise error
    
    time.sleep(1)
    iframes = driver.find_elements(By.TAG_NAME, "iframe")
    for iframe in iframes:
        driver.switch_to.frame(iframe)
        try:
            select_month = driver.find_element(By.XPATH, "//table[@id = 'outerTable']//select[@id = 'MonSelect']")
            select_mdropdown = Select(select_month)
            select_mdropdown.select_by_value(month)
            print("Select mount success")
            select_year = driver.find_element(By.XPATH, "//table[@id = 'outerTable']//select[@id = 'YearSelect']")
            select_ydropdown = Select(select_year)
            select_ydropdown.select_by_value(year)
            print("Select year success")
            date_element = driver.find_element(By.XPATH, f"//a[text()='{day}']")
            date_element.click()
            print("Select day success")
            driver.switch_to.default_content()
            break
        except Exception as e:
            driver.switch_to.default_content()  # กลับไปที่ document หลัก
            print(f"ไม่สามารถเลือกวันที่: {e}")
    
def select_checkbox_by_id(driver, assign_id):
    try:
        # รอให้แถวที่มีข้อความที่ตรงกันปรากฏ
        row_type = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, f"//table[@id='ContentPlaceHolder1_gvResult8']//tr[td/span[text()='{assign_id}']]"))
        )

        # ค้นหา checkbox ภายในแถวที่พบ
        checkbox = row_type.find_element(By.XPATH, ".//input[@type='checkbox']")

        # ใช้ JavaScript เพื่อคลิก checkbox
        driver.execute_script("arguments[0].click();", checkbox)
        print(f"คลิก checkbox ที่ตรงกับ {assign_id} สำเร็จ!")
    
    except Exception as e:
        print(f"เกิดข้อผิดพลาด: {e}")
    
    
if __name__ == '__main__':
    # current_date = datetime.now()
    current_date = datetime(2024, 8, 7)
    get_data(current_date)