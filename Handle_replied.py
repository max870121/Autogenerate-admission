from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import pandas as pd
from bs4 import BeautifulSoup
import time
import random

import pandas as pd
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.section import WD_ORIENT
from PIL import Image
from docx.oxml.ns import qn
import os
from admission_function import *
from datetime import datetime, timedelta
import pwinput
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
replied_text=""
with open("Replied.html", 'r',encoding="utf-8") as f:
	replied_text=replied_text+f.read()
soup = BeautifulSoup(replied_text, 'html.parser')
# print(soup)
print(soup.find('div', id="Chief_complaint"))



chrome_options = Options()
chrome_options.headless = True
# chrome_options.add_argument("--headless=new")  # 如果不需要顯示瀏覽器界面，可以啟用 headless 模式
# chrome_options.add_argument("--window-position=-2400,-2400")
chrome_options.add_argument('--log-level=3')

service = Service(executable_path=r'chromedriver.exe')
driver = webdriver.Chrome(service=service,options=chrome_options)

username=input("帳號 : ")

password = pwinput.pwinput(prompt='密碼: ', mask='*')
ID=input("病歷號 : ")
login_url = 'https://eip.vghtpe.gov.tw/login.php'  #
driver.get(login_url)

username_field = driver.find_element(By.ID, 'login_name')  # 替換成實際的字段名稱
password_field = driver.find_element(By.ID, 'password')  # 替換成實際的字段名稱

username_field.send_keys(username)  # 替換成實際的用戶名
password_field.send_keys(password)  # 替換成實際的密碼

# 提交表單
password_field.send_keys(Keys.RETURN)
time.sleep(5)


driver.get("https://web9.vghtpe.gov.tw/emr2/adminote/Admission.do?adistno="+ID+"&last=N&adicase=&action=add")
time.sleep(3)
Chief_complain = driver.find_element(By.ID, 'item02')
Chief_complain.send_keys(soup.find('div', id="Chief_complaint").text) 

PRESENT_ILLNESS = driver.find_element(By.ID, 'item03')
PRESENT_ILLNESS.send_keys(soup.find('div', id="PRESENT_ILLNESS").text) 

PAST_HISTORY = driver.find_element(By.ID, 'item04')
PAST_HISTORY.send_keys(soup.find('div', id="PAST_HISTORY").text) 

PERSONAL_HISTORY = driver.find_element(By.ID, 'item05')
PERSONAL_HISTORY.send_keys(soup.find('div', id="PERSONAL_HISTORY").text) 

FAMILY_HISTORY = driver.find_element(By.ID, 'item06')
FAMILY_HISTORY.send_keys(soup.find('div', id="FAMILY_HISTORY").text) 

IMPRESSION = driver.find_element(By.ID, 'item12')
IMPRESSION.send_keys(soup.find('div', id="IMPRESSION").text) 

Plan = driver.find_element(By.ID, 'item13')
Plan.send_keys(soup.find('div', id="PLAN").text) 

save_button = driver.find_element(By.NAME, 'save')
save_button.click()

WebDriverWait(driver, 10).until(EC.alert_is_present())
confirm = driver.switch_to.alert

# 獲取 confirm 的文本（可選）
print(confirm.text)

# 接受 confirm
confirm.accept()
breakpoint()