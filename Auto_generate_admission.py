#!/usr/bin/env python
# coding: utf-8

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
import pathlib
import chromedriver_autoinstaller
import subprocess
# 配置 WebDriver
chrome_options = Options()
chrome_options.headless = True
chrome_options.add_argument("--headless=new")  # 如果不需要顯示瀏覽器界面，可以啟用 headless 模式
# chrome_options.add_argument("--window-position=-2400,-2400")


chromedriver_autoinstaller.install()
service = Service(service_args=['--log-level=OFF'], log_output=subprocess.STDOUT)

download_path=str(pathlib.Path(__file__).parent.resolve())
chrome_options.add_experimental_option('prefs', {
"download.default_directory": download_path, #Change default directory for downloads
"download.prompt_for_download": False, #To auto download the file
"download.directory_upgrade": True,
"plugins.always_open_pdf_externally": True #It will not show PDF directly in chrome
})
chrome_options.add_argument('--log-level=3')
driver = webdriver.Chrome(service=service,options=chrome_options)

print("此程式為運用chatgpt利用ER 或是OPD note產生病歷，作者的燈號為: 8375K，如果有任何問題或建議，歡迎聯絡!!!")
print("請稍帶片刻...")
time.sleep(12)
username=input("帳號 : ")
password = pwinput.pwinput(prompt='密碼: ', mask='*')
api_key = input("api_key:")
login_url = 'https://eip.vghtpe.gov.tw/login.php'  #

OPD_or_ER= input("需要用門診紀錄請打OPD, 需要ER 請打ER:")
driver.get(login_url)

username_field = driver.find_element(By.ID, 'login_name')  # 替換成實際的字段名稱
password_field = driver.find_element(By.ID, 'password')  # 替換成實際的字段名稱

username_field.send_keys(username)  # 替換成實際的用戶名
password_field.send_keys(password)  # 替換成實際的密碼

# 提交表單
password_field.send_keys(Keys.RETURN)

time.sleep(0.5)

driver.get("https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm?action=findEmr&histno=50687768")
soup = BeautifulSoup(driver.page_source, 'html.parser')




patID=input("病歷號:")

admin_intro=get_admin_Intro(driver,patID)
VS=str(admin_intro.at[0, "主治醫師"])
VS=VS.split("(")[0]
print(VS)


prompt_text=""

with open("Lib/admission prompt.txt", 'r',encoding="utf-8") as f:
	# breakpoint()
	prompt_text=prompt_text+f.read()

try:
	Age=str(admin_intro.at[0, "生　日　"])
	Age=Age.split("（")[1]
	Age=Age.split("）")[0]
	Sex=str(admin_intro.at[0, "性　別　"])
	prompt_text=prompt_text+"Age:"+Age+"Sex"+Sex
except:
	pass



if OPD_or_ER == "OPD":
	try:
		prompt_text=prompt_text+"OPD note\n"
		OPD_note=get_OPD(driver, patID, VS)
		prompt_text=prompt_text+OPD_note+"\n"
	except:
		pass
else:
	try:
		prompt_text=prompt_text+"ER note\n"
		ER_note=get_ER(driver, patID)
		prompt_text=prompt_text+ER_note+"\n"
		prompt_text=prompt_text+"-----------------------------------------------------------------------------------\n"
	except:
		pass
	

prompt_text=prompt_text+"-----------------------------------------------------------------------------------\n"
try:
	nurse_note=get_nurse_note(driver, patID)
	prompt_text=prompt_text+"護理紀錄\n"+nurse_note+"\n"
	prompt_text=prompt_text+"\n-----------------------------------------------------------------------------------\n"
except:
	pass

try:
	dis_note=get_last_discharge(driver,patID)
	prompt_text=prompt_text+"The patient's last discharged note\n"+dis_note+"\n"
	prompt_text=prompt_text+"\n-----------------------------------------------------------------------------------\n"
except:
	pass

# try:
# 	nurse_note=get_nurse_note(driver, patID)
# 	prompt_text=prompt_text+nurse_note
# 	prompt_text=prompt_text+"\n-----------------------------------------------------------------------------------\n"
# except:
# 	pass

time.sleep(3*random.random())
report_num=20
report_name,recent_report=get_recent_report(driver, patID, report_num=report_num)
for i in range(len(report_name)):
	try:
		prompt_text=prompt_text+report_name[i]+"\n"
		prompt_text=prompt_text+recent_report[report_name[i]].to_string()
		prompt_text=prompt_text+"\n-----------------------------------------------------------------------------------\n"
	except:
		pass
# breakpoint()
# print(prompt_text)

path="prompt.txt"
with open(path, 'w',encoding="utf-8") as f:
	f.write(prompt_text)

print("請稍等，chatGPT 正在產生病歷當中...")
from openai import OpenAI

client = OpenAI(api_key = api_key)


completion = client.chat.completions.create(
    model="gpt-4o-mini",
    messages=[
        {"role": "system", 
        "content": "You are a resident doctor, who needs to write admission note based on ER note or OPD note."
        },
        {
            "role": "user",
            "content": prompt_text
        }
    ]
)
replied_text=completion.choices[0].message.content

print(replied_text)

# breakpoint()
path="Replied.html"
with open(path, 'w',encoding="utf-8") as f:
	f.write(replied_text)
print("已產生好病歷，並且儲存為Replied.html，正在回填當中")

soup = BeautifulSoup(replied_text, 'html.parser')

try:
	driver.get("https://web9.vghtpe.gov.tw/emr2/adminote/Admission.do?adistno="+patID+"&last=N&adicase=&action=add")
	time.sleep(3)
	try:
		Chief_complain = driver.find_element(By.ID, 'item02')
		Chief_complain.send_keys(soup.find('div', id="Cheif_complain").text) 
	except:
		pass

	Transfer_hospital = driver.find_element(By.ID, 'aditran')
	Transfer_hospital.send_keys("N/A") 

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

	try:
		Plan = driver.find_element(By.ID, 'item13')
		Plan.send_keys(soup.find('div', id="PLAN").text) 
	except:
		try:
			Plan = driver.find_element(By.ID, 'item13')
			Plan.send_keys(soup.find('div', id="Plan").text)
		except:
			pass

	# Review of system

	for i in range(17):
		ROS_AI_id="ROS_"+str(i+1)
		if i<9:
			ROS_id='item100'+str(i+1)
		else:
			ROS_id='item10'+str(i+1)
		ROS = driver.find_element(By.ID, ROS_id)
		ROS.send_keys(soup.find('div', id=ROS_AI_id).text) 

	## PE
	for i in range(17):
		PE_AI_id="PE_"+str(i+1)
		if i<9:
			PE_id='item110'+str(i+1)
		else:
			PE_id='item11'+str(i+1)
		PE = driver.find_element(By.ID, PE_id)
		PE.send_keys(soup.find('div', id=PE_AI_id).text) 




	save_button = driver.find_element(By.NAME, 'save')
	save_button.click()

	WebDriverWait(driver, 10).until(EC.alert_is_present())
	confirm = driver.switch_to.alert

	# 獲取 confirm 的文本（可選）
	# print(confirm.text)

	# 接受 confirm
	confirm.accept()


	time.sleep(10)

	input("已完成回填，按Enter 結束程式。")
except:
	input("無法填入病歷系統，請檢察是否已有入院病摘，或是病人尚未入院，按Enter 結束程式。")


driver.quit()