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
# 配置 WebDriver
chrome_options = Options()
chrome_options.headless = True
chrome_options.add_argument("--headless=new")  # 如果不需要顯示瀏覽器界面，可以啟用 headless 模式
chrome_options.add_argument("--window-position=-2400,-2400")
chrome_options.add_argument('--log-level=3')

service = Service(executable_path=r'Lib/chromedriver.exe')

download_path=str(pathlib.Path(__file__).parent.resolve())
chrome_options.add_experimental_option('prefs', {
"download.default_directory": download_path, #Change default directory for downloads
"download.prompt_for_download": False, #To auto download the file
"download.directory_upgrade": True,
"plugins.always_open_pdf_externally": True #It will not show PDF directly in chrome
})

driver = webdriver.Chrome(service=service,options=chrome_options)

username=input("帳號 : ")
password = pwinput.pwinput(prompt='密碼: ', mask='*')
api_key = input("api_key:")
login_url = 'https://eip.vghtpe.gov.tw/login.php'  #
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


try:
	prompt_text=prompt_text+"ER note\n"
	ER_note=get_ER(driver, patID)
	prompt_text=prompt_text+ER_note+"\n"
	prompt_text=prompt_text+"-----------------------------------------------------------------------------------\n"
except:
	prompt_text=prompt_text+"OPD note\n"
	OPD_note=get_OPD(driver, patID, VS)
	prompt_text=prompt_text+OPD_note+"\n"

prompt_text=prompt_text+"-----------------------------------------------------------------------------------\n"

try:
	nurse_note=get_nurse_note(driver, patID)
	prompt_text=prompt_text+"護理紀錄\n"++nurse_note+"\n"
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
report_num=30
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



driver.quit()

# def set_paragraph_spacing(doc, spacing=0):
#     """Set paragraph spacing for all paragraphs in the document."""
#     for paragraph in doc.paragraphs:
#         paragraph.paragraph_format.line_spacing = Pt(spacing)
#         paragraph.paragraph_format.space_before = Pt(spacing)
#         paragraph.paragraph_format.space_after = Pt(spacing)

# def set_font_size(doc, size):
#     for paragraph in doc.paragraphs:
#         for run in paragraph.runs:
#             run.font.size = Pt(size)

# def add_table(doc, df):
#     table = doc.add_table(rows=1, cols=len(df.columns))
    
#     # 設置表頭的字體大小
#     hdr_cells = table.rows[0].cells
#     for i, column_name in enumerate(df.columns):
#         hdr_cells[i].text = str(column_name)
#         for paragraph in hdr_cells[i].paragraphs:
#             for run in paragraph.runs:
#                 run.font.size = Pt(6)
#             paragraph.paragraph_format.line_spacing = Pt(0)
#             paragraph.paragraph_format.space_before = Pt(0)
#             paragraph.paragraph_format.space_after = Pt(0)
        
    
#     # 添加數據行
#     for index, row in df.iterrows():
#         row_cells = table.add_row().cells
#         for i, value in enumerate(row):
#             row_cells[i].text = str(value)
#             for paragraph in row_cells[i].paragraphs:
#                 for run in paragraph.runs:
#                     run.font.size = Pt(6)
#                 paragraph.paragraph_format.line_spacing = Pt(0)
#                 paragraph.paragraph_format.space_before = Pt(0)
#                 paragraph.paragraph_format.space_after = Pt(0)
                
#     for col in table.columns:
#         max_length = max(len(cell.text) for cell in col.cells)
#         col_width = Inches(max_length)
#         for cell in col.cells:
#             cell.width = col_width

# def convert_date(date_str):
#     date_str=date_str[3:8]
#     return date_str

# def add_line(doc):
#     doc.add_paragraph("-------------------------------")

# def convert_drug(data_drug):
#     data_drug=data_drug.split(" ")[:2]
#     data_drug=" ".join(data_drug)
#     return data_drug


# def generate_table_report(driver,doc, ID, row_cells,pat):
#     print(ID)
    
#     info_cell=row_cells[0]
#     paragraph = info_cell.paragraphs[0]
#     paragraph.add_run("\n".join(pat))

#     try:
#         TPR=get_TPR(driver,ID)
#         time.sleep(3*random.random())
#         run=paragraph.add_run("\n")
#         paragraph.add_run("\n".join(list(TPR[["體溫","心跳","呼吸","收縮壓","舒張壓"]].iloc[0])))
        
#     except:
#         pass
    
#     try:
#         run=paragraph.add_run()
#         TPR_img=get_TPR_img(driver,ID)
#         time.sleep(3*random.random())
#         image_path = 'temp_image.png'
#         TPR_img.save(image_path)
#         run.add_picture(image_path, width=Inches(1))  # 插入圖片
#         os.remove(image_path)
#     except:
#         pass

#     try:
#         BW_BL=get_BW_BL(driver,ID, adminID="all")
#         BW_BL=BW_BL[["身高","體重"]]
#         add_table(info_cell, BW_BL.head(2) )
#     except:
#         pass
 
#     assessment_cell=row_cells[1]
#     paragraph = assessment_cell.paragraphs[0]
#     try:
#         progress_note=get_progress_note(driver,ID,num=10)
#         time.sleep(3*random.random())

#         for i in range(len(progress_note)):
#             assessment=progress_note[i]["Assessment"]
#             if len(assessment)>5:
#                 break
#         paragraph.add_run(assessment)
#     except:
#         pass

#     Lab_cells = row_cells[2]

#     try:
#         patIO=get_drainage(driver, ID)
#         add_table(Lab_cells,patIO[["項目","總量"]])
#         # add_line(Lab_cells)
#     except:
#         pass
    

#     try:
#         report_num=3
#         report_name,recent_report=get_recent_report(driver, ID, report_num=report_num)
#         time.sleep(3*random.random())
#         for i in range(report_num):
#             Lab_cells.add_paragraph(report_name[i])
#             # add_table(doc, recent_report[report_name[i]])
#     except:
#         pass


#     try:
#         SMAC=get_res_report(driver,ID,resdtype="SMAC")
#         SMAC["日期"]=SMAC["日期"].apply(convert_date)
#         SMAC=SMAC[["日期","NA","K","BUN","CREA","ALT","BILIT","CRP"]]
#         SMAC = SMAC.loc[~(SMAC[["日期","NA","K","BUN","CREA","ALT","BILIT","CRP"]] == '-').all(axis=1)]
#         time.sleep(3*random.random())
#         add_table(Lab_cells, SMAC.tail(3) )
#     except:
#         pass

#     try:
#         CBC=get_res_report(driver,ID,resdtype="CBC")
#         time.sleep(3*random.random())
#         CBC["日期"]=CBC["日期"].apply(convert_date)
#         CBC=CBC[["日期","WBC","HGB","PLT",'SEG', 'PT', 'APTT']]
#         CBC = CBC.loc[~(CBC[["日期","WBC","HGB","PLT",'SEG', 'PT', 'APTT']] == '-').all(axis=1)]
#         add_table(Lab_cells, CBC.tail(3) )
#         # add_line(Lab_cells)
#     except:
#         pass

    
#     try:

#         def convert_drug(data_drug):
#             data_drug=data_drug.split(" ")[:2]
#             data_drug=" ".join(data_drug)
#             return data_drug
#         def convert_drug_date(data_drug_date):
#             data_drug_date=data_drug_date[5:10]
#             return data_drug_date
#         drug=get_drug(driver,ID)
#         drug["學名"]=drug["學名"].apply(convert_drug)
#         drug["開始日"]=drug["開始日"].apply(convert_drug_date)
#         time.sleep(3*random.random())
#         add_table(Lab_cells, drug[drug["狀態"]=="使用中"][["學名","劑量","途徑","頻次","開始日"] ])
#     except:
#         pass

# doc = Document()



# section = doc.sections[0]
# new_width, new_height = section.page_height, section.page_width
# section.orientation = WD_ORIENT.LANDSCAPE
# section.page_width=new_width
# section.page_height=new_height

# # 設定邊界
# section.top_margin = Pt(30)   # 0.5 inch
# section.bottom_margin = Pt(30) # 0.5 inch
# section.left_margin = Pt(30)   # 0.5 inch
# section.right_margin = Pt(30)  # 0.5 inch

# header = section.header
# paragraph=header.paragraphs[0]
# run = paragraph.add_run("日期:"+datetime.now().strftime('%Y-%m-%d')+" 醫師: "+docID)
# run.font.size = Pt(6)

# table = doc.add_table(rows=1, cols=3)
# table.style = 'Table Grid'

# hdr_cells = table.rows[0].cells
# hdr_cells[0].text = '病人資料'
# hdr_cells[1].text = 'Assessment'
# hdr_cells[2].text = 'Lab Data+drug'
# for cell in hdr_cells:
#     set_font_size(cell, 6)


# for pat in pat_data:
#     row_cells = table.add_row().cells
#     ID=pat[1]
#     generate_table_report(driver=driver,doc=doc, ID=ID, row_cells=row_cells,pat=pat)
#     for cell in row_cells:
#         set_font_size(cell, 6)

# for idx,col in enumerate(table.columns):
#     max_length = max(len(cell.text) for cell in col.cells)
#     col_width = Inches(max_length)
#     if idx==2:
#         col_width = Inches(max_length*0.8)
#     for cell in col.cells:
#         cell.width = col_width


# # 設置所有文本字體為 6 號
# set_font_size(doc, 6)
# set_paragraph_spacing(doc, spacing=0)

# # 保存 Word 文件
# doc.save(docID+'.docx')
# print("儲存為"+docID+'.docx')


