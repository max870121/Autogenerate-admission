from fastapi import FastAPI, Form, BackgroundTasks, HTTPException
from fastapi.responses import HTMLResponse
from fastapi.templating import Jinja2Templates
from fastapi import Request
import requests
import time
from submit_admission_note import *
import random
import pathlib
import subprocess
# from selenium import webdriver
# from selenium.webdriver.common.keys import Keys
# from selenium.webdriver.common.by import By
# from selenium.webdriver.chrome.service import Service
# from selenium.webdriver.chrome.options import Options
from openai import OpenAI
from typing import Optional
from bs4 import BeautifulSoup
from VGH_function import *
from AI_write import *
from VGH_login import VGHLogin
from fastapi.responses import RedirectResponse
from dotenv import load_dotenv

load_dotenv('.env')
OPENAI_API_KEY = os.getenv('OPENAI_API_KEY')
# Initialize FastAPI
app = FastAPI()
import uuid
templates = Jinja2Templates(directory="templates")
task_status_dict={}

Complete_text=""

# Define background task to handle the web scraping and report generation
def process_medical_report(username: str, password: str, api_key: str, patient_id: str, OPD_or_ER: str,task_id: str):
    vgh = VGHLogin()
        # Web login to the hospital system
    vgh.login(username, password)
    time.sleep(0.5)
    print("login_complete")
    # Patient-specific details page
    page_content = vgh.get_page_after_login("https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm?action=findEmr&histno=50687768")

    # Handle patient-related data
    admin_intro=get_admin_Intro(vgh,patient_id)
    VS=str(admin_intro.at[0, "主治醫師"])
    VS=VS.split("(")[0]
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
        prompt_text=prompt_text+"OPD note\n"
        OPD_note=get_OPD(vgh, patient_id, VS)
        prompt_text=prompt_text+OPD_note+"\n"
    else:
        prompt_text=prompt_text+"ER note\n"
        ER_note=get_ER(vgh, patient_id)
        prompt_text=prompt_text+ER_note+"\n"
        prompt_text=prompt_text+"-----------------------------------------------------------------------------------\n"
        

    prompt_text=prompt_text+"-----------------------------------------------------------------------------------\n"
    nurse_note=get_nurse_note(vgh, patient_id)
    prompt_text=prompt_text+"護理紀錄\n"+nurse_note+"\n"
    prompt_text=prompt_text+"\n-----------------------------------------------------------------------------------\n"

    try:
        pass
        dis_note=get_last_discharge(driver,patient_id)
        prompt_text=prompt_text+"The patient's last discharged note\n"+dis_note+"\n"
        prompt_text=prompt_text+"\n-----------------------------------------------------------------------------------\n"
    except:
        pass


    time.sleep(3*random.random())
    report_num=1
    report_name,recent_report=get_recent_report(vgh, patient_id, report_num=report_num)
    for i in range(len(report_name)):
        try:
            prompt_text=prompt_text+report_name[i]+"\n"
            prompt_text=prompt_text+recent_report[report_name[i]].to_string()
            prompt_text=prompt_text+"\n-----------------------------------------------------------------------------------\n"
        except:
            pass
    print("complete getting data")
    replied_text=Auto_write_admission_ChatGPT(prompt_text, OPENAI_API_KEY)
    print(replied_text)
    submit_medical_form_with_requests(patient_id, replied_text, vgh, task_id=None, task_status_dict=None)
 



@app.post("/generate_report/")
async def generate_report(
    username: str = Form(...),
    password: str = Form(...),
    api_key: str = Form(...),
    patient_id: str = Form(...),
    OPD_or_ER: str = Form(...),
    background_tasks: BackgroundTasks = BackgroundTasks()
):
    task_id = str(uuid.uuid4())
    task_status_dict[task_id] = "In Progress"  # 初始狀態為進行中
    # Start the background task to process the report
    background_tasks.add_task(process_medical_report, username, password, api_key, patient_id, OPD_or_ER, task_id)

    return RedirectResponse(url=f"/task_status/{task_id}", status_code=303)


@app.get("/task_status/{task_id}")
async def task_status(task_id: str):
    # 查詢任務的狀態
    status = task_status_dict.get(task_id)
    if not status:
        raise HTTPException(status_code=404, detail="Task not found.")
    return {"task_id": task_id, "status": status}

@app.get("/Admission_note/{task_id}", response_class=HTMLResponse)
async def admission_note(request: Request, task_id: str):
    # 查詢任務的狀態
    status = task_status_dict.get(task_id)
    if not status:
        raise HTTPException(status_code=404, detail="Task not found.")
    else:
        html=status[1]
        soup = BeautifulSoup(html, 'html.parser')
        sections = []

        for div in soup.find_all('div'):
            section_id = div.get("id")
            if section_id:
                content = div.get_text(separator="\n", strip=True)
                sections.append({"id": section_id, "content": content})
        
    return templates.TemplateResponse("admission_note.html", {"request": request, "sections": sections })
    # return HTMLResponse(content=status[1], status_code=200)


# Serve the HTML form from a separate template file
@app.get("/", response_class=HTMLResponse)
async def read_root(request: Request):
    task_status_dict={}
    return templates.TemplateResponse("index.html", {"request": request})

