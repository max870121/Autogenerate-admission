from fastapi import FastAPI, Form, BackgroundTasks, HTTPException
from fastapi.responses import HTMLResponse
from fastapi.templating import Jinja2Templates
from fastapi import Request
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import random
import pathlib
import subprocess
import chromedriver_autoinstaller
from openai import OpenAI
from typing import Optional
from bs4 import BeautifulSoup
from admission_function import *
from fastapi.responses import RedirectResponse
# Initialize FastAPI
app = FastAPI()

# Initialize Jinja2 Templates
templates = Jinja2Templates(directory="templates")
task_status={}

# Install ChromeDriver automatically if needed
chromedriver_autoinstaller.install()

# Setup ChromeDriver options
chrome_options = Options()
chrome_options.headless = True
chrome_options.add_argument("--headless=new")  # headless mode
download_path = str(pathlib.Path(__file__).parent.resolve())
chrome_options.add_experimental_option('prefs', {
    "download.default_directory": download_path,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "plugins.always_open_pdf_externally": True
})

# Define background task to handle the web scraping and report generation
def process_medical_report(username: str, password: str, api_key: str, patient_id: str, OPD_or_ER: str):
    service = webdriver.chrome.service.Service(service_args=['--log-level=OFF'], log_output=subprocess.STDOUT)
    driver = webdriver.Chrome(service=service, options=chrome_options)

    try:
        # Web login to the hospital system
        login_url = 'https://eip.vghtpe.gov.tw/login.php'
        driver.get(login_url)

        username_field = driver.find_element(By.ID, 'login_name')
        password_field = driver.find_element(By.ID, 'password')
        username_field.send_keys(username)
        password_field.send_keys(password)
        password_field.send_keys(Keys.RETURN)
        time.sleep(0.5)

        # Patient-specific details page
        driver.get("https://web9.vghtpe.gov.tw/emr/qemr/qemr.cfm?action=findEmr&histno=50687768")
        soup = BeautifulSoup(driver.page_source, 'html.parser')

        # Handle patient-related data
        admin_intro=get_admin_Intro(driver,patient_id)
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
            try:
                prompt_text=prompt_text+"OPD note\n"
                OPD_note=get_OPD(driver, patient_id, VS)
                prompt_text=prompt_text+OPD_note+"\n"
            except:
                pass
        else:
            try:
                prompt_text=prompt_text+"ER note\n"
                ER_note=get_ER(driver, patient_id)
                prompt_text=prompt_text+ER_note+"\n"
                prompt_text=prompt_text+"-----------------------------------------------------------------------------------\n"
            except:
                pass
            

        prompt_text=prompt_text+"-----------------------------------------------------------------------------------\n"
        try:
            nurse_note=get_nurse_note(driver, patient_id)
            prompt_text=prompt_text+"護理紀錄\n"+nurse_note+"\n"
            prompt_text=prompt_text+"\n-----------------------------------------------------------------------------------\n"
        except:
            pass

        try:
            dis_note=get_last_discharge(driver,patient_id)
            prompt_text=prompt_text+"The patient's last discharged note\n"+dis_note+"\n"
            prompt_text=prompt_text+"\n-----------------------------------------------------------------------------------\n"
        except:
            pass


        time.sleep(3*random.random())
        report_num=20
        report_name,recent_report=get_recent_report(driver, patient_id, report_num=report_num)
        for i in range(len(report_name)):
            try:
                prompt_text=prompt_text+report_name[i]+"\n"
                prompt_text=prompt_text+recent_report[report_name[i]].to_string()
                prompt_text=prompt_text+"\n-----------------------------------------------------------------------------------\n"
            except:
                pass
        print("complete getting data")

        # Send data to OpenAI API for report generation
        client = OpenAI(api_key=api_key)
        completion = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{
                "role": "system", 
                "content": "You are a resident doctor, who needs to write admission notes based on ER or OPD notes."
            }, {
                "role": "user", 
                "content": prompt_text
            }]
        )
        replied_text = completion.choices[0].message.content

        # Save and process the reply
        path = "Replied.html"
        with open(path, 'w', encoding="utf-8") as f:
            f.write(replied_text)

        # Further actions for saving the report or updating the medical system would go here
        soup = BeautifulSoup(replied_text, 'html.parser')

        driver.get("https://web9.vghtpe.gov.tw/emr2/adminote/Admission.do?adistno="+patient_id+"&last=N&adicase=&action=add")
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

        print("已完成回填，按Enter 結束程式。")

    except Exception as e:
        print(f"Error processing the medical report: {e}")
        raise HTTPException(status_code=500, detail="Error during report generation")
    finally:
        driver.quit()

# Define an endpoint to start the background task
@app.post("/generate_report/")
async def generate_report(
    username: str = Form(...),
    password: str = Form(...),
    api_key: str = Form(...),
    patient_id: str = Form(...),
    OPD_or_ER: str = Form(...),
    background_tasks: BackgroundTasks = BackgroundTasks()
):


    # Start the background task to process the report
    background_tasks.add_task(process_medical_report, username, password, api_key, patient_id, OPD_or_ER)
    
    # Set task status to "in progress"

    # Return a response with a task_id, which the frontend can check to know when the task is complete
    return RedirectResponse(url="/", status_code=303)




# Serve the HTML form from a separate template file
@app.get("/", response_class=HTMLResponse)
async def read_root(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

