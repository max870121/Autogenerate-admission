from bs4 import BeautifulSoup
from VGH_function import *
import requests
import time

def submit_medical_form_with_requests(patient_id, replied_text, vgh, task_id=None, task_status_dict=None):
    """
    使用 requests 版本提交醫療表單
    
    Args:
        patient_id: 病人ID
        replied_text: 回覆的HTML文本
        session: requests.Session 物件 (如果已經登入)
        task_id: 任務ID
        task_status_dict: 任務狀態字典
    """
    
    # 解析回覆文本
    
    soup = BeautifulSoup(replied_text, 'html.parser')
    # 如果沒有提供 session，建立新的 session
    
    # try:
        
    # 尋找所有隱藏欄位 (包括 CSRF token)
    form_url="https://web9.vghtpe.gov.tw/emr2/adminote/Admission.do?adistno="+patient_id+"&last=N&adicase=&action=add"
    page_content = vgh.get_page_after_login(form_url)
    form_soup = BeautifulSoup(page_content, 'html.parser')
    form_data ={
        "adistno": "",
        "adiname": "",
        "adidate": "",
        "aditime": "",
        "adicase": "",
        "adisect": "",
        "adiward": "",
        "adibed": "",
        "adifinal": "",
        "adisex": "",
        "adibirth": "",
        "adiaskdt": "",
        "adiasktm": "",
        "adioccup2": "",
        "adioccup": "",
        "adiblood": "",
        "aditran": "",
        "adimarry": "",
        "adimarry2": "",
        "nyha": "",
        "adippd": "",
        "adicyear": "",
        "adisyear": "",
        "item02": "",
        "item03": "",
        "item04": "",
        "item05": "",
        "item15": "",
        "item06": "",
        "item07": "",
        "item08": "",
        "item09": "",
        "item1001": "",
        "item1002": "",
        "item1003": "",
        "item1004": "",
        "item1005": "",
        "item1006": "",
        "item1007": "",
        "item1008": "",
        "item1009": "",
        "item1010": "",
        "item1011": "",
        "item1012": "",
        "item1013": "",
        "item1014": "",
        "item1015": "",
        "item1016": "",
        "item1017": "",
        "aditall": "",
        "adiweigh": "",
        "adibmi": "",
        "adiwaist": "",
        "adibt": "",
        "adipr": "",
        "adirr": "",
        "adibph": "",
        "adibpl": "",
        "item1101": "",
        "item1102": "",
        "item1103": "",
        "item1104": "",
        "item1105": "",
        "item1106": "",
        "item1107": "",
        "item1108": "",
        "item1109": "",
        "item1110": "",
        "item1111": "",
        "item1112": "",
        "item1113": "",
        "item1114": "",
        "item1115": "",
        "item1116": "",
        "item1117": "",
        "adinse8": "0",
        "adinse4": "",
        "item14": "",
        "item12": "",
        "item13": "",
        "adiplan.diagname": "",
        "adiplan.person": "",
        "adiplan.language": "",
        "adiplan.othrmemo": "",
        "adivs": "",
        "adivsnm": "",
        "adir": "",
        "adirname": "",
        "adir2": "",
        "adir2nam": "",
        "adiassid": "",
        "adiassnm": "",
        "memo": "",
        "drtype": "",
        "doctype": "",
        "adientid": "",
        "adientnm": "",
        "adientdt": "",
        "nseType": "",
        "showHbmi": "",
        "checkHf": "false",
        "checkPregweek": "false",
        "serial": "0",
        "version": "0",
        "status": "0",
        "userId": "",
        "userType": "",
        "smrmark": ""
    }
    inputs = form_soup.find_all('input')
    for input_tag in inputs:
        name = input_tag.get('name')
        value = input_tag.get('value', '')
        if name in form_data.keys():
            form_data[name] = value

    select = form_soup.find_all('select')
    for select_tag in select:
        name=select_tag.get('name')
        selected_option = select_tag.find('option', selected=True)
        if selected_option:
            value = selected_option.get('value')
        if name in form_data.keys():
            
            form_data[name] = value



    # hidden_inputs = form_soup.find_all('input', type='hidden')
    # for hidden_input in hidden_inputs:
    #     name = hidden_input.get('name')
    #     value = hidden_input.get('value', '')
    #     if name:
    #         form_data[name] = value
    
    # 3. 填入醫療資訊
    # Chief complain
    try:
        chief_complain_text = soup.find('div', id="Cheif_complain")
        if chief_complain_text:
            form_data['item02'] = chief_complain_text.text.strip()
    except:
        pass
    
    # Transfer hospital
    form_data['aditran'] = "N/A"
    
    # Present illness
    present_illness_text = soup.find('div', id="PRESENT_ILLNESS")
    if present_illness_text:
        form_data['item03'] = present_illness_text.text.strip()
    
    # Past history
    past_history_text = soup.find('div', id="PAST_HISTORY")
    if past_history_text:
        form_data['item04'] = past_history_text.text.strip()
    
    # Personal history
    personal_history_text = soup.find('div', id="PERSONAL_HISTORY")
    if personal_history_text:
        form_data['item05'] = personal_history_text.text.strip()
    
    # Family history
    family_history_text = soup.find('div', id="FAMILY_HISTORY")
    if family_history_text:
        form_data['item06'] = family_history_text.text.strip()
    
    # Impression
    impression_text = soup.find('div', id="IMPRESSION")
    if impression_text:
        form_data['item12'] = impression_text.text.strip()
    
    # Plan
    try:
        plan_text = soup.find('div', id="PLAN")
        if plan_text:
            form_data['item13'] = plan_text.text.strip()
        else:
            plan_text = soup.find('div', id="Plan")
            if plan_text:
                form_data['item13'] = plan_text.text.strip()
    except:
        pass
    
    # 4. Review of system (ROS)
    for i in range(17):
        ros_ai_id = f"ROS_{i+1}"
        if i < 9:
            ros_id = f'item100{i+1}'
        else:
            ros_id = f'item10{i+1}'
        
        ros_text = soup.find('div', id=ros_ai_id)
        if ros_text:
            form_data[ros_id] = ros_text.text.strip()
    
    # 5. Physical Examination (PE)
    for i in range(17):
        pe_ai_id = f"PE_{i+1}"
        if i < 9:
            pe_id = f'item110{i+1}'
        else:
            pe_id = f'item11{i+1}'
        
        pe_text = soup.find('div', id=pe_ai_id)
        if pe_text:
            form_data[pe_id] = pe_text.text.strip()
    
    # 6. 提交表單
    # 尋找表單的 action URL
    form_element = form_soup.find('form')
    if form_element:
        action_url = form_element.get('action')
        if action_url:
            # 如果是相對路徑，轉換為絕對路徑
            if action_url.startswith('/'):
                submit_url = f"https://web9.vghtpe.gov.tw{action_url}"
            elif action_url.startswith('http'):
                submit_url = action_url
            else:
                submit_url = f"https://web9.vghtpe.gov.tw/emr2/adminote/{action_url}"
        else:
            submit_url = form_url  # 如果沒有 action，使用原始 URL
    else:
        submit_url = form_url
    
    # 設定請求標頭
    headers = {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Referer': form_url,
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }
    
    # 提交表單
    breakpoint()
    form_data['save'] = "暫時存檔"
    submit_response = vgh.session.post("https://web9.vghtpe.gov.tw/emr2/adminote/Admission.do?action=save", data=form_data, headers=headers)
    submit_response.raise_for_status()
    
    # 檢查提交結果
    if submit_response.status_code == 200:
        print("表單提交成功")
        print("已完成回填")
        if task_status_dict and task_id:
            task_status_dict[task_id] = ["Completed", replied_text]
        return True
    else:
        print(f"表單提交失敗，狀態碼: {submit_response.status_code}")
        return False