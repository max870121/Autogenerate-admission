import requests
from bs4 import BeautifulSoup
import urllib.parse
import pandas as pd
import os
from VGH_function import *
from datetime import datetime, timedelta
import time
import random


class VGHLogin:
    def __init__(self):
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        })
        self.csrf_token = None
        self.base_url = "https://eip.vghtpe.gov.tw/login.php"
    
    def get_login_page(self):
        """取得登入頁面並解析CSRF token"""
        try:
            response = self.session.get(self.base_url)
            response.raise_for_status()
            
            # 解析HTML取得CSRF token
            soup = BeautifulSoup(response.text, 'html.parser')
            csrf_meta = soup.find('meta', {'name': 'csrf-token'})
            if csrf_meta:
                self.csrf_token = csrf_meta.get('content')
                # print(f"取得CSRF Token: {self.csrf_token}")
            
            return True
        except requests.RequestException as e:
            print(f"取得登入頁面失敗: {e}")
            return False
    
    def login(self, username, password):
        """執行登入"""
        if not self.get_login_page():
            return False
        
        # 準備登入資料
        login_data = {
            'login_name': username,
            'password': password,
            'loginCheck': '1',
            'fromAjax': '1'
        }
        
        # 設定headers
        headers = {
            'X-CSRF-TOKEN': self.csrf_token,
            'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
            'X-Requested-With': 'XMLHttpRequest',
            'Referer': self.base_url
        }
        
        try:
            # 發送登入請求
            login_url = urllib.parse.urljoin(self.base_url, '/login_action.php')
            response = self.session.post(
                login_url,
                data=login_data,
                headers=headers
            )
            response.raise_for_status()
            
            # 解析回應
            result = response.json()

            if 'error' in result:
                error_code = int(result['error'])
                if error_code == 0:
                    print("登入成功!")
                    if 'url' in result:
                        print(f"重定向至: {result['url']}")
                        # 可以選擇跟隨重定向
                        dashboard_response = self.session.get("https://eip.vghtpe.gov.tw/"+result['url'])
                        login_url="https://eip.vghtpe.gov.tw/"+dashboard_response.text.split("/")[1][:-2]
                        dashboard_response = self.session.get(login_url)
                        return True
                else:
                    print(f"登入失敗: {result.get('err_msg', '未知錯誤')}")
                    return False
            else:
                print("未預期的回應格式")
                return False
                
        except requests.RequestException as e:
            print(f"登入請求失敗: {e}")
            return False
        except ValueError as e:
            print(f"JSON解析失敗: {e}")
            return False
    
    def get_page_after_login(self, url):
        """登入後取得其他頁面"""
        try:
            response = self.session.get(url)
            response.raise_for_status()
            return response.text
        except requests.RequestException as e:
            print(f"取得頁面失敗: {e}")
            return None

    def get_img_after_login(self, url):
        """登入後取得其他頁面"""
        try:
            response = self.session.get(url)
            response.raise_for_status()
            return response
        except requests.RequestException as e:
            print(f"取得頁面失敗: {e}")
            return None