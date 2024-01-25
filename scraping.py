#!/usr/bin/env python
# coding: utf-8

# In[312]:


from selenium import webdriver
from bs4 import BeautifulSoup
from selenium.webdriver.support.select import Select
import time
import pandas as pd
import random
import os
from bs4 import BeautifulSoup
import urllib.request as req
import requests
import re
import datetime

from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import base64
from email.mime.text import MIMEText
import pickle


# In[332]:


# 文字列で”年”,”月”,”日”を取得
dt = datetime.datetime.today()
dt_year = f'{dt.year}'
dt_mon = f'{dt.month}'
dt_day = f'{dt.day}'


# In[314]:


driver = webdriver.Chrome(executable_path="/usr/local/bin/chromedriver")

# Instagram ログインページへ移動
loginUrl = "https://xxxxxxxxxxxxx/"
driver.get(loginUrl)
driver.implicitly_wait(30)


# In[315]:


username = "example@test.co.jp"
password = "hogehogehuga"
# Usernameを入力
usernameInput = driver.find_element_by_id("login_login_id")

usernameInput.send_keys(username)
# Passwordを入力
passwordInput = driver.find_element_by_id("login_password")
passwordInput.send_keys(password)


# In[316]:


# ログインボタンクリック
loginButton = driver.find_element_by_id("SubButton")
loginButton.click()
time.sleep(5)


# In[333]:


# 検索期間ドロップダウンリスト選択　月
dropdown_year = driver.find_element_by_id('s_from_year')
select = Select(dropdown_year)
select.select_by_value(dt_year)

# 検索期間　月
dropdown_mon = driver.find_element_by_id('s_from_month')
select = Select(dropdown_mon)
select.select_by_value(dt_mon)

# 検索期間　日
dropdown_day = driver.find_element_by_id('s_from_day')
select = Select(dropdown_day)
select.select_by_value(dt_day)

# 検索期間　年
dropdown_t_year = driver.find_element_by_id('s_to_year')
select = Select(dropdown_t_year)
select.select_by_value(dt_year)

# 検索期間　月
dropdown_t_mon = driver.find_element_by_id('s_to_month')
select = Select(dropdown_t_mon)
select.select_by_value(dt_mon)

# 検索期間　日
dropdown_t_day = driver.find_element_by_id('s_to_day')
select = Select(dropdown_t_day)
select.select_by_value(dt_day)

# 再検索ボタン
loginButton = driver.find_element_by_xpath('/html/body/div[1]/div[3]/div/div/div/div[4]/form/div[2]/input[1]')
loginButton.click()
time.sleep(5)


# In[270]:


# ステータス選択
dropdown_status = driver.find_element_by_id('s_status_choice')
select = Select(dropdown_status)
select.select_by_value('mi')


# In[271]:


# 担当者検索
dropdown_assigned = driver.find_element_by_id('s_assigned_choice')
select = Select(dropdown_assigned)
select.select_by_value('40')


# In[272]:


# 再検索ボタン
loginButton = driver.find_element_by_xpath('/html/body/div[1]/div[3]/div/div/div/div[4]/form/div[8]/input[1]')
loginButton.click()
time.sleep(5)


# In[274]:


# 表示したページのhtmlを取得する
html = driver.page_source


# In[275]:


# html.parser→HTMLを解析する
parse_html = BeautifulSoup(html, 'html.parser')


# In[276]:


# find_all（"タグ名","クラス名"）
atag = parse_html.find_all('a')


# In[277]:


# aタグのテキスト部分のみを取得

tag_list = []


for i in range(len(atag)):
    tag_text = atag[i].text
    tag_list.append(tag_text)
    
tag_list


# In[278]:


# 数値のみのテキストを取得
client_list = []

r = re.compile('^[0-9]+$')
for x in filter(r.match, tag_list):
    client_list.append(x)
client_list


# In[279]:


df = pd.DataFrame(columns=['企業名','部署名', '担当者', '役職', '電話番号', 'メールアドレス'])

for n in range(len(client_list)):
    url = 'https://xxxxxx.xx/xxxx/inquiry/detail/id/{}'.format(client_list[n])
    driver.get(url)
    driver.implicitly_wait(30)
    r = requests.get(url)
    html = driver.page_source
    parse_html = BeautifulSoup(html, 'html.parser')
    client_data = parse_html.find_all('td')[0:6]
    client_data_list = []
    for y in range(len(client_data)+1):
        if y <= 5 :
            client_data_item = client_data[y].text
            client_data_list.append(client_data_item)
        else:
            df_ = pd.Series(client_data_list,index=df.columns)
            df = df.append(df_, ignore_index=True)


# In[280]:


address = "[0-9a-zA-Z_.+-]+@[0-9a-zA-Z-]+\.[0-9a-zA-Z-.]+"
address_pattern = re.compile(address)

for x in range(len(df)):
    email_date = df.iat[x,5].replace('\n','')
    email_ = address_pattern.search(email_date)
    email = email_.group()
    df.iat[x,5] = email


# In[281]:


df.to_excel('output/example.xlsx', index=False) 


# In[282]:


# date_list = df.iloc[0].to_list()


# In[283]:


# from google_auth_oauthlib.flow import InstalledAppFlow
# from googleapiclient.discovery import build

# import base64
# from email.mime.text import MIMEText
# import pickle


# In[284]:


SCOPES = ["https://www.googleapis.com/auth/gmail.compose"]

def get_credential():
    if os.path.exists('token.pickle'):
        # 再利用（pickle読み取り）！
        with open('token.pickle', 'rb') as token:
                cred = pickle.load(token)
    else:
        launch_browser = True
        flow = InstalledAppFlow.from_client_secrets_file("hogehogehuga.apps.googleusercontent.com.json", SCOPES)
        flow.run_local_server()
        cred = flow.credentials
        with open('token.pickle', 'wb') as token:
            pickle.dump(cred, token)
    return cred

def create_message(sender, to, subject, message_text):
    enc = "utf-8"
    message = MIMEText(message_text.encode(enc), _charset=enc)
    message["to"] = to
    message["from"] = sender
    message["subject"] = subject
    encode_message = base64.urlsafe_b64encode(message.as_bytes())
    return {"raw": encode_message.decode()}

def create_draft(service, user_id, message_body):
    message = {'message':message_body}
    draft = service.users().drafts().create(userId=user_id, body=message).execute()
    return draft


def main(sender, to, subject, message_text):
    creds = get_credential()
    service = build("gmail", "v1", credentials=creds, cache_discovery=False)
    message = create_message(sender, to, subject, message_text)
    create_draft(service, "me", message)

for t in range(len(df)):
    date_list = df.iloc[t].to_list()
    if __name__ == "__main__":

        sender = 'example@test.co.jp'
        to = date_list[5]
        subject = '自動送信テスト'
        message_text = """{0}
        {1}　様

        お世話になります。
        xxxxx
        """.format(date_list[0],date_list[2])


        main(sender=sender, to=to, subject=subject, message_text=message_text)


# In[301]:





# In[ ]:




