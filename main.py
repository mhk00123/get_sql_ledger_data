import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
import openpyxl
import re
import time
import os

# 為會計科目表進行填色
import fill_asset_account_color

# Global Var
# 會計科目表
asset_account = list()
# 試算表
trial_balance = list()
# 損益表
income_statement = list()
# 資產負債表
balance_sheet = list()

# Selenium Setup
user_ip = '192.168.56.102' ## IP -> 192.168.56.xxx 
opts = Options()
opts.add_argument('--headless')  #不顯示Chrome
opts.add_argument('--disable-gpu')
webdriver_path = './chromedriver'
chrome = webdriver.Chrome(executable_path = webdriver_path, chrome_options = opts)

# ------------------- 1 ----------------------- #
## Login SQL-Ledger (user, 6263) 
def login_sql_ledger(user_id, user_password):
    url = 'http://{}/sql-ledger/login.pl'.format(user_ip)
    chrome.get(url)

    login_fill_id = chrome.find_element_by_xpath("/html/body/center/table/tbody/tr/td/form/table/tbody/tr/td/table/tbody/tr[1]/td/input")
    password_fill_id = chrome.find_element_by_xpath("/html/body/center/table/tbody/tr/td/form/table/tbody/tr/td/table/tbody/tr[2]/td/input")

    login_fill_id.send_keys(user_id)
    password_fill_id.send_keys(user_password)

    chrome.find_element_by_xpath('/html/body/center/table/tbody/tr/td/form/table/tbody/tr/td/input').click()
    
# ------------------- 2 ----------------------- #
## 爬取會計科目表
## Return DataFrame
def get_asset_account(user_ip):
    book_nanme = "會計科目表"
    url = 'http://{}/sql-ledger/ca.pl?path=bin/mozilla&action=chart_of_accounts&level=Reports--Chart%20of%20Accounts&login=user&js=1'.format(user_ip)
    chrome.get(url)
    time.sleep(1)
    
    soup = BeautifulSoup(chrome.page_source, 'lxml')
    th_title = list()
    rows = list()
    
    th_title = soup.find_all('tr', {'class':'listheading'})
    th_title = th_title[0].text.split('\n')
    th_title = th_title[0:len(th_title)-1]
    
    regex = re.compile('l*')
    find_tr = soup.find_all('tr', {'class' : regex})
    for i in find_tr:
        split_temp = i.get_text().split('\n')
        for index, element in enumerate(split_temp):
            split_temp[index] = element.replace('\xa0','')
        rows.append(split_temp[0:len(split_temp)-1])
    
    rows = rows[1:len(rows)]
    df_web_data = pd.DataFrame(rows, columns = th_title)

    return df_web_data
    
# ------------------- 3 ----------------------- #
## 爬取試算表
## Return DataFrame
def get_trial_balance(user_ip):
    url = 'http://{}/sql-ledger/rp.pl?path=bin/mozilla&action=report&level=Reports--Trial%20Balance&login=user&js=1&report=trial_balance'.format(user_ip)
    chrome.get(url)
    time.sleep(1)
    
    chrome.find_element_by_xpath('/html/body/form/table/tbody/tr[4]/td/table/tbody/tr[1]/td/input[2]').click()
    chrome.find_element_by_xpath('/html/body/form/input[1]').click()
    time.sleep(1)  
    # print(chrome.page_source)
    soup = BeautifulSoup(chrome.page_source, 'lxml')
    
    th_title = list()
    find_th = soup.find_all("tr")[3].find_all("th", attrs={'class' : 'listheading'})
    for i in find_th:
        th_title.append(i.text)
    
    regex = re.compile('list*')
    rows = list()
    find_tr = soup.find_all('tr')[3].find('table').find_all('tr',{'class':regex})
    for i in find_tr:
        split_temp = i.get_text().split('\n')
        for index, element in enumerate(split_temp):
            split_temp[index] = element.replace('\xa0','')
        rows.append(split_temp[1:len(split_temp)-1])
        
    df_web_data = pd.DataFrame(rows, columns = th_title)
    
    return df_web_data 

# ------------------- 4 ----------------------- #
## 爬取損益表
## Return DataFrame
def get_income_statement(user_ip):
    url = 'http://{}/sql-ledger/rp.pl?path=bin/mozilla&action=report&level=Reports--Income%20Statement&login=user&js=1&report=income_statement'.format(user_ip)
    chrome.get(url)
    time.sleep(0.5)
    
    chrome.find_element_by_xpath('/html/body/form/input[1]').click()
    time.sleep(0.5)
    soup = BeautifulSoup(chrome.page_source, 'lxml')
    print(soup)
    
    rows = list()
    find_tr = soup.find_all('tr')
    for i in find_tr:
        split_temp = i.get_text().split('\n')
        rows.append(split_temp[1:len(split_temp)-1])
        # print(split_temp)
        # print("--------------------------")
        # for i in rows:
            # print(i)
    rows[-1].insert(1,"")
    print(rows[-1])
    
    df_web_data = pd.DataFrame(rows)
    
    return df_web_data

# 寫入Excel Function
# data should be a list
def write_to_excel(file_name_w, sheet_name_w, data_w):
    file_name = '{}.xlsx'.format(file_name_w)
    path_w = os.path.join(os.getcwd(), file_name)
    
    if file_name_w == "損益表":
        header_select = False
    else:
        header_select = True
    
    with pd.ExcelWriter(engine='openpyxl', path=path_w, mode='w') as writer:
        book = writer.book
        try:    
            book.remove(book[sheet_name_w])
        except:
            pass
        data_w.to_excel(writer, sheet_name = sheet_name_w, index = False, header = header_select)
    
    print("Write {} successfully.".format(sheet_name_w))

if __name__ == "__main__":
    login_sql_ledger("user", "6263")
    
    income_statement = get_income_statement(user_ip)
    print(income_statement)
    write_to_excel("損益表", "損益表", income_statement)
    
    # asset_account = get_asset_account(user_ip)
    # trial_balance = get_trial_balance(user_ip)
    
    # print(asset_account)
    # print(trial_balance)
    
    # write_to_excel('試算表', '試算表', trial_balance)
    # write_to_excel('會計科目表', '會計科目表', asset_account)
    
    # fill_asset_account_color.fill_trial_balance_color()
    # fill_asset_account_color.fill_asset_account_color()