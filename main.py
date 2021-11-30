import pandas as pd
import time
import os
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import openpyxl
from openpyxl.styles import PatternFill
from bs4 import BeautifulSoup


# 為會計科目表進行填色
import fill_asset_account_color

# 全域變數
user_ip = '192.168.56.102' ## IP -> 192.168.56.xxx 
opts = Options()
# opts.add_argument('--headless')  #不顯示Chrome
opts.add_argument('--disable-gpu')
webdriver_path = './chromedriver'
chrome = webdriver.Chrome(executable_path = webdriver_path, chrome_options = opts)

# ------------------- 1 ----------------------- #
## Login SQL-Ledger (user) 
def login_sql_ledger(user_id, user_password):
    url = 'http://{}/sql-ledger/login.pl'.format(user_ip)
    chrome.get(url)

    login_fill_id = chrome.find_element_by_xpath("/html/body/center/table/tbody/tr/td/form/table/tbody/tr/td/table/tbody/tr[1]/td/input")
    password_fill_id = chrome.find_element_by_xpath("/html/body/center/table/tbody/tr/td/form/table/tbody/tr/td/table/tbody/tr[2]/td/input")

    login_fill_id.send_keys(user_id)
    password_fill_id.send_keys(user_password)

    chrome.find_element_by_xpath('/html/body/center/table/tbody/tr/td/form/table/tbody/tr/td/input').click()
    
    time.sleep(1)
    
# ------------------- 2 ----------------------- #
## 爬取會計科目表
def asset_account(user_ip):
    book_nanme = "會計科目表"
    url = 'http://{}/sql-ledger/ca.pl?path=bin/mozilla&action=chart_of_accounts&level=Reports--Chart%20of%20Accounts&login=user&js=1'.format(user_ip)
    chrome.get(url)
    time.sleep(21)

    web_data_sc = pd.read_html(url, encoding="utf-8")
    df_web_data = web_data_sc[0]
    df_web_data.columns = ["帳戶", "財務訊息通用索引(GIFI)", "說明", "借方", "貸方"]
    # print(web_data)

    print(type(df_web_data))
    print("----------------")
    print(df_web_data)
    
    # 寫入到Excel
    path = os.path.join(os.getcwd(), 'SQL-Ledger.xlsx') # 設定路徑及檔名
    writer = pd.ExcelWriter(path, engine='openpyxl') # 指定引擎openpyxl
    df_web_data.to_excel(writer, sheet_name=book_nanme ,index=False)
    writer.save()
    
    fill_asset_account_color.fill_asset_account_color()

# ------------------- 3 ----------------------- #
## 爬取試算表
def Spreadsheet(user_ip):
    url = 'http://{}/sql-ledger/rp.pl?path=bin/mozilla&action=report&level=Reports--Trial%20Balance&login=user&js=1&report=trial_balance'.format(user_ip)
    chrome.get(url)
    time.sleep(1)
    
    chrome.find_element_by_xpath('/html/body/form/input[1]').click()
    
    time.sleep(0.5)  
    # print(chrome.page_source)
    soup = BeautifulSoup(chrome.page_source, 'lxml')
    
    find_th = soup.find_all("tr")[3].find_all("th", attrs={'class' : 'listheading'})
    for i in find_th:
        print(i)

    
    
    # df = pd.DataFrame(df_list, columns = ["No", "帳戶", "說明", "起始餘額" ,"借方", "貸方", "餘額"])
    
if __name__ == "__main__":
    login_sql_ledger("user", "6263")
    # asset_account(user_ip)
    Spreadsheet(user_ip)