import requests
import openpyxl
import time

from selenium import webdriver
from selenium.webdriver.chrome.options import Options

from bs4 import BeautifulSoup

opts = Options()
# opts.add_argument('--headless')  #無頭chrome
opts.add_argument('--disable-gpu')
webdriver_path = 'C:/Users/leotam/Downloads/Code/get_sql_ledger/chromedriver'

# IP -> 192.168.56.xxx 
user_ip = '192.168.56.102'
url = 'http://{}/sql-ledger/login.pl'.format(user_ip)
chrome = webdriver.Chrome(executable_path=webdriver_path,chrome_options=opts)
chrome.get(url)

login_id = chrome.find_element_by_xpath("/html/body/center/table/tbody/tr/td/form/table/tbody/tr/td/table/tbody/tr[1]/td/input")
password_id = chrome.find_element_by_xpath("/html/body/center/table/tbody/tr/td/form/table/tbody/tr/td/table/tbody/tr[2]/td/input")

user_id = 'user'
login_id.send_keys(user_id)

url2 = 'http://{}/sql-ledger/rp.pl?path=bin/mozilla&action=report&level=Reports--Trial%20Balance&login=user&js=1&report=trial_balance'.format(user_ip)
chrome.get(url2)
time.sleep(2)

chrome.find_element_by_xpath('/html/body/form/input[1]').click()
time.sleep(2)



time.sleep(30)