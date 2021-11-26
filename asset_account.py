# 會計科目表


def asset_account(user_ip):
    book_nanme = "會計科目表"
    url = 'http://{}/sql-ledger/ca.pl?path=bin/mozilla&action=chart_of_accounts&level=Reports--Chart%20of%20Accounts&login=user&js=1'.format(user_ip)
    chrome.get(url)
    time.sleep(2)

    web_data_sc = pd.read_html(url2, encoding="utf-8")
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