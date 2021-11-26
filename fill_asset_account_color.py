# 填色
import openpyxl
from openpyxl.styles import PatternFill

def fill_asset_account_color():
    file_name = "SQL-Ledger.xlsx"
    sheet_name = "會計科目表"

    wb = openpyxl.load_workbook(file_name)
    wb_name = wb.sheetnames

    print(wb_name)

    # 選定表
    wb_sheet_select = wb[wb_name[0]]
    print(wb_sheet_select)

    # solid = 實色填充
    fill_color = PatternFill('solid', fgColor="FFBB02")

    rows = wb_sheet_select.max_row
    print(rows)

    check_list = [1, 21,22,231,232,24,251,252,261,262,263,264,271,272,281,291,3,4,5,6,7,821,822,831,832,891]

    for i in range(2, rows):
        cell_select = wb_sheet_select.cell(row = i, column=1)
        if (int(cell_select.value) in check_list):
            for j in range(1, wb_sheet_select.max_column+1):
                wb_sheet_select.cell(row=i, column=j).fill = fill_color
            print("found")
        i = i+1
        
    wb.save(file_name)