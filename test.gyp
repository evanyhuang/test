import openpyxl #
from openpyxl.styles import Font
import os

# os.chdir 是 python 切換到電腦指定路徑的方法
os.chdir(r"/Users/evany137/Downloads") #開檔案的動作
wb = openpyxl.load_workbook('producesSales.xlsx')#開檔案的動作
sheet = wb.worksheet[0]#開檔案的動作

price_updates_dict = {'Garlic': 3.07,
                      'Lemon': 1.27}

for rowNum in range (2, sheet.max_row, 1):#掃描excel檔案的第一個column
    produceName = sheet.cell(rowNum, 1).value
    if produceName in price_updates_dict:
        sheet.cell(rowNum, 2).value = price_updates_dict[produceName]#這個cell的value我要改成ＸＸＸ

        sheet.cell(rowNum, 2).font = Font(color='FF0000')#這個cell的format我要改成ＸＸＸ

wb.save('produceSales_update.xlsx')
