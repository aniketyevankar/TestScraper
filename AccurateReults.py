import pyautogui
import clipboard
import time
import csv
import os
import xlrd
import xlwt
import webbrowser

########################################
filename = 'CA Records.xlsx' ###########
########################################


def copyAll():
    pyautogui.click(1200,185)
    time.sleep(0.5)
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(2)
    pyautogui.hotkey('ctrl', 'c')
    time.sleep(0.5)
    raw_text = clipboard.paste()
    time.sleep(0.5)
    
    company_name = raw_text.rsplit('\nMap of ')
    company_name = company_name[-1]
    company_name = company_name.split('map expand')[0]

    company_address = raw_text.rsplit('Address: ')
    company_address = company_address[-1]
    company_address = company_address.split('\n')[0]
    
    return company_name, company_address

#Read:
book = xlrd.open_workbook(filename)
sheet = book.sheets()[0]
raw_arr = os.listdir()

#Write:
book2 = xlwt.Workbook(encoding="utf-8")
sheet1 = book2.add_sheet("Address Book")

if 'spreadsheet.xls' in raw_arr:
    os.remove(os.getcwd()+'\\spreadsheet.xls')

MyFun_CSV('Street Address, Searched Address')
count = 0

for i in range(len(sheet.col_values(1))):
    c_name = sheet.col_values(0)[i]
    c_name = c_name.replace(',',' ')
    c_name = c_name.replace('  ',' ')

    c_add = sheet.col_values(1)[i]
    c_add = c_add.replace(',',' ')
    c_add = c_add.replace('  ',' ')

    base_url = 'https://www.google.com/search?q='
    
    query = str(c_name)+' '+str(c_add)
    query = query.replace(' ','+')

    url = base_url+query

    webbrowser.open_new_tab(url)
    time.sleep(5)

    try:
        company_name, company_address = copyAll()
    except:
        company_name = ''
        company_address = ''
        pass

    sheet1.write(i, 1, c_name)
    sheet1.write(i, 2, c_add)
    sheet1.write(i, 3, company_name)
    sheet1.write(i, 4, company_address)
    book2.save("spreadsheet.xls")

