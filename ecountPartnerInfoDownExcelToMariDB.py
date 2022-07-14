from selenium import webdriver #selenium 드라이버 import
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
#key는 값을 넣는데 사용됨
#by는 얻어오는데 사용됨

import datetime #date날짜를 사용하기 위함
import time #페이지로딩 지연을 위해 import
import os #컴퓨터에서 파일들의 위치를 지정, 찾기 위함
#import shutil #파일 이동처리를 하기위
#import win32com.client #엑셀제어
#import openpyxl

def DownXl():
    driver = webdriver.Chrome('./chromedriver')
    driver.implicitly_wait(1) # 로딩 지연
    driver.get('https://login.ecount.com/Login/')
    driver.maximize_window()

    com_code = driver.find_element_by_name('com_code')
    com_code.send_keys("179830") #회사코드를 입력해주세요
    ecount_id = driver.find_element_by_name('id')
    ecount_id.send_keys("sdjo") #아이디를 입력해주세요
    ecount_pw = driver.find_element_by_name('passwd')
    ecount_pw.send_keys("png1234!@!@") #비밀번호를 입력해주세요
    # 로그인
    btn_click = driver.find_element_by_id('save')
    btn_click.click()
    driver.implicitly_wait(3)

    #새로운 기기 로그인알림의 등록 클릭
    btn_basicRegistration = driver.find_element_by_css_selector("#ecdivpop > div > div.footer.footer-fixed > div > button.btn.btn-primary")
    btn_basicRegistration.click()
    driver.implicitly_wait(10)
    #거래처정보 ㅋ
    item_reg = driver.find_element_by_id('save')
    item_reg.click()
    time.sleep(10)
    xl_down = driver.find_element_by_css_selector('#outputExcel')
    xl_down.click()
    time.sleep(5)
    driver.implicitly_wait(10)
"""
def FileCheck():
    now = datetime.datetime.now()
    nowDate = now.strftime('%Y%m%d')
    #print(nowDate)
    #날짜정보

    fileDir = r'C:/Users/png-20210701/Downloads'
    workDir = r'G:\공유 드라이브\사후원가\사후원가_조성동'
    fileName = "ESA009M.xlsx"
    newFileName = nowDate + '_품목코드정보.xlsx'
    path = fileDir + "/" + fileName
    desPath = workDir + "/" + newFileName
    desMove = workDir + "/" + "품목코드정보.xlsx"
    #os.path.isfile(fileName)
    if os.path.isfile(path):
        print("Yes. it is a file")
        shutil.copy(path, desPath)
        shutil.move(path, desMove)
    elif os.path.isdir(path):
        print("Yes. it is a directory")
    elif os.path.exists(path):
        print("Something exist")
    else :
        print("Nothing")

def Xl_TitleRow_Del():
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel_file = excel.Workbooks.Open('G:\공유 드라이브\사후원가\사후원가_조성동\품목코드정보.xlsx')
    w_sheet = excel_file.ActiveSheet
    w_sheet.Rows(1).delete
    excel_file.save
    excel.Quit()

DownXl()
print("stop")
#FileCheck()
#Xl_TitleRow_Del()


