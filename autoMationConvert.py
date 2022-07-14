import subprocess
import time
import pyautogui
import pygetwindow
import os
# import xlwings as xw
from tkinter import messagebox

#Step1
#디렉토리 파일 리스트 받음.
dirPath = r"\\192.168.0.100\data\TFT\300.mdData\TEST_blcAutoMerge\PDF"
saveDir = r"\\192.168.0.100\data\TFT\300.mdData\TEST_blcAutoMerge\ConvertExcel"
"""
dirPath = r"\\192.168.0.100\data\TFT\300.mdData\blcAutoMerge\PDF"
saveDir = r"\\192.168.0.100\data\TFT\300.mdData\blcAutoMerge\ConvertExcel"
"""
fileList = os.listdir(dirPath)

print("fileList : {}".format(fileList))

#try:
for i in fileList :
    filePath = dirPath + "/" + i
    print("{} 파일을 실행합니다.".format(filePath))
    time.sleep(2)
    # 한셀로 pdf파일 실행
    callFile = subprocess.Popen([r"C:/Program Files (x86)/Hnc/Office 2020/HOffice110/Bin/HCell.exe", filePath])
    print("실행완료")

    # 5초 대기.
    print("Sleep 5 seconds from now on...")

    print("wake up!")

    try:
        win = pygetwindow.getWindowsWithTitle('한셀')[0]
        print("정상동작")
    except:
        print("예외처리")
        time.sleep(5)
        win = pygetwindow.getWindowsWithTitle('한셀')[0]
        print(win)

    win.activate()

    #파일저장작업
    time.sleep(1)
    pyautogui.hotkey('ctrl', 's')
    time.sleep(1)
    pyautogui.hotkey('ctrl', 'c')
    time.sleep(1)
    pyautogui.press("f4")
    time.sleep(1)
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(1)
    pyautogui.press("delete")
    time.sleep(2)
    pyautogui.write(saveDir)
    time.sleep(1)
    pyautogui.press("enter")
    time.sleep(2)
    pyautogui.hotkey('alt', 's')
    time.sleep(4)
    win.close()
    #os.system("taskkill /f /im HCell.exe")

    time.sleep(8)

messagebox.showinfo("complete","변환완료완료")

#except:
 #   messagebox.showinfo("complete", "작업이 실행되지 않았습니다.")

#Step2
"""
wb = xw.Book(r"\\192.168.0.100\data\TFT\300.mdData\blcAutoMerge/Merge.xlsm")
#엑셀 VBA의 매크로 함수 'test'를 파이썬 함수로 지정
macro_test = wb.macro('test')

#VBA 함수 실행
macro_test()

#함수를 실행한 엑셀파일 따로 저장하기
wb.save(r"\\192.168.0.100\data\TFT\300.mdData\blcAutoMerge/py1_result.xlsm")

#WorkBook 객체 닫기
wb.close()
wb.close
print("종료")
"""

"""
#subprocess.run("C:\Program Files (x86)\Hnc\Office 2020\HOffice110\Bin\HCell.exe")
#한셀로 pdf파일 실행
callFile = subprocess.Popen([r"C:/Program Files (x86)/Hnc/Office 2020/HOffice110/Bin/HCell.exe", r"C:/Users/png-20210701/Downloads/PO 정보 이상있는 경우/Purchase Order_PO1981-Japan_1634079670244.pdf"])

print("실행완료")

#5초 대기.
print ("Sleep 5 seconds from now on...")
time.sleep(5)
print("wake up!")

win = pygetwindow.getWindowsWithTitle('한셀')[0]
print(win)
win.activate()

saveDir = "C:/Users/png-20210701/Downloads/"
pyautogui.hotkey('ctrl', 's')
pyautogui.hotkey('ctrl', 'c')
pyautogui.press("f4")
pyautogui.hotkey('ctrl', 'a')
pyautogui.press("delete")
time.sleep(2)
pyautogui.write(saveDir)
pyautogui.press("enter")
time.sleep(2)
pyautogui.hotkey('alt', 's')
time.sleep(2)
win.close()


#pyautogui.keyDown("a")

#pyautogui.keyDown("Purchase Order_PO1981-Japan_1634079670244.pdf")
"""