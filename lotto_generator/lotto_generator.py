import sys
import os
import random
from PyQt5.QtWidgets import *
from PyQt5 import uic

#form_class = uic.loadUiType("lotto_generator.ui")[0]

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

form = resource_path('lotto_generator.ui')
lotto_generator = uic.loadUiType(form)[0]

class WindowClass(QMainWindow, lotto_generator) :
    global btn_ck
    def __init__(self) :
        super().__init__()
        self.setupUi(self)

        # self.lineEdit_in.textChanged.connect(self.lineCtr)
        #button 클리거
        #self.makeNum.clicked.connect(self.makeNum_func)
        self.makeNum.clicked.connect(self.makeNum_multiple)

    #아래는 함수 정의
    def makeNum_multiple(self):
        self.resultNum_1.setText(self.makeNum_func())
        self.resultNum_2.setText(self.makeNum_func())
        self.resultNum_3.setText(self.makeNum_func())
        self.resultNum_4.setText(self.makeNum_func())
        self.resultNum_5.setText(self.makeNum_func())

    def makeNum_func(self):
        strNum = self.strNum.text()
        endNum = self.endNum.text()
        xitNum = self.xitNum.text()
        numList = []

        #추출 대상 설정
        for i in range(int(strNum), int(endNum)+1):
            numList.append(i)

        #중복 제거 대상 만들기 set
        if xitNum == "":
            print("제외값없음")
        else:
            setXitNum = set(xitNum)
            if "," in setXitNum:
                setXitNum.remove(",")
                print(setXitNum)
            else:
                print(setXitNum)

            setXitNumList = list(setXitNum)

        #중복 제거대상 제거 시작
        #for i in setXitNumList:
        for i in range(len(setXitNumList)+1):
            values = i
            print("삭제대상", values)
            if int(values) in numList:
                numList.remove(int(values))
        print("제거완료 :",numList)

        randomList = random.sample(numList,6)
        print(randomList)
        print("결과입력시작")
        #self.setValues_func(randomList)
        resultValues = str(randomList)
        resultValues.strip("[")
        resultValues.strip("]")

        return resultValues
        #self.resultNum_1.setText(resultValues)
        #print(self.lineEdit_in.text())

if __name__ == "__main__" :
    # QApplication : 프로그램을 실행시켜주는 클래스
    app = QApplication(sys.argv)

    # WindowClass의 인스턴스 생성
    myWindow = WindowClass()

    # 프로그램 화면을 보여주는 코드
    myWindow.show()

    # 프로그램을 이벤트루프로 진입시키는(프로그램을 작동시키는) 코드
    app.exec_()
