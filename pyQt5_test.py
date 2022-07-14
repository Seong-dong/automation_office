import sys
import os
from PyQt5.QtWidgets import *
from PyQt5 import uic
import mdPosMerge_selectFile as mdMerge_class
import saPosMerge_selectFile as sdMerge_class

#UI파일 연결
#단, UI파일은 Python 코드 파일과 같은 디렉토리에 위치해야한다.
#dir_path = os.path.dirname(os.path.abspath(__file__))
#form_class = uic.loadUiType("./merge_py.ui")[0]
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

form = resource_path('merge_py.ui')
form_class = uic.loadUiType(form)[0]

#화면을 띄우는데 사용되는 Class 선언
class WindowClass(QMainWindow, form_class) :
    def __init__(self) :
        super().__init__()
        self.setupUi(self)
        """
        시그널 이력 부분
        시그널이 작동할때 실행되는 기능은 보통 이 클래스의 멤버함수로 작성.
        """
        self.mdMerge.clicked.connect(self.mdMerge_func)
        self.saMerge.clicked.connect(self.saMerge_func)
    def mdMerge_func(self):
        print("md")
        mdMerge_class.run_func()

    def saMerge_func(self):
        print("sd")
        sdMerge_class.run_func()

if __name__ == "__main__" :
    #QApplication : 프로그램을 실행시켜주는 클래스
    app = QApplication(sys.argv)

    #WindowClass의 인스턴스 생성
    myWindow = WindowClass()

    #프로그램 화면을 보여주는 코드
    myWindow.show()

    #프로그램을 이벤트루프로 진입시키는(프로그램을 작동시키는) 코드
    app.exec_()