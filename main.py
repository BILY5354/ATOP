from PySide2.QtWidgets import QApplication, QFileDialog,QLineEdit,QMainWindow
from PySide2.QtUiTools import QUiLoader
import os
import sys
from data.getDbData import get_dbData
from model.outputExcel import output_excel

# 目标目录
targetDir = r'\\192.168.0.11\Data\GeRun\SecondaryWire\4#\VT报告文件db3'
#获取数据
db_data = get_dbData(targetDir)
#输出excel
output_excel(db_data)



class MainWin:

    def __init__(self):
        # 从文件中加载UI定义
        self.ui = QUiLoader().load('ui/mainwin.ui')
        self.initUI()

    def initUI(self):
        self.ui.pathLineEd.setText(targetDir) 
        
        self.ui.excelButton.clicked.connect(self.Generate)
        self.ui.pathBut.clicked.connect(self.SelectPath)

    def SelectPath(self):
        print("选择路径")
        mW = QMainWindow()
        FileDialog  = QFileDialog(mW)
        folderPath = FileDialog.getExistingDirectory(mW, 'Select Folder')
        if folderPath:
            self.ui.pathLineEd.setText(folderPath) 

    def Generate(self):
        print("生成excel")



if __name__ == '__main__':
    app = QApplication(sys.argv)
    mainWin = MainWin()
    mainWin.ui.show()
    app.exec_()