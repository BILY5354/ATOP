from PySide2.QtWidgets import QApplication, QFileDialog,QLineEdit,QMainWindow
from PySide2.QtUiTools import QUiLoader
import sys
from data.getDbData import get_db_defect_data,get_yied_dur_data
from model.outputExcel import output_excel

# 目标目录
targetDir = r'\\192.168.0.11\Data\GeRun\SecondaryWire\4#\VT报告文件db3'




class MainWin:

    def __init__(self):
        # 从文件中加载UI定义
        self.ui = QUiLoader().load('ui/mainwin.ui')
        self.InitUI()
    def InitUI(self):
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
        db_defect_data = get_db_defect_data(targetDir) #获取缺陷数据
        yied_dur_data_dict = get_yied_dur_data(targetDir)
        output_excel(db_defect_data) #输出excel
        print("生成完成")



if __name__ == '__main__':
    app = QApplication(sys.argv)
    mainWin = MainWin()
    mainWin.ui.show()
    app.exec_()