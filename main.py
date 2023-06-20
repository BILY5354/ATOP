import sys
from datetime import datetime

from PySide2.QtWidgets import QApplication, QFileDialog, QMainWindow
from PySide2.QtUiTools import QUiLoader
from PySide2.QtCore import Signal, QThread

from model.outputExcel import output_excel
from data.getFilePath import get_specific_ver_vrp
from data.getDbData import *


# 目标目录
# targetDir = r'\\192.168.0.11\Data\GeRun\SecondaryWire\4#\VT报告文件db3'  # 格润4#
targetDir = r'\\192.168.0.130\gerun\日常批量仿真'

# 将查询数据库放入子线程中运行
class Worker(QThread):
    # 定义查询数据库的信號
    queryDbFinishedStatu = Signal()

    def __init__(self):
        super().__init__()

    # ! 后期这里需要修改 传入目录参数
    def run(self):
        mulVerDefects = get_db_defect_data(targetDir)  # 获取数据
        mulYiedTime = get_yied_dur_data(targetDir)[0]
        ver_list = get_yied_dur_data(targetDir)[1]  # 版本
        # 输出excel
        output_excel(mulVerDefects, mulYiedTime, ver_list, get_file_path(targetDir), get_specific_ver_vrp(targetDir))
        self.queryDbFinishedStatu.emit()  # 发送完成信号


class MainWin:

    def __init__(self):
        # 从文件中加载UI定义
        self.ui = QUiLoader().load('ui/mainwin.ui')
        self.InitUI()

    def InitUI(self):
        self.ui.pathLineEdit.setText(targetDir)
        self.ui.pathBut.clicked.connect(self.SelectPath)
        self.ui.excelBut.clicked.connect(self.StartGenExcelThread)
        self.folderPath = targetDir

    def SelectPath(self):
        print("选择路径")
        mW = QMainWindow()
        FileDialog = QFileDialog(mW)
        self.folderPath = FileDialog.getExistingDirectory(mW, 'Select Folder')
        if self.folderPath:
            self.ui.pathLineEdit.setText(self.folderPath)
            self.ui.logPlainTextEdit.appendPlainText(
                f'{self.GetNowTime()}目录已更新为 {self.folderPath}')

    def StartGenExcelThread(self):
        self.ui.excelBut.setDisabled(True)
        self.ui.logPlainTextEdit.appendPlainText(
            str(self.GetNowTime()) + "正在开始查询数据，请稍等")

        self.thread = Worker()
        self.thread.queryDbFinishedStatu.connect(self.FinishedGenExcelThead)
        self.thread.start()

    def FinishedGenExcelThead(self):
        self.ui.excelBut.setDisabled(False)
        self.ui.logPlainTextEdit.appendPlainText(
            str(self.GetNowTime()) + "完成查询")
        self.thread.quit()

    def GetNowTime(self):
        return datetime.now().strftime('%Y/%m/%d %H:%M:%S ')


if __name__ == '__main__':
    app = QApplication(sys.argv)
    mainWin = MainWin()
    mainWin.ui.show()
    app.exec_()
