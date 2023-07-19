import sys
from datetime import datetime

from PySide2.QtWidgets import QApplication, QFileDialog, QMainWindow
from PySide2.QtUiTools import QUiLoader
from PySide2.QtCore import Signal, QThread

from data.getFilePath import get_number_of_batches
from model.outputExcel import output_excel
from data.getDbData import *


targetDir = r'\\192.168.0.11\Data\GeRun\SecondaryWire\4#\VT报告文件db3'  # 格润4#


# 将查询数据库放入子线程中运行
class Worker(QThread):
    # 定义查询数据库的信號
    queryDbFinishedStatu = Signal()
    queryDbError = Signal()

    def __init__(self):
        super().__init__()
        self.folder_path = targetDir

    def select_folder(self, new_folder_path):
        self.folder_path = new_folder_path

    def run(self):
        try:
            total_version_defects_dict, total_query_filepath_info = get_db_defect_data(self.folder_path)  # 获取数据
            total_yield_dur_dict = get_yied_dur_data(self.folder_path)[0]  # 良率
            total_ver_list = get_yied_dur_data(self.folder_path)[1]  # 版本
            output_excel(total_version_defects_dict, total_yield_dur_dict, total_ver_list,
                         total_query_filepath_info['totalBatchVerPath'],
                         total_query_filepath_info['totalVrpDict'],
                         get_number_of_batches(self.folder_path))
            self.queryDbFinishedStatu.emit()  # 发送完成信号
        except Exception as e:
            self.queryDbError.emit(str(e))


class MainWin:

    def __init__(self):
        # 从文件中加载UI定义
        self.ui = QUiLoader().load('ui/mainwin.ui')
        self.ui.pathLineEdit.setText(targetDir)
        self.ui.pathBut.clicked.connect(self.select_path)
        self.ui.excelBut.clicked.connect(self.start_gen_excel_thread)
        self.thread = Worker()
        self.thread.queryDbError.connect(self.handle_query_db_error)
        self.folderPath = targetDir

    def handle_query_db_error(self, error):
        self.ui.logPlainTextEdit.appendPlainText(str(self.get_now_time()) + "查询出错：" + error)

    def select_path(self):
        print("选择路径")
        mW = QMainWindow()
        file_dialog = QFileDialog(mW)
        self.folderPath = file_dialog.getExistingDirectory(mW, 'Select Folder')
        if self.folderPath:
            self.ui.pathLineEdit.setText(self.folderPath)
            self.ui.logPlainTextEdit.appendPlainText(
                f'{self.get_now_time()}目录已更新为 {self.folderPath}')
            self.thread.select_folder(self.folderPath)

    def start_gen_excel_thread(self):
        self.ui.excelBut.setDisabled(True)
        self.ui.logPlainTextEdit.appendPlainText(
            str(self.get_now_time()) + "正在开始查询数据，请稍等")

        self.thread.select_folder(self.folderPath)  # 传递文件夹路径
        self.thread.queryDbFinishedStatu.connect(self.finish_gen_excel_thread)
        self.thread.start()

    def finish_gen_excel_thread(self):
        self.ui.excelBut.setDisabled(False)
        self.ui.logPlainTextEdit.appendPlainText(
            str(self.get_now_time()) + "完成查询")
        self.thread.quit()

    def get_now_time(self):
        return datetime.now().strftime('%Y/%m/%d %H:%M:%S ')


if __name__ == '__main__':
    app = QApplication(sys.argv)
    mainWin = MainWin()
    mainWin.ui.show()
    app.exec_()
