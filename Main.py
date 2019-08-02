import sys
import Surface
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox
from PyQt5.QtCore import QBasicTimer, QStringListModel, QTimer, QThread, pyqtSignal
from PyQt5.QtGui import QIcon
from Function import ManageClass


class ThreadTransfer(QThread):
    signOut = pyqtSignal(str, float)

    def __init__(self, zip_file_path, excel_file_path, type):
        super(ThreadTransfer, self).__init__()
        self.zip_file_path = zip_file_path
        self.excel_file_path = excel_file_path
        self.type = type

    def run(self):
        manage = ManageClass(self.zip_file_path, self.excel_file_path)

        message_zip_deal = manage.get_zip_file()
        if message_zip_deal[0]:
            self.signOut.emit('压缩包成功解压', 0)
            message_excel_deal = manage.get_excel_file()
            if message_excel_deal[0]:
                self.signOut.emit('Excel文件读取成功', 0)
            else:
                self.signOut.emit('Excel文件读取失败，开启默认命名', 0)

            manage.coding_conversion()

            list_filename = manage.return_list_filename()
            # proportion = 100/len(list_filename)

            for index, filename in enumerate(list_filename):
                if self.type == 1:
                    message_img_deal = manage.convert_image(filename)
                    if message_img_deal[0]:
                        self.signOut.emit('文件转换成功,文件名：{0}'.format(filename), 90)
                else:
                    message_img_deal = manage.convert_image_colours(filename)
                    if message_img_deal[0]:
                        self.signOut.emit('文件转换成功,文件名：{0}'.format(filename), 90)

            manage.del_temp_file()
            self.signOut.emit('删除临时文件夹', 100)
        else:
            self.signOut.emit('压缩包解压失败,程序停止,错误信息:{0}'.format(message_zip_deal[1]), 100)


class QRCodeTransfer(QMainWindow, Surface.Ui_MainWindow):
    """
    Interface the user watch
    """
    def __init__(self):
        QMainWindow.__init__(self)
        Surface.Ui_MainWindow.__init__(self)
        self.setupUi(self)

        self.message = []
        self.slm = QStringListModel()

        self.workthread = None
        # self.timer = QBasicTimer()  # 初始化一个时钟
        # self.timer = QTimer()
        # self.timer.timeout.connect(self.set_progress_bar)
        # self.timer.start(100)
        self.radioButton_colours.toggled.connect(self.change_type)
        self.radioButton_blank.toggled.connect(self.change_type)
        self.step = 0  # 进度条的值
        self.type = 1
        self.progressBar_Progress.setValue(0)
        self.setWindowIcon(QIcon('./icon.ico'))
        self.Button_SelectZipFile.clicked.connect(self.select_zip_file)
        self.Button_SelectExcelFile.clicked.connect(self.select_excel_file)
        self.Button_Start.clicked.connect(self.start_process)

    def select_zip_file(self):
        filename_choose, filetype = QFileDialog.getOpenFileName(self, '打开', r'./', 'Zip Files (*.zip);;All Files (*)')
        self.lineEdit_SelectZipFile.setText(filename_choose)

    def select_excel_file(self):
        filename_choose, filetype = QFileDialog.getOpenFileName(self, '打开', r'./', 'Excel Files 2003 (*.xls);;Excel Files 2007 (*.xlsx);;ALL Files (*)')
        self.lineEdit_SelectExcelFile.setText(filename_choose)

    def start_process(self):
        if not self.lineEdit_SelectZipFile.text() and not self.lineEdit_SelectExcelFile.text():
            QMessageBox.information(self, '提示', '请选择压缩包文件和Excel文件!')
        else:
            pass
        self.workthread = ThreadTransfer(self.lineEdit_SelectZipFile.text(),
                                         self.lineEdit_SelectExcelFile.text(),
                                         self.type)

        self.workthread.signOut.connect(self.list_add)
        self.Button_Start.setEnabled(False)
        self.Button_Start.setText('正在处理')
        self.workthread.start()

    def set_progress_bar(self):
        self.step += 1
        self.progressBar_Progress.setValue(self.step)

    def list_add(self, message, statu):
        self.message.append(message)
        self.slm.setStringList(self.message)
        self.listView_Info.setModel(self.slm)
        self.listView_Info.scrollToBottom()
        self.progressBar_Progress.setValue(statu)
        if statu >= 100:
            self.Button_Start.setEnabled(True)
            self.Button_Start.setText('开始处理')
            QMessageBox.information(self, "提示", "程序处理完成")

    def change_type(self):
        if self.radioButton_blank.isChecked():
            self.type = 1
        else:
            self.type = 2


if __name__ == '__main__':
    app = QApplication(sys.argv)
    # MainWindow = QMainWindow()
    ui = QRCodeTransfer()
    # ui.setupUi(MainWindow)
    ui.show()
    sys.exit(app.exec_())
