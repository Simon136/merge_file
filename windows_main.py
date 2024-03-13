from PyQt5.QtWidgets import QMainWindow, QApplication
from plug_Qt.untitled import Ui_MainWindow
from PyQt5.QtWidgets import *
from PyQt5 import QtWidgets
import os
import sys
from plug_Qt import function_function


class MainWindow(QMainWindow, Ui_MainWindow):

    def __init__(self, parent=None):
        super(MainWindow, self).__init__(parent)
        self.setupUi(self)
        # 初始化各类文件
        self.template_path = ""
        self.merge_path_file = ""
        self.try_name = ""
        self.merge_file = ""
        self.save_file_path = ""
        self.header_type = ""
        self.row_content = ""
        # 定义解压文件夹存在
        self.unzip_folder = ""
        # 定义初始操作，避免重复解压文件
        self.first_use = True
        # 定义模板表读取表头行
        self.final_row = int
        self.pushButton.setEnabled(False)
        self.execute_judge = bool
        # 定义模板表头内容数组
        self.header_columns = list
        # 定义处理文件夹存在
        self.merge_path_target = str
        self.start_main()

    def start_main(self):
        self.try_name = self.comboBox.currentText()
        start_text = "欢迎使用海雅合并软件\n开发团队：信息部，有任何使用问题与建议请联系信息部\n仅供公司内部使用，禁止外泄！！！！"
        self.textEdit.insertPlainText(start_text)
        self.textEdit.append("= = = = = = = = = = = = = = = = = = = = =")
        self.textEdit.append("目前选择文件合并类型为："+self.try_name)
        self.header_type = self.header_try.currentText()
        content_text = "默认选择合并为第一行表头合并："+self.header_type
        self.textEdit.append(content_text)
        self.textEdit.append("= = = = = = = = = = = = = = = = = = = = =")
        self.label_5.setText("标题所在行：")
        self.header_try.currentIndexChanged.connect(self.button_state)
        self.comboBox.currentIndexChanged.connect(self.selectionchange)
        self.optional_module.clicked.connect(self.select_template_function)
        self.select_file.clicked.connect(self.select_merge_file)
        self.save_button.clicked.connect(self.select_save_file)
        self.execute.clicked.connect(self.execute_function)
        self.pushButton.clicked.connect(self.merge_application)
        self.reset_button.clicked.connect(self.reset_function)

    def input_text(self):
        self.textEdit.insertPlainText("解压开始\n请稍等机器人运行……")
        self.textEdit.append(self.comboBox.currentText())

    def button_state(self):
        self.header_type = self.header_try.currentText()
        if self.header_type == "按标题所在行合并":
            self.label_5.setText("标题所在行：")
            self.textEdit.append("选择按标题所在行合并")
        else:
            self.label_5.setText("标题名称：")
            self.textEdit.append("选择输入关键字标题名称合并")

    def selectionchange(self):
        """
        更改文件选择类型
        :return:
        """
        self.textEdit.append("目前选择文件合并类型已经更改为："+self.comboBox.currentText())
        self.merge_file = self.file_input.text()
        if self.merge_file != "":
            self.file_input.setText(None)
            self.textEdit.append("已更换合并类型，请重新选择合并文件")
        self.try_name = self.comboBox.currentText()

    def select_template_function(self):
        """
        选择合并文件模板
        :return:
        """
        fileName, fileType = QtWidgets.QFileDialog.getOpenFileName(None, "选取文件", os.getcwd(),
                                                                   "All Files(*.xlsx *.csv *.xls)")
        if fileName == "":
            self.textEdit.append("未选择对应数据模板")
        else:
            self.textEdit.append("已选择模板为：" + fileName)
        self.template_input.setText(fileName)
        self.template_path = fileName

    def select_merge_file(self):
        """
        选择合并文件或文件夹
        :return:
        """
        if self.try_name == "压缩包合并":
            fileName, fileType = QtWidgets.QFileDialog.getOpenFileName(None, "选取文件", os.getcwd(),
                                                                       "All Files(*.zip *.tar *.gz)")
            if fileName == "":
                self.textEdit.append("未选择对应数据合并压缩包")
            else:
                self.textEdit.append("已选择合并压缩包为：" + fileName)
            self.file_input.setText(fileName)
            self.merge_path_file = fileName
        else:
            fileName = QtWidgets.QFileDialog.getExistingDirectory(None, "请选择文件夹路径", "C:/")
            if fileName == "":
                self.textEdit.append("未选择对应数据合并路径")
            else:
                self.textEdit.append("已选择的合并文件夹为：" + fileName)
            self.file_input.setText(fileName)
            self.merge_path_file = fileName

    def select_save_file(self):
        """
        选择保存文件夹路径
        :return:
        """
        fileName = QtWidgets.QFileDialog.getExistingDirectory(None, "请选择文件夹路径", "C:/")
        if fileName == "":
            self.textEdit.append("未选择保存文件夹")
        else:
            self.textEdit.append("已选择解压合并保存文件夹为：" + fileName)
        self.save_input.setText(fileName)
        self.save_file_path = fileName

    def execute_function(self):
        """
        执行函数处理
        :return:
        """
        if self.merge_path_file != "":
            if self.header_type == "按标题名称合并":
                if self.lineEdit.text() == "":
                    self.textEdit.append("未输入标题名称，请输入标题名称！！！")
                    self.execute_judge = False
                else:
                    self.textEdit.append("选择按标题名称合并，标题名称为："+self.lineEdit.text())
                    self.row_content = self.lineEdit.text()
                    self.execute_judge = True
            else:
                if self.lineEdit.text() == "":
                    self.textEdit.append("未输入选择标题所在行，默认第一行合并")
                    self.row_content = 1
                    self.execute_judge = True
                else:
                    self.textEdit.append("选择按标题所在行合并，选择第"+self.lineEdit.text()+"行作为表头合并")
                    row_int = self.lineEdit.text()
                    try:
                        row_int = int(row_int)
                        self.row_content = row_int
                        self.execute_judge = True
                    except ValueError:
                        self.textEdit.append("输入非数值类型，无法选择对应行，请修改！！！")
                        self.execute_judge = False

            if self.execute_judge:
                if self.template_path != "" and self.save_file_path != "":
                    self.textEdit.append("全部材料已齐全，程序正在解压合并文件，请稍等……")
                elif self.template_path == "" and self.save_file_path != "":
                    self.textEdit.append("文件模板未选择，将使用合并文件第一个文件作为文件模板")
                elif self.template_path != "" and self.save_file_path == "":
                    self.save_file_path = "\\".join(self.merge_path_file.split("/")[:-1])
                    self.save_input.setText(self.save_file_path)
                    self.textEdit.append("保存文件夹路径未选择，将使用合并文件上一层文件路径\n"+self.save_file_path)
                else:
                    self.textEdit.append("文件模板与保存路径将全部使用默认")
                    self.save_file_path = "\\".join(self.merge_path_file.split("/")[:-1])
                    self.save_input.setText(self.save_file_path)
                    self.textEdit.append("合并文件保存路径\n"+self.save_file_path)
                self.execute.setEnabled(False)
                self.textEdit.append("= = = = = = = = = = = = = = = = = = = = =")
                print(self.final_row)
                try:
                    self.template_feedback = function_function.New_Thread(self.merge_path_file, self.save_file_path, self.row_content, self.header_type, self.template_path, self.first_use, self.unzip_folder, self.final_row)
                    self.template_feedback.finishSignal.connect(self.decompressing)
                    self.template_feedback.start()
                    self.first_use = False
                    self.unzip_folder = self.template_feedback.merge_path_file
                    self.final_row = self.template_feedback.final_row
                    print(self.final_row)
                except Exception as e:
                    self.textEdit.append(str(e))
                    self.textEdit.append("请将此信息发至信息部，信息部核查此异常情况，感谢您的反馈")
                    self.execute.setEnabled(True)
                self.execute.setEnabled(True)
            else:
                self.textEdit.append("请完善对应信息再点击执行！！")
                self.textEdit.append("= = = = = = = = = = = = = = = = = = = = =")
        else:
            self.execute_judge = False
            self.textEdit.append("未选择合并文件夹或压缩包，请选择合并文件夹或压缩包！")

    def decompressing(self, msg):
        """
        线程回传函数
        :param msg: 数据处理回传值
        :return: None
        """
        print(msg)
        if msg['状态']:
            self.textEdit.append("如下为模板的表头，请核查是否正确")
            self.textEdit.append(msg['内容'])
            self.textEdit.append("若是正确则点击确定！！！")
            self.final_row = msg['所在行']
            self.header_columns = msg['表头内容']
            self.merge_path_target = msg['文件所在文件夹']
            self.pushButton.setEnabled(True)
        else:
            self.pushButton.setEnabled(False)
            self.textEdit.append("输入内容核查失败，失败原因："+msg['内容'])

    def merge_application(self):
        """
        确定执行合并函数
        :return:
        """
        self.template_feedback.terminate()
        self.pushButton.setEnabled(False)
        self.execute.setEnabled(False)
        self.textEdit.append("流程正在执行，请稍等……")
        self.summary_thread = function_function.Summary_thread(self.header_columns, self.final_row, self.save_file_path, self.merge_path_target)
        self.summary_thread.finishSignal.connect(self.executive_feedback)
        self.summary_thread.start()

    def executive_feedback(self, msg):
        """
        合并回传函数
        :param msg:
        :return:
        """
        print(msg)
        if msg['状态']:
            self.textEdit.append('流程运行结束，以下为整理文件位置:')
            self.textEdit.append('汇总存放于：'+msg['内容']['汇总表'])
            self.textEdit.append('无法汇总记录表存放于：'+msg['内容']['无法汇总记录表'])
            self.textEdit.append('无法汇总文件夹存放于：'+msg['内容']['无法汇总文件迁移'])
        else:
            self.textEdit.append('流程运行异常错误，错误原因：')
            self.textEdit.append(msg['内容'])
        self.pushButton.setEnabled(True)
        self.execute.setEnabled(True)

    def reset_function(self):
        """
        重置操作
        :return:
        """
        # 定义初始操作，避免重复解压文件
        self.first_use = True
        # 初始化各类文件
        self.template_path = ""
        self.merge_path_file = ""
        self.try_name = ""
        self.merge_file = ""
        self.save_file_path = ""
        self.header_type = ""
        self.row_content = ""
        # 定义解压文件夹存在
        self.unzip_folder = ""
        # 定义模板表读取表头行
        self.final_row = int
        self.pushButton.setEnabled(False)
        self.execute_judge = bool
        # 定义模板表头内容数组
        self.header_columns = list
        # 定义处理文件夹存在
        self.merge_path_target = str
        self.template_input.setText(None)
        self.file_input.setText(None)
        self.save_input.setText(None)
        self.lineEdit.setText(None)
        self.try_name = self.comboBox.currentText()
        self.header_type = self.header_try.currentText()
        try:
            self.summary_thread.terminate()
        except Exception as e:
            print(e)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()
    sys.exit(app.exec_())
