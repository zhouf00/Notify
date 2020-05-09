import sys, cgitb, os
import qtawesome
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QGridLayout, QAction, qApp, QMenu, QPushButton,
                             QLabel, QComboBox,QTableWidget, QMessageBox, QInputDialog, QLineEdit,QFileDialog,
                             QTableWidgetItem, QHeaderView, QDialog
                             )
from PyQt5.QtCore import Qt
from PyQt5.QAxContainer import QAxWidget
# from PyQt5.Qt import QCursor
from PyQt5.QtGui import QPixmap

from setting import Setting
from get_data import FileProcessing
from con_email import load_file, SendMail
from Mythreading import MyThread


class MainUI(QMainWindow):

    def __init__(self):
        super().__init__()
        # self.BASE_DIR = os.getcwd()
        try:
            self._init_ui()
            self._load_setting()
        except Exception as e:
            QMessageBox.critical(self, '错误', '%s' % e,QMessageBox.Yes|QMessageBox.No,QMessageBox.Yes)
        # self._init_ui()
        # self._load_setting()
        self.show()

    def _init_ui(self):
        self.setFixedSize(800, 600)
        self.setWindowTitle('群发软件测试v3.1')
        self.axWidget = QAxWidget(self)
        main_widget = QWidget(self)
        main_widget.setObjectName('main_widget')
        main_layout = QGridLayout()
        main_widget.setLayout(main_layout)
        self.setting = Setting()
        self._datas = None
        self._thread_list = []

        # self.statusBar().showMessage("状态栏")

        # 创建左侧部件
        left_widget = QWidget()
        left_widget.setObjectName('left_widget')
        left_layout = QGridLayout()
        left_widget.setLayout(left_layout)

        # 创建右侧部件
        right_widget = QWidget()
        right_widget.setObjectName('right_widget')
        self.right_layout = QGridLayout()
        right_widget.setLayout(self.right_layout)

        # 菜单栏设置
        self._menu_bar()

        # 工具栏
        self._tools_bar()

        # 左侧信息
        self._left_layout(left_layout)

        # 右侧信息
        self._right_layout(self.right_layout)

        main_layout.addWidget(left_widget, 0, 0, 12, 2)
        main_layout.addWidget(right_widget, 0, 2, 12, 10)
        self.setCentralWidget(main_widget)
        # self.setWindowFlag(Qt.FramelessWindowHint)  # 隐藏边框
        # self.setWindowOpacity(0.9)  # 设置窗口透明度
        # self.setAttribute(Qt.WA_TintedBackground)   # 设置窗口背景透明
        # main_layout.setSpacing(0)   # 去缝隙

    # 菜单栏信息
    def _menu_bar(self):
        mail_act = QAction('邮箱设置', self)
        about_act = QAction('关于', self)

        menu_bar = self.menuBar()

        file_menu = menu_bar.addMenu('文件')
        file_menu.addAction(mail_act)

        help_menu = menu_bar.addMenu('帮助')
        help_menu.addAction(about_act)

        mail_act.triggered.connect(self._func_set_mail)

    # 工具栏信息
    def _tools_bar(self):
        self.send_act = QAction(qtawesome.icon('fa.paper-plane', color='white'), '发送邮件', self)
        self.send_act.setDisabled(True)
        self.generate_act = QAction(qtawesome.icon('fa.play', color='green'), '生成PDF文件', self)
        self.generate_act.setDisabled(True)
        self.mailfresh_act = QAction(qtawesome.icon('fa.refresh',color='green'), '邮箱数据刷新', self)
        self.mailfresh_act.setDisabled(True)
        self.stop_act = QAction(qtawesome.icon('fa.stop', color='red'), '停止',self)

        # 按键绑定工具栏
        toolbar = self.addToolBar('操作')
        toolbar.setToolButtonStyle(Qt.ToolButtonTextUnderIcon)
        toolbar.addAction(self.send_act)
        toolbar.addAction(self.generate_act)
        toolbar.addAction(self.mailfresh_act)
        toolbar.addAction(self.stop_act)

        # 按键绑定
        self.send_act.triggered.connect(self._func_send)
        self.generate_act.triggered.connect(self._func_generate)
        self.mailfresh_act.triggered.connect(self._func_mailfresh)
        self.stop_act.triggered.connect(self._func_stop)

    # 配置文件加载
    def _load_setting(self, task=None):
        task_conf = self.setting.get_task()
        task_list = task_conf.keys()
        if task:
            task_ini = task_conf[task]
        else:
            task_ini = None
        conf = self.setting.load_conf(task_ini)
        self.left_chk1.clear()
        self.left_chk1.addItems(task_list)
        if task:
            task_index = list(task_list).index(task)
        else:
            task_index = list(task_list).index(conf['task'])
        self.left_chk1.setCurrentIndex(task_index)
        self.left_btn_1.setText(conf['email_file'])
        self.left_btn_3.setText(conf['data_execl'])
        self.left_btn_4.setText(conf['model_word'])
        self.left_btn_5.setText(conf['model_html'])
        self.action = FileProcessing(self.setting.get_model(self.left_btn_3.text()), self.setting.TEMP_DIR,
                                     self.setting.get_model(self.left_btn_4.text()))
        header_list = self.action.get_data
        btn_list = [self.left_chk2, self.left_chk3, self.left_chk4, self.left_chk5, self.left_chk6]
        btn_str = ['s_str', 'e_str', 'type', 'to_user', 'cc_user']
        try:
            for i in range(len(btn_list)):
                btn_list[i].clear()
                btn_list[i].addItems(header_list)
                if conf[btn_str[i]] in header_list:
                    btn_list[i].setCurrentIndex(header_list.index(conf[btn_str[i]]))
                else:
                    pass
        except Exception as e:
            print('报错%s'%e)

    # 左侧显示信息
    def _left_layout(self, layout):
        left_label_1 = QPushButton('邮箱文件设置')
        left_label_1.setObjectName('left_label')
        left_label_2 = QPushButton('发送内容设置')
        left_label_2.setObjectName('left_label')
        left_label_3 = QPushButton('发送文件设置')
        left_label_3.setObjectName('left_label')
        left_label_4 = QPushButton('任务管理')
        left_label_4.setObjectName('left_label')
        left_label_5 = QLabel('开始')
        left_label_5.setObjectName('left_label_btn')
        left_label_6 = QLabel('结束')
        left_label_6.setObjectName('left_label_btn')
        left_label_7 = QLabel('信息归属')
        left_label_7.setObjectName('left_label_btn')
        left_label_8 = QLabel('收件人')
        left_label_8.setObjectName('left_label_btn')
        left_label_10 = QLabel('抄送人')
        left_label_10.setObjectName('left_label_btn')
        left_label_9 = QLabel(self)
        left_label_9.setGeometry(0,0,200,100)
        logo = QPixmap(os.path.join(os.getcwd(), 'conf', 'LOGO.png'))\
            .scaled(left_label_9.width(), left_label_9.height())
        left_label_9.setPixmap(logo)


        self.left_chk1 = QComboBox()
        left_chk1_btnADD = QPushButton('新建')
        left_chk1_btnADD.setObjectName('left_btn')
        left_chk1_btnUSE = QPushButton('使用')
        left_chk1_btnUSE.setObjectName('left_btn')
        self.left_btn_1 = QPushButton('请添加员工邮箱文件')
        self.left_btn_1.setObjectName('left_btn_file')
        self.left_btn_2 = QPushButton('确定发送内空')
        self.left_btn_2.setObjectName('left_btn')
        self.left_chk2 = QComboBox()
        # self.left_chk2.addItems(['日期'])
        self.left_chk3 = QComboBox()
        # self.left_chk3.addItems(['金额'])
        self.left_chk4 = QComboBox()
        # self.left_chk4.addItems(['报销人'])
        self.left_chk5 = QComboBox()
        # self.left_chk5.addItems(['报销人'])
        self.left_chk6 = QComboBox()
        self.left_btn_3 = QPushButton('请添加EXCEL文件')
        self.left_btn_3.setObjectName('left_btn_file')
        self.left_btn_4 = QPushButton('请添加模板文件')
        self.left_btn_4.setObjectName('left_btn_file')
        self.left_btn_5 = QPushButton('请添加签名模板')
        self.left_btn_5.setObjectName('left_btn_file')

        layout.addWidget(left_label_4, 0, 0, 1, 3)
        layout.addWidget(self.left_chk1, 1, 0, 1, 3)
        layout.addWidget(left_chk1_btnADD, 3, 0, 1, 1)
        layout.addWidget(left_chk1_btnUSE, 3, 2, 1, 1)
        layout.addWidget(left_label_1, 5, 0, 1, 3)
        layout.addWidget(self.left_btn_1, 7, 0, 1, 3)
        layout.addWidget(self.left_btn_2, 21, 0, 1, 3)
        layout.addWidget(left_label_2, 11, 0, 1, 3)
        layout.addWidget(left_label_5, 13, 0, 1, 1)
        layout.addWidget(self.left_chk2, 13, 1, 1, 2)
        layout.addWidget(left_label_6, 15, 0, 1, 1)
        layout.addWidget(self.left_chk3, 15, 1, 1, 2)
        layout.addWidget(left_label_7, 17, 0, 1, 1)
        layout.addWidget(self.left_chk4, 17, 1, 1, 2)
        layout.addWidget(left_label_8, 19, 0, 1, 1)
        layout.addWidget(self.left_chk5, 19, 1, 1, 2)
        layout.addWidget(left_label_10, 20, 0, 1, 1)
        layout.addWidget(self.left_chk6, 20, 1, 1, 2)
        layout.addWidget(left_label_3, 22, 0, 1, 3)
        layout.addWidget(self.left_btn_3, 23, 0, 1, 3)
        layout.addWidget(self.left_btn_4, 25, 0, 1, 3)
        layout.addWidget(self.left_btn_5, 27, 0, 1, 3)
        layout.addWidget(left_label_9, 29, 0, 1, 3)

        layout.setAlignment(Qt.AlignTop)

        # 按键绑定任务
        left_chk1_btnADD.clicked.connect(self._func_add)
        left_chk1_btnUSE.clicked.connect(self._func_use)
        self.left_btn_1.clicked.connect(self._func_mail_file)
        self.left_btn_2.clicked.connect(self._func_sure_file)
        self.left_btn_3.clicked.connect(self._func_execl_file)
        self.left_btn_4.clicked.connect(self._func_word_file)
        self.left_btn_5.clicked.connect(self._func_html_file)

    # 右侧显示信息
    def _right_layout(self, layout):
        tableHeader = ['日期', '公司', '金额', '报销人']
        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setRowCount(3)
        self.table.setHorizontalHeaderLabels(tableHeader)
        layout.addWidget(self.table)

    def _func_add(self):
        value, ok = QInputDialog.getText(self, '消息框', '请输入新任务名：', QLineEdit.Normal, '工资条发送')
        if ok:
            self.setting.write_task(value)
            self.left_chk1.addItems([value])

    def _func_use(self):
        task = self.left_chk1.currentText()
        self._load_setting(task)

    def _func_mail_file(self, e):
        self.axWidget.clear()
        file = self.left_btn_1.text()
        os.system('explorer %s'%self.setting.get_conf(file))

        # if not self.axWidget.setControl('Excel.Application'):
        #     return QMessageBox.critical(self, '错误', '没有安装  %s' % 'Excel.Application')
        # self.axWidget.dynamicCall(
        #     'SetVisible (bool Visible)', 'false')  # 不显示窗体
        # self.axWidget.setProperty('DisplayAlerts', False)
        # self.axWidget.setControl(self.setting.get_conf(file))
        # self.axWidget.show()

    def _func_sure_file(self):
        print('确定发送内容')
        self.generate_act.setEnabled(True)
        if self.send_act.isEnabled():
            self.send_act.setDisabled(True)
            self.mailfresh_act.setDisabled(True)
        conf = self.setting.conf_ini
        conf['s_str'] = self.left_chk2.currentText()
        conf['e_str'] = self.left_chk3.currentText()
        conf['type'] = self.left_chk4.currentText()
        conf['to_user'] = self.left_chk5.currentText()
        conf['cc_user'] = self.left_chk6.currentText()
        self.setting.write_conf(conf)
        datas = self.action.get_excelData(conf['s_str'], conf['e_str'], conf['type'], conf['to_user'], conf['cc_user'])
        self._func_update_table(self.action.get_data, len(datas), datas, (conf['s_str'], conf['e_str']))
        self._datas = datas

    def _func_update_table(self, col_header, row, datas, index=None):
        self.table.deleteLater()
        self.table = QTableWidget()
        self.table.setRowCount(row)
        if index:
            s = col_header.index(index[0])
            e = col_header.index(index[1])
            tableHeader = col_header[s:e+1]
            self.table.setColumnCount(len(tableHeader))
            self.table.setHorizontalHeaderLabels(tableHeader)
            # self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
            # self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Interactive)
            for data, i in zip(datas, range(row)):
                for var in data[1]['value']:
                    for j in range(len(var)):
                        self.table.setItem(i, j, QTableWidgetItem(str(var[j])))
        else:
            tableHeader = col_header
            self.table.setColumnCount(len(tableHeader))
            self.table.setHorizontalHeaderLabels(tableHeader)

        self.right_layout.addWidget(self.table)

    def _func_execl_file(self):
        # print('请添加execl文件')
        file = self.left_btn_3.text()
        # print(self.setting.get_model(file))
        os.system('explorer %s' % self.setting.get_model(file)) # 打开文件

    def _func_word_file(self):
        # print('请添加word文件')
        file = self.left_btn_4.text()
        # print(self.setting.get_model(file))
        os.system('explorer %s' % self.setting.get_model(file))  # 打开文件

    def _func_html_file(self):
        # print('请添加html文件')
        file = self.left_btn_5.text()
        # print(self.setting.get_model(file))
        os.system('explorer %s' % self.setting.get_model(file))  # 打开文件

    def _func_send(self):
        self.send_act.setDisabled(True)
        mail_conf = self.setting.get_data(self.setting.get_conf('mail_conf.ini'))
        start_send = MyThread(target=self._thread_send, args=(mail_conf, self._datas, ))
        self._thread_list.append(start_send)
        start_send._signal.connect(self._message)
        start_send.start()

    def _thread_send(self, mail_conf, datas):
        try:
            send_mail = SendMail(mail_conf.values())
            html = send_mail.get_html(self.setting.get_model(self.left_btn_5.text()))
            for name, data in datas:
                send_mail.send(name, data['to_user'], data['to'], data['cc'], data['file'],
                               html)
        except Exception as e:
            return 'warning,%s'%e
        else:
            return 'information,邮件发送成功，欢迎使用！'


    def _func_generate(self):
        header = ['收件人', '抄送人', '文件']
        self.generate_act.setDisabled(True)
        self.mailfresh_act.setEnabled(True)
        emails = load_file(self.setting.get_conf(self.left_btn_1.text()))
        self._func_update_table(header, len(self._datas), None)
        start_generate = MyThread(target=self._thread_generate, args=(emails, ))
        self._thread_list.append(start_generate)
        start_generate._signal.connect(self._message)
        start_generate.start()

    def _thread_generate(self, emails, flag=True):
        datas = self._datas
        val_list = ['to', 'cc', 'file']
        if datas:
            try:
                for (name, data), i in zip(datas, range(len(datas))):
                    if flag:
                        self.action.spanned_file(name, data)
                    to = ''
                    cc = ''
                    for user in data['to_user'].strip().split('、'):
                        if not user:
                            pass
                        elif user in emails:
                            to += emails[user]+','
                        else:
                            to += user+','
                    if 'cc_user' in data and data['cc_user']:
                        for user in data['cc_user'].strip().split('、'):
                            if not user:
                                pass
                            elif user in emails:
                                cc += emails[user] + ','
                            else:
                                cc += user + ','
                    data['to'] = to
                    if cc:
                        data['cc'] = cc
                    else:
                        data['cc'] = ''
                    for j in range(len(val_list)):
                        self.table.setItem(i, j, QTableWidgetItem('%s'%data[val_list[j]]))
            except Exception as e:
                return 'warning,%s'%e
            else:
                return 'information,转换PDF成功，请检查邮箱、文件等信息后再发送'
            finally:
                if not self.send_act.isEnabled():
                    self.send_act.setEnabled(True)
        else:
            pass

    def _func_mailfresh(self):
        emails = load_file(self.setting.get_conf(self.left_btn_1.text()))
        self._thread_generate(emails, flag=False)

    def _func_stop(self):
        pass
        print(self.send_act.isEnabled())
        print(self.generate_act.isEnabled())
        # for thread in self._thread_list:
        #     thread.stop()

    def _func_set_mail(self):
        mail_conf = self.setting.get_data(self.setting.get_conf('mail_conf.ini'))
        mail_file = self.setting.get_conf('mail_conf.ini')
        ShowDialog(mail_conf, mail_file)

    def _message(self, msg):
        if msg:
            title, text = msg.split(',')
            if title == 'warning':
                QMessageBox.warning(self, title, text, QMessageBox.Yes)
            elif title == 'information':
                QMessageBox.information(self, title, text, QMessageBox.Yes)

    # 鼠标移动窗口
    # def mousePressEvent(self, event):
    #     if event.button() == Qt.LeftButton:
    #         self.m_flag = True
    #         self.m_Position = event.globalPos() - self.pos()  # 获取鼠标相对窗口的位置
    #         event.accept()
    #         self.setCursor(QCursor(Qt.OpenHandCursor))  # 更改鼠标图标
    #
    # def mouseMoveEvent(self, QMouseEvent):
    #     if Qt.LeftButton and self.m_flag:
    #         self.move(QMouseEvent.globalPos() - self.m_Position)  # 更改窗口位置
    #         QMouseEvent.accept()
    #
    # def mouseReleaseEvent(self, QMouseEvent):
    #     self.m_flag = False
    #     self.setCursor(QCursor(Qt.ArrowCursor))


class ShowDialog(QDialog):

    def __init__(self, mail_conf, mail_file):
        super().__init__()
        self._mail_conf = mail_conf
        self._mail_file = mail_file
        self._init()
        self.exec_()

    def _init(self):
        print(self._mail_conf)
        self.setWindowTitle('邮箱设置')
        self.setFixedSize(250, 300)
        self.setWindowModality(Qt.ApplicationModal)

        dialog_layout = QGridLayout()
        self.label_1 = QLabel('发件服务器：')
        self.label_2 = QLabel('端口：')
        self.label_3 = QLabel('Email帐号：')
        self.label_4 = QLabel('密码：')
        self.label_5 = QLabel('手机号：')
        self.label_6 = QLabel('名字(签名用)：')
        self.label_7 = QLabel('邮箱标题：')
        self.label_8 = QLabel('正文插入内容：')
        self.edit_1 = QLineEdit(self._mail_conf['email_host'])
        self.edit_2 = QLineEdit(self._mail_conf['email_port'])
        self.edit_3 = QLineEdit(self._mail_conf['mail_name'])
        self.edit_4 = QLineEdit(self._mail_conf['mail_passwd'])
        self.edit_4.setEchoMode(QLineEdit.Password)
        self.edit_5 = QLineEdit(self._mail_conf['mail_name_phone'])
        self.edit_6 = QLineEdit(self._mail_conf['name'])
        self.edit_7 = QLineEdit(self._mail_conf['header'])
        self.edit_8 = QLineEdit(self._mail_conf['text'])

        btn_sure = QPushButton('确定')

        dialog_layout.addWidget(self.label_1, 1, 0)
        dialog_layout.addWidget(self.edit_1, 1, 1)
        dialog_layout.addWidget(self.label_2, 2, 0)
        dialog_layout.addWidget(self.edit_2, 2, 1)
        dialog_layout.addWidget(self.label_3, 3, 0)
        dialog_layout.addWidget(self.edit_3, 3, 1)
        dialog_layout.addWidget(self.label_4, 4, 0)
        dialog_layout.addWidget(self.edit_4, 4, 1)
        dialog_layout.addWidget(self.label_5, 5, 0)
        dialog_layout.addWidget(self.edit_5, 5, 1)
        dialog_layout.addWidget(self.label_6, 6, 0)
        dialog_layout.addWidget(self.edit_6, 6, 1)
        dialog_layout.addWidget(self.label_7, 7, 0)
        dialog_layout.addWidget(self.edit_7, 7, 1)
        dialog_layout.addWidget(self.label_8, 8, 0)
        dialog_layout.addWidget(self.edit_8, 8, 1)

        dialog_layout.addWidget(btn_sure, 9, 0, 1, 2)

        self.setLayout(dialog_layout)

        btn_sure.clicked.connect(self._save_data)

    def _save_data(self):
        self._mail_conf['email_host'] =self.edit_1.text()
        self._mail_conf['email_port'] = self.edit_2.text()
        self._mail_conf['mail_name'] = self.edit_3.text()
        self._mail_conf['mail_passwd'] = self.edit_4.text()
        self._mail_conf['mail_name_phone'] = self.edit_5.text()
        self._mail_conf['name'] = self.edit_6.text()
        self._mail_conf['header'] = self.edit_7.text()
        self._mail_conf['text'] = self.edit_8.text()
        with open(self._mail_file, 'w', encoding='utf8') as f:
            for key, value in self._mail_conf.items():
                f.writelines('%s=%s\n'%(key, value))
        self.close()


if __name__ == '__main__':
    from Tool import QSSTool
    cgitb.enable(format='text')
    app = QApplication(sys.argv)
    win = MainUI()
    QSSTool.setQssToObj('qtCSS.qss', app)
    sys.exit(app.exec_())