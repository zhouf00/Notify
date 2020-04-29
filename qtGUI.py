import sys, cgitb, os
import qtawesome
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QGridLayout, QAction, qApp, QMenu, QPushButton,
                             QLabel, QComboBox,QTableWidget, QMessageBox, QInputDialog, QLineEdit
                             )
from PyQt5.QtCore import Qt
# from PyQt5.Qt import QCursor
from PyQt5.QtGui import QPixmap


class MainUI(QMainWindow):

    def __init__(self):
        super().__init__()
        self.BASE_DIR = os.getcwd()
        self._init_ui()
        self.show()
        self.left_chk1_conf = os.path.join(self.BASE_DIR, 'conf', 'n_conf.ini')

    def _init_ui(self):
        self.setFixedSize(800, 550)
        # self.setWindowTitle('群发软件测试v0.1')
        main_widget = QWidget()
        main_widget.setObjectName('main_widget')
        main_layout = QGridLayout()
        main_widget.setLayout(main_layout)

        # self.statusBar().showMessage("状态栏")

        # 创建左侧部件
        left_widget = QWidget()
        left_widget.setObjectName('left_widget')
        left_layout = QGridLayout()
        left_widget.setLayout(left_layout)

        # 创建右侧部件
        right_widget = QWidget()
        right_widget.setObjectName('right_widget')
        right_layout = QGridLayout()
        right_widget.setLayout(right_layout)

        # 菜单栏设置
        # self._menu_bar()

        # 工具栏
        self._tools_bar()

        # 左侧信息
        self._left_layout(left_layout)

        # 右侧信息
        self._right_layout(right_layout)

        main_layout.addWidget(left_widget, 0,0, 12, 2)
        main_layout.addWidget(right_widget, 0, 2, 12, 10)
        self.setCentralWidget(main_widget)
        # self.setWindowFlag(Qt.FramelessWindowHint)  # 隐藏边框
        # self.setWindowOpacity(0.9)  # 设置窗口透明度
        # self.setAttribute(Qt.WA_TintedBackground)   # 设置窗口背景透明
        # main_layout.setSpacing(0)   # 去缝隙


    # 菜单栏信息
    def _menu_bar(self):
        open_act = QAction('设置', self)
        about_act = QAction('关于', self)

        menu_bar = self.menuBar()

        file_menu = menu_bar.addMenu('文件')
        file_menu.addAction(open_act)

        help_menu = menu_bar.addMenu('帮助')
        help_menu.addAction(about_act)


    # 工具栏信息
    def _tools_bar(self):
        self.send_act = QAction(qtawesome.icon('fa.paper-plane', color='white'), '发送邮件', self)
        self.generate_act = QAction(qtawesome.icon('fa.play', color='green'), '生成PDF文件', self)
        self.conversion_act = QAction(qtawesome.icon('fa.random',color='#ebd200'), '转换模式', self)
        self.stop = QAction(qtawesome.icon('fa.stop', color='red'), '停止',self)

        # 按键绑定工具栏
        toolbar = self.addToolBar('操作')
        toolbar.setToolButtonStyle(Qt.ToolButtonTextUnderIcon)
        toolbar.addAction(self.send_act)
        toolbar.addAction(self.generate_act)
        toolbar.addAction(self.conversion_act)
        toolbar.addAction(self.stop)

        # 按键绑定
        self.send_act.triggered.connect(self._func_send)
        self.generate_act.triggered.connect(self._func_generate)

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
        left_label_7 = QLabel('分类')
        left_label_7.setObjectName('left_label_btn')
        left_label_8 = QLabel('收件人')
        left_label_8.setObjectName('left_label_btn')
        left_label_9 = QLabel(self)
        left_label_9.setGeometry(0,0,200,100)
        logo = QPixmap(os.path.join(os.getcwd(), 'conf', 'LOGO.png'))\
            .scaled(left_label_9.width(), left_label_9.height())
        left_label_9.setPixmap(logo)

        alist = ['工资发送']
        self.left_chk1 = QComboBox()
        self.left_chk1.addItems(alist)
        left_chk1_btnADD = QPushButton('新建')
        left_chk1_btnADD.setObjectName('left_btn')
        left_chk1_btnUSE = QPushButton('使用')
        left_chk1_btnUSE.setObjectName('left_btn')
        self.left_btn_1 = QPushButton('请添加员工邮箱文件')
        self.left_btn_1.setObjectName('left_btn_file')
        self.left_btn_2 = QPushButton('确定发送内空')
        self.left_btn_2.setObjectName('left_btn')
        self.left_chk2 = QComboBox()
        self.left_chk2.addItems(['日期'])
        self.left_chk3 = QComboBox()
        self.left_chk3.addItems(['金额'])
        self.left_chk4 = QComboBox()
        self.left_chk4.addItems(['报销人'])
        self.left_chk5 = QComboBox()
        self.left_chk5.addItems(['报销人'])
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
        layout.addWidget(self.left_btn_2, 20, 0, 1, 3)
        layout.addWidget(left_label_2, 11, 0, 1, 3)
        layout.addWidget(left_label_5, 13, 0, 1, 1)
        layout.addWidget(self.left_chk2, 13, 1, 1, 2)
        layout.addWidget(left_label_6, 15, 0, 1, 1)
        layout.addWidget(self.left_chk3, 15, 1, 1, 2)
        layout.addWidget(left_label_7, 17, 0, 1, 1)
        layout.addWidget(self.left_chk4, 17, 1, 1, 2)
        layout.addWidget(left_label_8, 19, 0, 1, 1)
        layout.addWidget(self.left_chk5, 19, 1, 1, 2)
        layout.addWidget(left_label_3, 21, 0, 1, 3)
        layout.addWidget(self.left_btn_3, 23, 0, 1, 3)
        layout.addWidget(self.left_btn_4, 25, 0, 1, 3)
        layout.addWidget(self.left_btn_5, 27, 0, 1, 3)
        layout.addWidget(left_label_9, 29, 0, 1, 3)

        layout.setAlignment(Qt.AlignTop)

        # 按键绑定任务
        left_chk1_btnADD.clicked.connect(self._func_add)
        left_chk1_btnUSE.clicked.connect(self._func_use)
        self.left_btn_1.clicked.connect(self._func_mail_file)
        self.left_btn_2.clicked.connect(self._func_change_file)
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
        alist = []
        if ok:
            alist.append(value)
            self.left_chk1.addItems(alist)
        self.left_chk1.currentText()
        f = open(self.left_chk1_conf, 'a')
        f.writelines()

    def _func_use(self):
        print('使用')

    def _func_mail_file(self):
        print('员工邮箱文件')

    def _func_change_file(self):
        print('确定发送内容')

    def _func_execl_file(self):
        print('请添加execl文件')

    def _func_word_file(self):
        print('请添加word文件')

    def _func_html_file(self):
        print('请添加html文件')

    def _func_send(self):

        print('发送成功')
        self.send_act.setDisabled(True)

    def _func_generate(self):
        print('生成附件')

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



if __name__ == '__main__':
        cgitb.enable(format='text')
        app = QApplication(sys.argv)
        win = MainUI()
        with open('qtCSS.qss', 'r') as f:
            content = f.read()
            app.setStyleSheet(content)
        sys.exit(app.exec_())