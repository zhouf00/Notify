import wx
from old_data.rd_config import HeadConf, MailConf
from threading import Thread
#from my_threading import MyThread
from old_data.rd_data2 import ReadData
#from rd_data import ReadDataLines
from old_data.con_email import _send_mail


class QfmailFrame(wx.Frame):

    def __init__(self, *args, **kw):
        super(QfmailFrame, self).__init__(*args, **kw)

        # 设置大小
        self.SetSize((450, 600))
        self.SetMaxSize((450, 600))
        self.SetMinSize((450, 600))
        self.SetTitle(u"邮件群发（附件）")

        # 初始化值
        conf_name = u"config.ini"
        conf_mail_name = u"conf_mail.ini"
        self.count_nb = 0

        # 加入图形化画板
        self._init_frame()

        # 导入包

        try:
            self.xl_name, self.xl_tb_name, self.doc_name = HeadConf(conf_name).r_head()
            # 单行文件
            self.data = ReadData(self.xl_name, self.xl_tb_name, self.doc_name)
            self.full_value = self.data.rd_xl_value()
            # 多行文件
            #self.data_lines = ReadDataLines(self.xl_name, self.xl_tb_name, self.doc_name)
            #self.data_lines.rd_xl_dict()
            #self.data_lines.rd_xl_value()
        except Exception as e:
            #print(str(e).split("'")[1])
            wx.MessageBox("请检查文件是否存在 %s"%str(e).split("'")[1], u"错误", wx.OK | wx.ICON_INFORMATION)
            exit()

        # 绑定事件
        self._func_event()

        self.Center()
        self.Show()

    def _init_frame(self):
        # 设置画板
        pnl = wx.Panel(self)
        vb = wx.BoxSizer(wx.VERTICAL)
        hb1 = wx.BoxSizer(wx.HORIZONTAL)
        hb2 = wx.BoxSizer(wx.VERTICAL)
        hb3 = wx.BoxSizer(wx.HORIZONTAL)

        ######################################
        # 第一层提示信息
        ######################################
        self.gSizer = wx.FlexGridSizer(0, 6, 5, 20)
        test1 = wx.StaticText(pnl, wx.ID_ANY, u"员工清单", wx.DefaultPosition, wx.DefaultSize, 0)
        test2 = wx.StaticText(pnl, wx.ID_ANY, u"附件总数:", wx.DefaultPosition, wx.DefaultSize, 0)
        self.test3 = wx.StaticText(pnl, wx.ID_ANY, "0", wx.DefaultPosition, wx.DefaultSize, 0)
        test4 = wx.StaticText(pnl, wx.ID_ANY, u"文件数量", wx.DefaultPosition, wx.DefaultSize, 0)
        test5 = wx.StaticText(pnl, wx.ID_ANY, u"总数：", wx.DefaultPosition, wx.DefaultSize, 0)
        self.test6 = wx.StaticText(pnl, wx.ID_ANY, "0", wx.DefaultPosition, wx.DefaultSize, 0)
        self.gSizer.AddMany([(test1, 0, wx.ALIGN_RIGHT), (test2, 0, wx.ALIGN_RIGHT),
                        (self.test3, 0, wx.ALIGN_LEFT), (test4, 0, wx.ALIGN_RIGHT),
                        (test5, 0, wx.ALIGN_RIGHT), (self.test6, 0, wx.ALIGN_LEFT)])
        self.gSizer.AddGrowableCol(1, 1)
        self.gSizer.AddGrowableCol(3, 3)
        self.gSizer.AddGrowableCol(4, 1)

        hb1.Add(self.gSizer, 1, wx.EXPAND | wx.ALL, 15)

        ######################################
        # 第二层 员工名字，邮箱
        ######################################

        listctrl = wx.ListCtrl(pnl, wx.ID_ANY, pos=(10, 10), size=(300, 100),
                               style=wx.LC_REPORT | wx.LC_SORT_ASCENDING)
        listctrl.InsertColumn(0, u"姓名")
        listctrl.InsertColumn(1, u"邮箱")
        listctrl.InsertColumn(2, u"备注")

        listctrl.SetColumnWidth(0, 70)
        listctrl.SetColumnWidth(1, 150)
        listctrl.SetColumnWidth(2, 190)

        self.listctrl = listctrl
        hb2.Add(listctrl, 0, wx.EXPAND | wx.ALL, 5)

        ######################################
        # 第二层 打印框
        ######################################
        self.hb2_text_pt = wx.TextCtrl(pnl, style=wx.TE_MULTILINE)
        hb2.Add(self.hb2_text_pt, 1, wx.EXPAND | wx.ALL, 5)

        ######################################
        # 第三层 按键板块
        ######################################
        self.start_button = wx.Button(pnl, label=u"开始转换(单行)")
        #self.start_lines_button = wx.Button(pnl, label=u"开始转换(多行)")
        self.send_button = wx.Button(pnl, label=u"发送")
        hb3.AddMany([(self.start_button, 0, wx.EXPAND | wx.ALIGN_LEFT | wx.ALL, 5),  # 单行按键绑定
                     #(self.start_lines_button, 0, wx.EXPAND | wx.ALIGN_LEFT | wx.ALL, 5), # 多行按键绑定
                     (self.send_button, 0, wx.EXPAND | wx.ALIGN_RIGHT | wx.ALL, 5), ])

        ######################################
        # 总的画板板块添加
        ######################################
        vb.Add(hb1, 0, wx.EXPAND | wx.ALIGN_CENTER | wx.ALL, 5)
        vb.Add(hb2, 1, wx.EXPAND | wx.ALIGN_CENTER | wx.ALL, 5)
        vb.Add(hb3, 0, wx.EXPAND | wx.ALIGN_CENTER | wx.ALL, 5)
        pnl.SetSizer(vb)

        ######################################
        # 文件菜单
        ######################################
        menu = wx.Menu()
        self.head_Menu = menu.Append(wx.ID_ANY, u"文件设置", "")
        menu.AppendSeparator()
        self.mail_Menu = menu.Append(wx.ID_ANY, u"邮件发送设置", "")
        menu_bar = wx.MenuBar()
        menu_bar.Append(menu, u"文件")
        self.SetMenuBar(menu_bar)
        self.CreateStatusBar()

    ######################################
    #  事件绑定函数
    ######################################
    def _func_event(self):
        # 对listCtrl进行事件绑定
        #self.Bind(wx.EVT_LIST_ITEM_ACTIVATED, self.list_values, self.listctrl)
        self.Bind(wx.EVT_BUTTON, self._change_event, self.start_button)
        #self.Bind(wx.EVT_BUTTON, self._change_lines_event, self.start_lines_button)
        self.Bind(wx.EVT_BUTTON, self._send_event, self.send_button)
        #self.Bind(wx.EVT_BUTTON, self._sendlines_event, self.send_button)
        self.Bind(wx.EVT_MENU, self._Head_Menu_event, self.head_Menu)
        self.Bind(wx.EVT_MENU, self._Mail_Menu_event, self.mail_Menu)

    ######################################
    #  listCtrl赋值函数
    ######################################
    def _list_values(self, e):
        index = e.GetIndex()
        #print(index)
        data = {}
        # 设置各个列的值
        items = data.items()  # items返回列表返回遍历的元组数组

        for key, values in items:
            #print(values)
            index = self.listctrl.InsertItem(11111, values[0])
            #print("index:", index)
            for i in range(len(values[1:])):
                self.listctrl.SetItem(index, i + 1, values[i + 1])
            self.listctrl.SetItemData(index, index)

    ######################################
    #  转换读取表格（单行）转为PDF函数
    ######################################
    def _change_event(self, e):
        self.hb2_text_pt.AppendText(u"%s表格正在转换为PDF\n"%self.data.p_time())
        #Thread(target=self.change_and_print, args=(self.full_value.__next__(),)).start()
        self.cg_thread = Thread(target=self._change_and_print, args=(self.full_value,))
        self.cg_thread.start()

    ######################################
    #  转换读取表格（多行）转为PDF函数
    ######################################
    def _change_lines_event(self, e):
        self.hb2_text_pt.AppendText(u"%s表格正在转换为PDF\n"%self.data_lines.p_time())
        #Thread(target=self.change_and_print, args=(self.full_value.__next__(),)).start()
        self.cgl_thread = Thread(target=self._change_and_print_lines, args=(self.data_lines.value_list,))
        self.cgl_thread.start()

    def _send_event(self, e):
        log_value = self.data._rd_log()
        rd_mail_config = MailConf(mail_confg_file)
        email_host, email_port, mail_name, mail_passwd, mail_name_phone, name, month \
        = rd_mail_config.r_mail()
        self.sd_thread = Thread(target=self._send_and_mail, args=(log_value, email_host, email_port, mail_name, mail_passwd, mail_name_phone, name, month ))
        self.sd_thread.start()

    def _sendlines_event(self, e):
        log_value = self.data_lines._rd_log()
        rd_mail_config = MailConf(mail_confg_file)
        email_host, email_port, mail_name, mail_passwd, mail_name_phone, name, month \
        = rd_mail_config.r_mail()
        self.sd_thread = Thread(target=self._sendlines_and_mail, args=(log_value, email_host, email_port, mail_name, mail_passwd, mail_name_phone, name, month ))
        self.sd_thread.start()

    def _Head_Menu_event(self, e):
        dialog = MyDialog_Head(self)
        dialog.ShowModal()

    def _Mail_Menu_event(self, e):
        dialog = MyDialog_Mail(self)
        dialog.ShowModal()

    def _change_and_print(self, rec_value):
        try:
            while True:
                row_name, row_email, pdf_file, pdf_name = self.data._change(rec_value.__next__())
                var = ("%s %s[%s] %s\n")%(self.data.p_time(), row_name, row_email, pdf_name)
                self.data._log_file(row_name, row_email, pdf_file, pdf_name)
                self._count_file()
                self.hb2_text_pt.AppendText(var)
        except IndexError as e:
            wx.MessageBox(u"excel列长度大于word列长度\n")
        except StopIteration:
            exit()
        finally:
            self.data._open_file()
            self.hb2_text_pt.AppendText(u"全部转换完成,请检查附件内容，请检查收件人与邮箱是否正确\n")
            # print("测试成功", self.data.p_time())

    def _change_and_print_lines(self, value_list):
        try:
            for key in self.data_lines.name_dict:
                row_name, row_email, pdf_file, pdf_name = \
                    self.data_lines._change(key, self.data_lines.name_dict[key], value_list)
                var = ("%s %s[%s] %s\n") % (self.data_lines.p_time(), row_name, row_email, pdf_name)
                self._count_file()
                self.data_lines._log_file(row_name, row_email, pdf_file, pdf_name)
                self.hb2_text_pt.AppendText(var)
        except IndexError as e:
            wx.MessageBox(u"excel列长度大于word列长度\n")
        else:
            self.data_lines._open_file()
            self.hb2_text_pt.AppendText(u"全部转换完成,请检查附件内容，请检查收件人与邮箱是否正确\n")
        finally:
            pass
            # print("测试成功", self.data_lines.p_time())

    def _send_and_mail(self, rec_value, email_host, email_port, mail_name, mail_passwd, mail_name_phone, name, month):
        try:
            while True:
                row_name, row_email, pdf_name, pdf_file = rec_value.__next__()
                #print(row_name, row_email, pdf_name, pdf_file)
                var = _send_mail(email_host, email_port, mail_name, mail_passwd, mail_name_phone, name, month, row_name, row_email, pdf_name, pdf_file)
                var = "%s %s"%(self.data.p_time(), var)
                self.data._err_log(var)
                self.hb2_text_pt.AppendText(var)
        except FileNotFoundError:
            wx.MessageBox("检查文件是否存在 %s"%pdf_file)
        except StopIteration:
            exit()
        except Exception as e:
            wx.MessageBox("身份验证失败，帐号或密码有误")
        finally:
            self.data._open_logall()
            self.hb2_text_pt.AppendText(u"邮件发送完毕，请检查是否有失败\n")

    def _sendlines_and_mail(self, rec_value, email_host, email_port, mail_name, mail_passwd, mail_name_phone, name, month):
        try:
            while True:
                row_name, row_email, pdf_name, pdf_file = rec_value.__next__()
                var = _send_mail(email_host, email_port, mail_name, mail_passwd, mail_name_phone, name, month, row_name, row_email, pdf_name, pdf_file)
                var = "%s %s"%(self.data_lines.p_time(), var)
                self.data_lines._err_log(var)
                self.hb2_text_pt.AppendText(var)
        except FileNotFoundError:
            wx.MessageBox("检查文件是否存在 %s" % pdf_file)
        except StopIteration:
            exit()
        finally:
            self.data_lines._open_logall()
            self.hb2_text_pt.AppendText(u"邮件发送完毕，请检查是否有失败\n")


    def _count_file(self):
        self.count_nb += 1
        self.test3.SetLabel("%s"%self.count_nb)
        self.gSizer.Layout()


class MyDialog_Head(wx.Dialog):

    def __init__(self, *args, **kw):
        super(MyDialog_Head, self).__init__(*args, **kw)

        self.SetTitle(u"添加文件测试")
        self.SetSize((260, 300))
        self.head_conf = HeadConf("config.ini")
        xl_name, xl_tb, wd_name = self.head_conf.r_head()

        pnl = wx.Panel(self)
        gSizer = wx.BoxSizer(wx.VERTICAL)
        gSizer_b = wx.BoxSizer(wx.HORIZONTAL)

        self.xl_name = wx.TextCtrl(pnl, size=(210, 25))
        self.xl_name.SetValue(xl_name)
        self.xl_tb = wx.TextCtrl(pnl, size=(210, 25))
        self.xl_tb.SetValue(xl_tb)
        self.wd_name = wx.TextCtrl(pnl, size=(210, 25))
        self.wd_name.SetValue(wd_name)
        #################################################################################
        # 创建静态框
        ser_xl = wx.StaticBox(pnl, label=u"Excel文件名")
        ser_tb = wx.StaticBox(pnl, label=u"Excel表格名")
        ser_wd = wx.StaticBox(pnl, label=u"Word模版文件名")
        # 创建水平方向box布局管理器
        hsbox_xl = wx.StaticBoxSizer(ser_xl, wx.HORIZONTAL)
        hsbox_tb = wx.StaticBoxSizer(ser_tb, wx.HORIZONTAL)
        hsbox_wd = wx.StaticBoxSizer(ser_wd, wx.HORIZONTAL)
        # 添加到hsbox布局管理器
        hsbox_xl.Add(self.xl_name, 0, wx.EXPAND | wx.BOTTOM, 5)
        hsbox_tb.Add(self.xl_tb, 0, wx.EXPAND | wx.BOTTOM, 5)
        hsbox_wd.Add(self.wd_name, 0, wx.EXPAND | wx.BOTTOM, 5)
        # 按键添加
        self.change_affirm = wx.Button(pnl, label=u"修改配置", size=(80, 25))
        self.init_value = wx.Button(pnl, label=u"恢复默认", size=(80, 25))
        gSizer_b.AddMany([(self.change_affirm, 0, wx.CENTER | wx.ALL, 5),
                          (self.init_value, 0, wx.ALIGN_RIGHT | wx.ALL, 5)])
        # 为添加按钮组件绑定事件处理
        self.change_affirm.Bind(wx.EVT_BUTTON, self.changeaffirm)
        self.init_value.Bind(wx.EVT_BUTTON, self.initvalue)

        gSizer.AddMany([(hsbox_xl, 0, wx.CENTER | wx.ALL, 5),
                        (hsbox_tb, 0, wx.CENTER | wx.ALL, 5),
                        (hsbox_wd, 0, wx.CENTER | wx.ALL, 5),
                        (gSizer_b, 0, wx.CENTER | wx.ALL, 5)])

        pnl.SetSizer(gSizer)

    def changeaffirm(self, e):
        xl_name = self.xl_name.GetValue()
        xl_tb = self.xl_tb.GetValue()
        wd_name = self.wd_name.GetValue()
        try:
            self.head_conf.w_head(xl_name, xl_tb, wd_name)
        except Exception as e:
            wx.MessageBox(u"输入有误", u"错误", wx.OK | wx.ICON_INFORMATION)
        else:
            wx.MessageBox(u"保存成功", u"成功", wx.OK | wx.ICON_INFORMATION)
            self.Close()

    def initvalue(self, e):
        try:
            self.head_conf.w_head(u"工资条.xlsx", u"工资", u"工资条模板.docx")
        except Exception as e:
            wx.MessageBox(u"输入有误", u"错误", wx.OK | wx.ICON_INFORMATION)
        else:
            wx.MessageBox(u"恢复成功", u"成功", wx.OK | wx.ICON_INFORMATION)
            self.Close()


class MyDialog_Mail(wx.Dialog):

    def __init__(self, *args, **kw):
        super(MyDialog_Mail, self).__init__(*args, **kw)

        self.SetTitle(u"邮件文件测试")
        self.SetSize((260, 450))
        self.Mail_conf = MailConf(mail_confg_file)

        email_host, email_port, mail_name, mail_passwd, mail_name_phone, name, month = \
        self.Mail_conf.r_mail()

        pnl = wx.Panel(self)
        gSizer = wx.BoxSizer(wx.VERTICAL)
        gSizer_b = wx.BoxSizer(wx.HORIZONTAL)

        self.email_host = wx.TextCtrl(pnl, size=(210, 25))
        self.email_host.SetValue(email_host)
        self.email_port = wx.TextCtrl(pnl, size=(210, 25))
        self.email_port.SetValue(email_port)
        self.mail_name = wx.TextCtrl(pnl, size=(210, 25))
        self.mail_name.SetValue(mail_name)
        self.mail_passwd = wx.TextCtrl(pnl, size=(210, 25), style=wx.TE_PASSWORD)
        self.mail_passwd.SetValue(mail_passwd)
        self.mail_name_phone = wx.TextCtrl(pnl, size=(210, 25))
        self.mail_name_phone.SetValue(mail_name_phone)
        self.name = wx.TextCtrl(pnl, size=(210, 25))
        self.name.SetValue(name)
        self.month = wx.TextCtrl(pnl, size=(210, 25))
        self.month.SetValue(month)
        #################################################################################
        # 创建静态框
        ser_host = wx.StaticBox(pnl, label=u"邮箱默认服务器及端口")
        ser_recmail = wx.StaticBox(pnl, label=u"发件人帐号及密码")
        ser_name = wx.StaticBox(pnl, label=u"发件人名字")
        ser_month = wx.StaticBox(pnl, label=u"月份")
        # 创建水平方向box布局管理器
        hsbox_host = wx.StaticBoxSizer(ser_host, wx.VERTICAL)
        hsbox_recmail = wx.StaticBoxSizer(ser_recmail, wx.VERTICAL)
        hsbox_name = wx.StaticBoxSizer(ser_name, wx.HORIZONTAL)
        hsbox_month = wx.StaticBoxSizer(ser_month, wx.HORIZONTAL)
        # 添加到hsbox布局管理器
        hsbox_host.AddMany([(self.email_host, 0, wx.EXPAND | wx.BOTTOM, 5),
                            (self.email_port, 0, wx.EXPAND | wx.BOTTOM, 5)])
        hsbox_recmail.AddMany([(self.mail_name, 0, wx.EXPAND | wx.BOTTOM, 5),
                               (self.mail_passwd, 0, wx.EXPAND | wx.BOTTOM, 5),
                               (self.mail_name_phone, 0, wx.EXPAND | wx.BOTTOM, 5)])
        hsbox_name.Add(self.name, 0, wx.EXPAND | wx.BOTTOM, 5)
        hsbox_month.Add(self.month, 0, wx.EXPAND | wx.BOTTOM, 5)
        # 按键添加
        self.change_affirm = wx.Button(pnl, label=u"修改配置", size=(80, 25))
        self.init_value = wx.Button(pnl, label=u"恢复默认", size=(80, 25))
        gSizer_b.AddMany([(self.change_affirm, 0, wx.CENTER | wx.ALL, 5),
                          (self.init_value, 0, wx.ALIGN_RIGHT | wx.ALL, 5)])
        # 为添加按钮组件绑定事件处理
        self.change_affirm.Bind(wx.EVT_BUTTON, self.changeaffirm)
        self.init_value.Bind(wx.EVT_BUTTON, self.initvalue)

        gSizer.AddMany([(hsbox_host, 0, wx.CENTER | wx.ALL, 5),
                        (hsbox_recmail, 0, wx.CENTER | wx.ALL, 5),
                        (hsbox_name, 0, wx.CENTER | wx.ALL, 5),
                        (hsbox_month, 0, wx.CENTER | wx.ALL, 5),
                        (gSizer_b, 0, wx.CENTER | wx.ALL, 5)])
        pnl.SetSizer(gSizer)

    def changeaffirm(self, e):
        email_host = self.email_host.GetValue()
        email_port = self.email_port.GetValue()
        mail_name = self.mail_name.GetValue()
        mail_passwd = self.mail_passwd.GetValue()
        mail_name_phone = self.mail_name_phone.GetValue()
        name = self.name.GetValue()
        month = self.month.GetValue()
        try:
            self.Mail_conf.w_mail(email_host, email_port,
                                  mail_name, mail_passwd, mail_name_phone, name, month)
        except Exception as e:
            wx.MessageBox(u"输入有误", u"错误", wx.OK | wx.ICON_INFORMATION)
        else:
            wx.MessageBox(u"保存成功", u"成功", wx.OK | wx.ICON_INFORMATION)
            self.Close()

    def initvalue(self, e):
        try:
            self.Mail_conf.w_mail("smtp.exmail.qq.com", "465", "doc@windit.com.cn",
                                  "Zzqa2018", "18167000916", u"文档管理", u"3月")

        except Exception as e:
            wx.MessageBox(u"输入有误", u"错误", wx.OK | wx.ICON_INFORMATION)
        else:
            wx.MessageBox(u"恢复成功", u"成功", wx.OK | wx.ICON_INFORMATION)
            self.Close()


if __name__ == '__main__':

    data2 = ["1","abc"]
    mail_confg_file = "conf_mail.ini"
    app = wx.App()
    QfmailFrame(None)
    app.MainLoop()