import os


class HeadConf(object):

    def __init__(self, conf_name):
        self.config_file = os.path.join(os.getcwd(), conf_name)

    ######################################
    #  读取配置文件信息
    ######################################
    def r_head(self):
        f = open(self.config_file, "r", encoding="utf8")

        xl_name, xl_tb, wd_name = \
            [var.split("=")[1].strip() for var in f.readlines()]
        f.close()
        return xl_name, xl_tb, wd_name

    ######################################
    #  重写配置文件
    ######################################
    def w_head(self, xl_name, xl_tb, wd_name):
        f = open(self.config_file, "w", encoding="utf8")
        f.writelines(self.conf_file(xl_name, xl_tb, wd_name))
        f.close()

    ######################################
    #  默认配置文件格式
    ######################################
    def conf_file(self, xl_name, xl_tb, wd_name):
        return """excel_file_name = %s\nexcel_table_name = %s\nword_model_name = %s\n"""\
               %(xl_name, xl_tb, wd_name)


class MailConf(object):

    def __init__(self, conf_name):
        self.config_file = os.path.join(os.getcwd(), conf_name)

    def r_mail(self):
        f = open(self.config_file, "r", encoding="utf8")
        #print(f.readlines())
        email_host, email_port, mail_name, mail_passwd, mail_name_phone, name, month = \
        [var.split("=")[1].strip() for var in f.readlines()]
        f.close()
        return email_host, email_port, mail_name, mail_passwd, mail_name_phone, name, month

    ######################################
    #  重写配置文件
    ######################################
    def w_mail(self, email_host, email_port, mail_name, mail_passwd, mail_name_phone, name, month):
        f = open(self.config_file, "w", encoding="utf8")
        f.writelines(self.conf_mail(email_host, email_port, mail_name, mail_passwd , mail_name_phone, name, month))
        f.close()

    ######################################
    #  默认配置文件格式
    ######################################
    def conf_mail(self, email_host, email_port, mail_name, mail_passwd, mail_name_phone, name, month):
        sub = "%s工资条-%s"
        return \
            """email_host = %s\nemail_port = %s\nmail_name = %s\nmail_passwd = %s\nmail_name_phone = %s\nname = %s\nmonth = %s"""\
               %(email_host, email_port, mail_name, mail_passwd, mail_name_phone, name, month)


if __name__ == '__main__':

    mconf = MailConf("conf_mail.ini")
    print(mconf.config_file)
    #a = mconf.r_mail()
