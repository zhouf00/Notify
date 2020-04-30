import os, shutil
from datetime import datetime


class Setting(object):

    def __init__(self):
        BASE_DIR = os.path.join(os.getcwd())
        self._CONF_DIR = os.path.join(BASE_DIR, 'conf')
        self._FILE_DIR = os.path.join(BASE_DIR, 'files')
        self._MODEL_DIR = os.path.join(BASE_DIR, 'model')
        self.TEMP_DIR = os.path.join(BASE_DIR, 'tmp')
        self.LOG_DIR = os.path.join(BASE_DIR, 'Logs')
        self._DEFAULTINI = os.path.join(self._CONF_DIR, 'default.ini')
        self._TASK_CONFINI = os.path.join(self._CONF_DIR, 'task_conf.ini')
        self._init(self._CONF_DIR, self._FILE_DIR, self._MODEL_DIR, self.TEMP_DIR,
                   self.LOG_DIR)
        self.conf_ini = ''
        self._task_ini = ''

    def _init(self, *args):
        for dir in args:
            if not os.path.exists(dir):
                os.mkdir(dir)

    def _info_conf(self, task_name):
        str = """task=%s\nemail_file=邮箱管理.xlsx\ndata_execl=%s.xlsx\nmodel_word=%s模板.docx\nmodel_html=%s.html\ns_str=\ne_str=\ntype=\nto=\ncc=\n"""
        str = str%(task_name, task_name, task_name, task_name)
        return str

    def get_task(self):
        with open(self._TASK_CONFINI, 'r', encoding="utf8") as f:
            conf_dict = {}
            for var in f.readlines():
                key, value = var.strip().split('=')
                conf_dict[key] = value
        self._task_ini = conf_dict
        return conf_dict

    def load_conf(self, conf=None):
        if not conf:
            conf = self._DEFAULTINI
        else:
            conf = os.path.join(self._CONF_DIR, conf)
        with open(conf, 'r', encoding="utf8") as f:
            conf_dict = {}
            for var in f.readlines():
                key, value = var.strip().split('=')
                conf_dict[key] = value
        self.conf_ini = conf_dict
        self._task_name = conf_dict['task']
        return conf_dict

    def write_task(self, task):
        if task in self.get_task().keys():
            raise Exception('任务名已经存在')
        confini = 'conf_%s.ini'%datetime.today().strftime('%H%M%S')
        with open(self._TASK_CONFINI, 'a', encoding='utf8') as f:
            f.writelines('%s=%s\n'%(task, confini))
        new_ini = os.path.join(self._CONF_DIR, confini)
        self._mkdir(task)
        shutil.copy(os.path.join(self._MODEL_DIR, '工资条.xlsx'), os.path.join(self._FILE_DIR, task, '%s.xlsx'%task))
        shutil.copy(os.path.join(self._MODEL_DIR, '工资条模板.docx'), os.path.join(self._FILE_DIR, task, '%s模板.docx' % task))
        shutil.copy(os.path.join(self._MODEL_DIR, 'signature.html'), os.path.join(self._FILE_DIR, task, '%s.html' % task))
        with open(new_ini, 'w', encoding='utf8') as f:
            f.write(self._info_conf(task))

    def write_conf(self, datas):
        with open(os.path.join(self._CONF_DIR, self._task_ini[self._task_name]),
                  'w', encoding='utf8') as f:
            for key, value in datas.items():
                print(key, value)
                f.writelines('%s=%s\n'%(key, value))

    def get_conf(self, file):
        return os.path.join(self._CONF_DIR, file)

    def get_model(self, file):
        return os.path.join(self._FILE_DIR, self._task_name, file)

    def _mkdir(self, task):
        new_dir = os.path.join(self._FILE_DIR, task)
        if not os.path.exists(new_dir):
            os.mkdir(new_dir)


if __name__ == '__main__':
    a = Setting()
    # conf = os.path.join(os.getcwd(), 'conf', 'default.ini')
    conf = os.path.join(os.getcwd(), 'conf', 'conf_171113.ini')
    with open(conf, 'r', encoding="utf8") as f:
        conf_dict = {}
        for var in f.readlines():
            print(var)