import os
from datetime import datetime


class Setting(object):

    def __init__(self):
        self._BASE_DIR = os.path.join(os.getcwd(), 'conf')
        self._MODEL_DIR = os.path.join(os.getcwd(), '模板')
        self.TEMP_DIR = os.path.join(os.getcwd(), 'tmp')
        if not os.path.exists(self.TEMP_DIR):
            os.mkdir(self.TEMP_DIR)
        self._DEFAULTINI = os.path.join(self._BASE_DIR, 'default.ini')
        self._TASK_CONFINI = os.path.join(self._BASE_DIR, 'task_conf.ini')
        self._conf = None

    def get_task(self):
        with open(self._TASK_CONFINI, 'r', encoding="utf8") as f:
            conf_dict = {}
            for var in f.readlines():
                key, value = var.strip().split('=')
                conf_dict[key] = value
        self._conf = conf_dict
        return conf_dict

    def load_conf(self, conf=None):
        if not conf:
            conf = self._DEFAULTINI
        with open(conf, 'r', encoding="utf8") as f:
            conf_dict = {}
            for var in f.readlines():
                key, value = var.strip().split('=')
                conf_dict[key] = value
        return conf_dict

    def write_conf(self, task):
        if task in self.get_task().keys():
            raise Exception('任务名已经存在')
        confini = 'conf_%s.ini'%datetime.today().strftime('%H%M%S')
        with open(self._TASK_CONFINI, 'a', encoding='utf8') as f:
            f.writelines('%s=%s\n'%(task, confini))
        new_ini = os.path.join(self._BASE_DIR, confini)
        with open(new_ini, 'w', encoding='utf8') as f:
            pass

    def get_conf(self, file):
        return os.path.join(self._BASE_DIR, file)

    def get_model(self, file):
        return os.path.join(self._MODEL_DIR, file)





if __name__ == '__main__':
    a = Setting()
    print(a.get_task())