import os


class Setting(object):

    def __init__(self):
        BASE_DIR = os.path.join(os.getcwd(), 'conf')
        self._NAME_CONFINI = os.path.join(BASE_DIR, 'n_conf.ini')

    def load_conf(self):
        f = open(self._NAME_CONFINI, 'r', encoding="utf8")
        # for var in f.readlines():
        #     task, file =var.split('=')
        #     print(task, file)
        print([var.split("=") for var in f.readlines()])
        f.close()
        # return task, file


if __name__ == '__main__':
    a = Setting()
    print(a._NAME_CONFINI)
    a.load_conf()