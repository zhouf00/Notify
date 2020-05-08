from threading import Thread, Event
from time import ctime
from PyQt5.QtCore import QThread, pyqtSignal


class MyThread(QThread):

    counter_value = pyqtSignal()
    _signal = pyqtSignal(str)

    def __init__(self, target, args, name=""):
        QThread.__init__(self)
        self.target = target
        self.args = args
        self.is_running = True

    def run(self):
        #print("starting",self.name, "at:",ctime())
        self.res = self.target(*self.args)
        print(self.res)
        self._signal.emit(self.res)

    def stop(self):
        self.terminate()