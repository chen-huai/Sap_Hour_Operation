# global a
# a = 1
#
#
# def op1(func):
#     def warraper(*args, **kwargs):
#         a = 0
#         func(*args, **kwargs)
#
#     return warraper
#
#
# @op1
# def op3(a):
#     a += 1
#     print(3, a)
#
#
# op3(a)


from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtWidgets import QMainWindow, QApplication, QPushButton, QLabel, QVBoxLayout, QWidget


class WorkerThread(QThread):
    # 子线程信号：发送数据到主界面
    value_signal = pyqtSignal(int)

    def __init__(self):
        super().__init__()
        self.current_value = 0

    def run(self):
        while True:
            # 每隔1秒发送一次数据到主界面
            self.current_value += 1
            self.value_signal.emit(self.current_value)
            self.msleep(1000)

    # 接收主界面数据的方法
    def update_value(self, new_value):
        self.current_value = new_value


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        # 创建界面元素
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        self.value_label = QLabel('0')
        self.start_button = QPushButton('开始')
        self.add_button = QPushButton('加10')

        layout.addWidget(self.value_label)
        layout.addWidget(self.start_button)
        layout.addWidget(self.add_button)

        # 连接按钮信号
        self.start_button.clicked.connect(self.start_thread)
        self.add_button.clicked.connect(self.add_value)

        self.worker = None

    def start_thread(self):
        if not self.worker:
            self.worker = WorkerThread()
            # 连接子线程信号到主界面更新方法
            self.worker.value_signal.connect(self.update_label)
            self.worker.start()
            self.start_button.setEnabled(False)

    def update_label(self, value):
        # 更新界面显示
        self.value_label.setText(str(value))

    def add_value(self):
        if self.worker:
            # 获取当前值
            current = int(self.value_label.text())
            # 加10
            new_value = current + 10
            # 更新界面
            self.value_label.setText(str(new_value))
            # 发送新值给子线程
            self.worker.update_value(new_value)


if __name__ == '__main__':
    app = QApplication([])
    window = MainWindow()
    window.show()
    app.exec_()