# from PyQt5.QtWidgets import QMainWindow, QAction, QApplication
# from PyQt5.QtGui import QIcon
# # from theme_manager_theme import ThemeManager
#
# class MainWindow(QMainWindow):
#     def __init__(self):
#         super().__init__()
#         self.init_ui()
#
#     def init_ui(self):
#         # 假设已经设置了菜单栏
#         menubar = self.menuBar()
#         view_menu = menubar.addMenu('视图')
#
#         # 创建主题管理器
#         self.theme_manager = ThemeManager(QApplication.instance())
#
#         # 创建切换主题的 action
#         theme_action = QAction('切换主题', self)  # 移除了 QIcon
#         theme_action.setStatusTip('切换应用主题')
#         theme_action.triggered.connect(self.toggle_theme)
#
#         # 将 action 添加到菜单
#         view_menu.addAction(theme_action)
#
#         # 也可以将 action 添加到工具栏
#         toolbar = self.addToolBar('主题')
#         toolbar.addAction(theme_action)
#
#     def toggle_theme(self):
#         self.theme_manager.toggle_theme()
#         # 可以在这里添加其他需要在主题切换后更新的UI元素
