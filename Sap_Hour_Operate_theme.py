import sys
import os
import re
import time
import math
import pandas as pd
import csv
import copy
import numpy as np
import win32com.client
import datetime
import chicon  # 引用图标
# from PyQt5 import QtCore, QtGui, QtWidgets
# from PyQt5.QtWidgets import QApplication, QMainWindow
from PyQt5.QtWidgets import QApplication, QFileDialog, QMainWindow, QMessageBox, QVBoxLayout, QPushButton, QAction, QLabel
from PyQt5.QtCore import QDate
from PyQt5.QtGui import QIcon, QFont
from Get_Data import *
from File_Operate import *
from Sap_Function import *
from Sap_Operate_Ui import Ui_MainWindow
from Data_Table import *
from Logger import *
from Excel_Field_Mapper import excel_field_mapper
from theme_manager_theme import ThemeManager
from Revenue_Operate import *
import qt_material
import shutil
import logging

# 导入自动更新模块
from auto_updater import AutoUpdater, UI_AVAILABLE
from auto_updater.config_constants import CURRENT_VERSION





class MyMainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super(MyMainWindow, self).__init__(parent)
        self.setupUi(self)

        # 初始化日志记录器
        self.logger = logging.getLogger(__name__)

        self.theme_manager = ThemeManager(QApplication.instance())
        self.init_theme_action()

        # 设置默认窗口大小和字体
        self.resize(1400, 950)
        font = QFont()
        font.setPointSize(11)
        QApplication.instance().setFont(font)

        self.theme_manager = ThemeManager(QApplication.instance())

        layout = QVBoxLayout()

        toggle_button = QPushButton("Toggle Theme")
        toggle_button.clicked.connect(self.theme_manager.toggle_theme)
        layout.addWidget(toggle_button)
        


        self.actionExport.triggered.connect(self.exportConfig)
        self.actionImport.triggered.connect(self.importConfig)
        self.actionExit.triggered.connect(MyMainWindow.close)
        self.actionHelp.triggered.connect(self.showVersion)
        self.actionAuthor.triggered.connect(self.showAuthorMessage)
        self.theme_manager.set_theme("blue")  # 设置默认主题
        self.pushButton_63.clicked.connect(lambda: self.get_hour_file_url(self.lineEdit_30))
        self.pushButton_71.clicked.connect(lambda: self.get_hour_file_url(self.lineEdit_38))
        self.pushButton_76.clicked.connect(lambda: self.get_hour_file_url(self.lineEdit_37))
        self.pushButton_78.clicked.connect(lambda: self.get_hour_file_url(self.lineEdit_31))
        self.pushButton_72.clicked.connect(self.get_hour_combine_file)
        self.pushButton_77.clicked.connect(self.get_department_hour)
        self.pushButton_73.clicked.connect(self.get_average_person_hour)
        # self.pushButton_73.clicked.connect(self.get_person_hour)
        self.pushButton_79.clicked.connect(self.hour_Operate)
        self.pushButton_80.clicked.connect(self.clear_hour_gui)
        self.pushButton_81.clicked.connect(lambda: self.open_file(self.lineEdit_30.text()))
        self.pushButton_82.clicked.connect(lambda: self.open_file(self.lineEdit_37.text()))
        self.pushButton_83.clicked.connect(lambda: self.open_file(self.lineEdit_38.text()))
        self.pushButton_84.clicked.connect(lambda: self.open_file(self.lineEdit_31.text()))

        # 集成自动更新功能
        self._setup_auto_update()

        # 在状态栏显示版本号
        self._setup_status_bar()

    def init_theme_action(self):
        theme_action = QAction(QIcon('theme_icon.png'), 'Toggle Theme', self)
        theme_action.setStatusTip('Toggle Theme')
        theme_action.triggered.connect(self.toggle_theme)

        # 将 action 添加到菜单（如果有的话）
        if hasattr(self, 'menuBar'):
            view_menu = self.menuBar().addMenu('Theme')
            view_menu.addAction(theme_action)

        # # 将 action 添加到工具栏
        # toolbar = self.addToolBar('主题')
        # toolbar.addAction(theme_action)

    def toggle_theme(self):
        self.theme_manager.set_random_theme()
        # 可以在这里添加其他需要在主题切换后更新的UI元素

    def _setup_status_bar(self):
        """设置状态栏，永久显示版本号"""
        try:
            # 创建版本标签
            version_label = QLabel(f"当前版本: {CURRENT_VERSION}")
            version_label.setStyleSheet("padding: 0px 10px;")  # 添加左右内边距

            # 使用 addPermanentWidget 永久显示在状态栏右侧
            # 这样不会被其他临时消息覆盖
            self.statusBar().addPermanentWidget(version_label)

            self.logger.info(f"状态栏永久显示版本号: {CURRENT_VERSION}")
        except Exception as e:
            self.logger.error(f"设置状态栏失败: {e}", exc_info=True)

    def _setup_auto_update(self):
        """设置自动更新功能"""
        try:
            if UI_AVAILABLE:
                # 初始化自动更新器
                self.auto_updater = AutoUpdater(self)

                # 连接Update按钮到更新检查功能
                self.actionUpdate.triggered.connect(
                    self._on_check_update_clicked
                )

                # 启动时静默检查更新（开发环境会自动跳过）
                self._startup_update_check()

                self.logger.info("自动更新功能初始化成功")
            else:
                self.logger.warning("UI组件不可用，跳过自动更新功能初始化")
                self.auto_updater = None

        except Exception as e:
            self.logger.error(f"自动更新器初始化失败: {e}", exc_info=True)
            self.auto_updater = None

    def _on_check_update_clicked(self):
        """处理Update按钮点击事件"""
        try:
            if self.auto_updater:
                self.logger.info("用户手动触发更新检查")
                # 强制检查更新并显示UI
                self.auto_updater.check_for_updates_with_ui(force_check=True)
            else:
                self.logger.warning("自动更新器未初始化")
                QMessageBox.warning(
                    self,
                    "更新功能不可用",
                    "自动更新功能初始化失败，请检查日志或联系技术支持。"
                )
        except Exception as e:
            self.logger.error(f"检查更新失败: {e}", exc_info=True)
            QMessageBox.warning(
                self,
                "检查更新失败",
                f"检查更新时发生错误：{str(e)}"
            )

    def _startup_update_check(self):
        """应用启动时静默检查更新"""
        try:
            if self.auto_updater:
                # 使用定时器延迟1秒执行，避免阻塞主窗口启动
                from PyQt5.QtCore import QTimer
                QTimer.singleShot(1000, self._perform_silent_check)

        except Exception as e:
            self.logger.error(f"启动更新检查失败: {e}", exc_info=True)

    def _perform_silent_check(self):
        """执行静默更新检查"""
        try:
            has_update, remote_version, local_version, error = \
                self.auto_updater.check_for_updates(
                    force_check=False,
                    is_silent=True
                )

            if has_update:
                self.logger.info(f"发现新版本: {remote_version} (当前版本: {local_version})")
                # 显示更新提示对话框
                reply = QMessageBox.question(
                    self,
                    "发现新版本",
                    f"检测到新版本 {remote_version} (当前版本: {local_version})\n\n"
                    f"是否立即查看更新详情?",
                    QMessageBox.Yes | QMessageBox.No,
                    QMessageBox.Yes
                )

                if reply == QMessageBox.Yes:
                    # 用户选择查看更新，显示完整更新UI
                    self.auto_updater.check_for_updates_with_ui(force_check=True)
            elif error:
                # 静默模式下仅记录错误，不显示给用户
                self.logger.debug(f"启动更新检查: {error}")

        except Exception as e:
            # 静默模式下仅记录错误，不影响应用启动
            self.logger.debug(f"静默更新检查异常: {e}")

    def closeEvent(self, event):
        """应用退出事件处理"""
        try:
            # 清理自动更新器资源
            if hasattr(self, 'auto_updater') and self.auto_updater:
                self.logger.info("正在清理自动更新器资源...")
                self.auto_updater.cleanup()
                self.auto_updater = None

        except Exception as e:
            self.logger.error(f"清理自动更新器资源失败: {e}", exc_info=True)

        # 接受关闭事件
        event.accept()

    def getConfig(self):
        # 初始化，获取或生成配置文件
        global configFileUrl
        global desktopUrl
        global now
        global last_time
        global today
        global oneWeekday
        global fileUrl

        date = datetime.datetime.now() + datetime.timedelta(days=1)
        now = int(time.strftime('%Y'))
        last_time = now - 1
        today = time.strftime('%Y.%m.%d')
        oneWeekday = (datetime.datetime.now() + datetime.timedelta(days=7)).strftime('%Y.%m.%d')
        desktopUrl = os.path.join(os.path.expanduser("~"), 'Desktop')
        configFileUrl = '%s/config' % desktopUrl
        configFile = os.path.exists('%s/config_sap_hour.csv' % configFileUrl)
        # print(desktopUrl,configFileUrl,configFile)
        if not configFile:  # 判断是否存在文件夹如果不存在则创建为文件夹
            reply = QMessageBox.question(self, '信息', '确认是否要创建配置文件', QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
            if reply == QMessageBox.Yes:
                if not os.path.exists(configFileUrl):
                    os.makedirs(configFileUrl)
                MyMainWindow.createConfigContent(self)
                MyMainWindow.getConfigContent(self)
                self.textBrowser.append("创建并导入配置成功")
            else:
                exit()
        else:
            MyMainWindow.getConfigContent(self)

    # 获取配置文件内容
    def getConfigContent(self):
        # 配置文件
        csvFile = pd.read_csv('%s/config_sap_hour.csv' % configFileUrl, names=['A', 'B', 'C'])
        global configContent
        global username
        global role
        global staff_dict
        configContent = {}
        staff_dict = {}
        # configContent[configContent.get('Business_Department','CS')] = []
        # configContent[configContent.get('Lab_1','PHY')] = []
        # configContent[configContent.get('Lab_2','CHM')] = []
        username = list(csvFile['A'])
        number = list(csvFile['B'])
        role = list(csvFile['C'])
        for i in range(len(username)):
            configContent['%s' % username[i]] = number[i]
            if role[i] == configContent.get('Business_Department', 'CS'):
                # 使用 setdefault 确保键存在且为列表类型
                staff_dict.setdefault(configContent.get('Business_Department', 'CS'), []).append(username[i])
            if role[i] == configContent.get('Lab_1', 'PHY'):
                # 使用 setdefault 确保键存在且为列表类型
                staff_dict.setdefault(configContent.get('Lab_1', 'PHY'), []).append(username[i])
            if role[i] == configContent.get('Lab_2', 'CHM'):
                # 使用 setdefault 确保键存在且为列表类型
                staff_dict.setdefault(configContent.get('Lab_2', 'CHM'), []).append(username[i])

        MyMainWindow.getDefaultInformation(self)

        try:
            self.textBrowser_4.append("配置获取成功")
        except AttributeError:
            QMessageBox.information(self, "提示信息", "已获取配置文件内容", QMessageBox.Yes)
        else:
            pass

    # 创建配置文件
    def createConfigContent(self):
        global monthAbbrev
        months = "JanFebMarAprMayJunJulAugSepOctNovDec"
        n = time.strftime('%m')
        pos = (int(n) - 1) * 3
        monthAbbrev = months[pos:pos + 3]

        config = np.array(configContent)
        configContent = [
            ['基础信息', '内容', '备注'],
            ['Business_Department', 'CS', '业务部门,名称会用于后续'],
            ['Lab_1', 'PHY', '代表实验室，会用于后续'],
            ['Lab_2', 'CHM', '代表实验室，会用于后续'],
            ['T20', 'PHY', '代表实验室，会用于后续'],
            ['T75', 'CHM', '代表实验室，会用于后续'],
            ['Hourly Rate', '金额', '备注'],
            ['CS_Hourly_Rate', 300, '客服时薪'],
            ['PHY_Hourly_Rate', 300, '物理时薪'],
            ['CHM_Hourly_Rate', 300, '化学时薪'],
            ['计划成本', '数值', '备注'],
            ['Plan_Cost_Parameter', 0.9, '实际的90%，预留10%利润'],
            ['Significant_Digits', 0, '保留几位有效数值'],
            ['实验室成本比例', '数值', '备注'],
            ['CHM_Cost_Parameter', 0.3, '给到CHM30%'],
            ['PHY_Cost_Parameter', 0.3, '给到PHY30%'],
            # 新增分配规则参数
            ['405_Item_1000', 0.5, '405分配规则'],
            ['405_Item_2000', 0.5, '405分配规则'],
            ['441_Item_1000', 0.8, '441分配规则'],
            ['441_Item_2000', 0.2, '441分配规则'],
            ['430_Item_1000', 0.8, '430分配规则'],
            ['430_Item_2000', 0.2, '430分配规则'],
            # 新增特殊MC规则
            ['T20-430-A2', 'PHY_1000/CHM_2000', '1000/2000对应的lab，强制1000设置在前，2000在后'],
            ['T20-430-A2_mc', 'T20-430-00/T75-430-00', '1000/2000对应的mc'],
            ['T75-441-A2', 'CHM_1000/PHY_2000', '1000/2000对应的lab，强制1000设置在前，2000在后'],
            ['T75-441-A2_mc', 'T75-441-00/T20-441-00', '1000/2000对应的mc'],
            ['T75-405-A2', 'CHM_1000/PHY_2000', '1000/2000对应的lab，强制1000设置在前，2000在后'],
            ['T75-405-A2_mc', 'T75-405-00/T20-405-00', '1000/2000对应的mc'],
            ['T75-405-D2', 'CHM_1000/PHY_2000', '1000/2000对应的lab，计算hour后强制都转为1000'],
            ['T75-405-D2_mc', 'T75-405-D2/T75-405-D2', '1000/2000对应的mc'],
            ['T75-405-D3', 'CHM_1000/PHY_2000', '1000/2000对应的lab，计算hour后强制都转为1000'],
            ['T75-405-D3_mc', 'T75-405-D3/T75-405-D3', '1000/2000对应的mc'],
            ['T75-441-D2', 'CHM_1000/PHY_2000', '1000/2000对应的lab，计算hour后强制都转为1000'],
            ['T75-441-D2_mc', 'T75-441-D2/T75-441-D2', '1000/2000对应的mc'],
            ['T75-441-D3', 'CHM_1000/PHY_2000', '1000/2000对应的lab，计算hour后强制都转为1000'],
            ['T75-441-D3_mc', 'T75-441-D3/T75-441-D3', '1000/2000对应的mc'],
            # Hour 参数
            ['Max_Hour', 8, '最大工作时长'],
            ['Hours_Combine_Key', "Order Number;Material Code;Primary CS",'以;分隔，数据透视字段'],
            ['Hour_Files_Import_URL', "N:\\XM Softlines\\6. Personel\\5. Personal\\Supporting Team\\2.财务\\2.SAP\\1.ODM Data - XM\\3.Hours",'Hour文件导入路径'],
            ['Hour_Files_Export_URL', "N:\\XM Softlines\\6. Personel\\5. Personal\\Supporting Team\\2.财务\\2.SAP\\1.ODM Data - XM\\3.Hours",'Hour文件导出路径'],
            ['Hour_Field_Mapping', "{'staff_id': 'staff_id','week': 'week','order_no': 'order_no','allocated_hours': 'allocated_hours','office_time':'office_time','material_code': 'material_code','item': 'item','allocated_day': 'allocated_day','staff_name': 'staff_name'}", '对应字段映射'],
            ['名称', '编号', '角色'],
            ['chen, frank', '6375108', 'CS'],
            ['chen, frank', '6375108', 'Sales'],
        ]
        df = pd.DataFrame(config)
        df.to_csv('%s/config_sap_hour.csv' % configFileUrl, index=0, header=0, encoding='utf_8_sig')
        self.textBrowser_4.append("配置文件创建成功")
        QMessageBox.information(self, "提示信息",
                                "默认配置文件已经创建好，\n如需修改请在用户桌面查找config文件夹中config_sap_hour.csv，\n将相应的文件内容替换成用户需求即可，修改后记得重新导入配置文件。",
                                QMessageBox.Yes)

    # 导出配置文件
    def exportConfig(self):
        # 重新导出默认配置文件
        reply = QMessageBox.question(self, '信息', '确认是否要创建默认配置文件', QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
        if reply == QMessageBox.Yes:
            MyMainWindow.createConfigContent(self)
        else:
            QMessageBox.information(self, "提示信息", "没有创建默认配置文件，保留原有的配置文件", QMessageBox.Yes)

    # 导入配置文件
    def importConfig(self):
        # 重新导入配置文件
        reply = QMessageBox.question(self, '信息', '确认是否要导入配置文件', QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
        if reply == QMessageBox.Yes:
            MyMainWindow.getConfigContent(self)
        else:
            QMessageBox.information(self, "提示信息", "没有重新导入配置文件，将按照原有的配置文件操作", QMessageBox.Yes)

    # 界面设置默认配置文件信息
    def getDefaultInformation(self):
        # 默认登录界面信息
        try:
            # hour界面操作
            self.spinBox_10.setValue(int(configContent['Max_Hour']))
            self.lineEdit_39.setText(configContent['Hours_Combine_Key'])
            today_hours = datetime.date.today()
            first_day = today_hours.replace(day=1)
            self.dateEdit.setDate(QDate(first_day.year, first_day.month, first_day.day))  # 当月第一天
            self.dateEdit_2.setDate(QDate.currentDate())  # 当天日期
            self.doubleSpinBox_14.setValue(float(format(float(configContent['CS_Hourly_Rate']), '.2f')))
            self.doubleSpinBox_16.setValue(float(format(float(configContent['CHM_Hourly_Rate']), '.2f')))
            self.doubleSpinBox_15.setValue(float(format(float(configContent['PHY_Hourly_Rate']), '.2f')))
            self.doubleSpinBox_11.setValue(float(format(float(configContent['Plan_Cost_Parameter']), '.2f')))
            self.doubleSpinBox_13.setValue(float(format(float(configContent['CHM_Cost_Parameter']), '.2f')))
            self.doubleSpinBox_12.setValue(float(format(float(configContent['PHY_Cost_Parameter']), '.2f')))
            self.spinBox_11.setValue(int(configContent['Significant_Digits']))
            self.lineEdit_28.setText(configContent['Lab_1'])
            self.lineEdit_29.setText(configContent['Lab_2'])
        except Exception as msg:
            self.textBrowser_4.append("错误信息：%s" % msg)
            self.textBrowser_4.append('----------------------------------')
            app.processEvents()
            reply = QMessageBox.question(self, '信息', '错误信息：%s。\n是否要重新创建配置文件' % msg, QMessageBox.Yes | QMessageBox.No,
                                         QMessageBox.Yes)
            if reply == QMessageBox.Yes:
                MyMainWindow.createConfigContent(self)
                self.textBrowser_4.append("创建并导入配置成功")
                self.textBrowser_4.append('----------------------------------')
                app.processEvents()

    def showAuthorMessage(self):
        # 关于作者
        QMessageBox.about(self, "关于",
                          "人生苦短，码上行乐。\n\n\n        ----Frank Chen")

    def showVersion(self):
        """显示版本信息"""
        try:
            version_info = f"SAP 小时操作工具\n\n当前版本: {CURRENT_VERSION}\n\n© 2022-2024 Frank Chen"
            QMessageBox.about(self, "版本", version_info)
        except Exception as e:
            self.logger.error(f"显示版本信息失败: {e}", exc_info=True)
            # 降级处理：显示固定版本
            QMessageBox.about(self, "版本", f"当前版本: {CURRENT_VERSION}")

    def getHourGuiData(self):
        guiHourData = {}
        guiHourData['Hours_Combine_Key'] = self.lineEdit_39.text()
        guiHourData['Max_Hour'] = int(self.spinBox_10.text())
        guiHourData['CS_Hourly_Rate'] = float(self.doubleSpinBox_14.text())
        guiHourData['CHM_Hourly_Rate'] = float(self.doubleSpinBox_16.text())
        guiHourData['PHY_Hourly_Rate'] = float(self.doubleSpinBox_15.text())
        guiHourData['Plan_Cost_Parameter'] = float(self.doubleSpinBox_11.text())
        guiHourData['Significant_Digits'] = float(self.spinBox_11.text())
        guiHourData['CHM_Cost_Parameter'] = float(self.doubleSpinBox_13.text())
        guiHourData['PHY_Cost_Parameter'] = float(self.doubleSpinBox_12.text())
        guiHourData['Lab_1'] = self.lineEdit_28.text()
        guiHourData['Lab_2'] = self.lineEdit_29.text()
        guiHourData['Business_Department'] = self.lineEdit_26.text()
        guiHourData['Start_Date'] = self.dateEdit.date().toString("yyyy.MM.dd")
        guiHourData['End_Date'] = self.dateEdit_2.date().toString("yyyy.MM.dd")
        return guiHourData

    # 获取文件
    def getFile(self, path):
        selectBatchFile = QFileDialog.getOpenFileName(self, '选择ODM导出文件',
                                                      '%s\\%s' % (path, today),
                                                      'files(*.docx;*.xls*;*.csv)')
        fileUrl = selectBatchFile[0]
        return fileUrl

    def get_hour_file_url(self, position):
        fileUrl = myWin.getFile(configContent['Hour_Files_Import_URL'])
        if fileUrl:
            position.setText(fileUrl)
            app.processEvents()
        else:
            self.textBrowser_2.append("请重新选择ODM文件")
            QMessageBox.information(self, "提示信息", "请重新选择ODM文件", QMessageBox.Yes)

    def get_hour_combine_file(self):
        fileUrl = self.lineEdit_30.text()
        pivot_table_key = self.lineEdit_39.text().split(';')
        if fileUrl and pivot_table_key:
            try:
                self.textBrowser_4.append("数据开始合并")
                app.processEvents()
                newData = Get_Data()
                file_data = newData.getFileTableData(fileUrl)
                # 删除
                deleteRowList = {'Order Number': ''}
                newData.deleteTheRows(deleteRowList)
                # 合并
                valus_key = ['Revenue', 'Total Subcon Cost']
                pivot_table_data = newData.pivotTable(pivot_table_key, valus_key)
                current_time = datetime.datetime.now().strftime('%Y-%m-%d %H.%M.%S')
                pivot_table_data_path = '%s\\%s' % (configContent['Hour_Files_Export_URL'], '1.order data %s.xlsx' % current_time)
                pivot_table_data_file = pivot_table_data.to_excel(pivot_table_data_path, merge_cells=False)
                self.lineEdit_37.setText(pivot_table_data_path)
                self.textBrowser_4.append("合并完成")
                self.textBrowser_4.append("文件路径：%s" % pivot_table_data_path)
            except Exception as errorMsg:
                self.textBrowser_4.append("<font color='red'>出错信息：%s </font>" % errorMsg)
                app.processEvents()
        elif pivot_table_key == []:
            self.textBrowser_4.append("请输入合并的key")
        else:
            self.textBrowser_4.append("请重新选择ODM文件")

    def update_config_content(self, update_data):
        # 创建配置字典的深拷贝以避免污染原始配置
        config_content = copy.deepcopy(configContent)
        config_content.update(update_data)
        return config_content  # 返回修改后的副本
    
    def get_department_hour(self):
        """
        计算部门工时并保存结果
        """
        order_data_path = self.lineEdit_37.text()
        hour_gui_data = myWin.getHourGuiData()
        if order_data_path:
            self.textBrowser_4.append("部门开始计算")
            app.processEvents()
            
            # 更新配置内容
            config_content = self.update_config_content(hour_gui_data)
            
            # 获取订单数据
            order_data_obj = Get_Data()
            order_datas = order_data_obj.getFileTableData(order_data_path)
            
            # 初始化结果DataFrame
            all_results = []
            
            # 调用hour方法处理每个订单
            revenue_allocator_obj = RevenueAllocator()
            for _, order_data in order_datas.iterrows():
                # 将Series转换为字典
                order_dict = order_data.to_dict()
                # 计算部门工时
                order_revenue_data = revenue_allocator_obj.allocate_department_hours(order_dict, config_content)

                all_results.extend(order_revenue_data)
            
            # 创建结果DataFrame
            result_df = pd.DataFrame(all_results)

            # material_code包含D2或D3，更新字段item=1000
            mask = result_df['material_code'].str.contains(r'D[23]', case=False, na=False, regex=True)
            result_df.loc[mask, 'item'] = '1000'

            # 生成输出文件名
            current_time = datetime.datetime.now().strftime('%Y-%m-%d %H.%M.%S')
            dept_hour_path = f"{configContent['Hour_Files_Export_URL']}\\2.dept hour {current_time}.xlsx"
            
            # 保存结果
            result_df.to_excel(dept_hour_path, index=False)
            
            # 更新UI
            self.lineEdit_38.setText(dept_hour_path)
            self.textBrowser_4.append("部门计算完成")
            self.textBrowser_4.append(f"文件路径：{dept_hour_path}")
            app.processEvents()
        else:
            self.textBrowser_4.append("请重新选择合并后的文件")
            app.processEvents()

    def get_person_hour(self):
        """
        分配人员工时并保存结果
        """
        dept_hour_path = self.lineEdit_38.text()
        if dept_hour_path:
            self.textBrowser_4.append("开始分配人员")
            app.processEvents()
            
            # 获取配置数据
            hour_gui_data = myWin.getHourGuiData()
            config_content = self.update_config_content(hour_gui_data)
            
            # 获取参数
            max_hours_per_day = int(config_content['Max_Hour'])
            start_date = datetime.datetime.strptime(config_content['Start_Date'], '%Y.%m.%d').date()
            end_date = datetime.datetime.strptime(config_content['End_Date'], '%Y.%m.%d').date()
            
            # 获取部门工时数据
            dept_hour_obj = Get_Data()
            dept_hour_datas = dept_hour_obj.getFileTableData(dept_hour_path)
            
            # 计算各部门总工时
            dept_total_hours = dept_hour_datas.groupby('dept')['dept_hours'].sum().to_dict()
            
            # 初始化结果列表
            all_results = []
            
            # 处理每个部门工时记录
            revenue_allocator_obj = RevenueAllocator()
            for _, dept_hour in dept_hour_datas.iterrows():
                # 将Series转换为字典
                dept_hour_dict = dept_hour.to_dict()
                # 将单个记录转换为列表形式
                dept_hour_list = [dept_hour_dict]
                self.textBrowser_4.append(f"处理Order Number：{dept_hour_dict['order_no']}")
                app.processEvents()
                # 分配人员工时
                person_hour_data = revenue_allocator_obj.allocate_person_average_hours(
                    dept_hour_list,
                    max_hours_per_day, 
                    start_date, 
                    end_date, 
                    staff_dict,
                    dept_total_hours,  # 添加部门总工时参数
                    config_content
                )
                all_results.extend(person_hour_data)
            
            # 创建结果DataFrame
            result_df = pd.DataFrame(all_results)
            
            # 生成输出文件名
            current_time = datetime.datetime.now().strftime('%Y-%m-%d %H.%M.%S')
            person_hour_path = f"{configContent['Hour_Files_Export_URL']}\\3.person hour {current_time}.xlsx"
            
            # 保存结果
            result_df.to_excel(person_hour_path, index=True, index_label='ID')
            
            # 打开结果文件
            os.startfile(person_hour_path)
            
            # 更新UI
            self.lineEdit_31.setText(person_hour_path)
            self.textBrowser_4.append("分配人员完成")
            self.textBrowser_4.append(f"文件路径：{person_hour_path}")
            app.processEvents()
        else:
            self.textBrowser_4.append("请先完成部门工时计算")
            app.processEvents()

    def get_average_person_hour(self):
        """
        使用平均分配方式分配人员工时并保存结果
        """
        dept_hour_path = self.lineEdit_38.text()
        if dept_hour_path:
            self.textBrowser_4.append("开始平均分配人员")
            app.processEvents()

            # 获取配置数据
            hour_gui_data = myWin.getHourGuiData()
            config_content = self.update_config_content(hour_gui_data)

            # 获取参数
            max_hours_per_day = int(config_content['Max_Hour'])
            start_date = datetime.datetime.strptime(config_content['Start_Date'], '%Y.%m.%d').date()
            end_date = datetime.datetime.strptime(config_content['End_Date'], '%Y.%m.%d').date()

            # 获取部门工时数据
            dept_hour_obj = Get_Data()
            dept_hour_datas = dept_hour_obj.getFileTableData(dept_hour_path)

            # 计算各部门总工时
            dept_total_hours = dept_hour_datas.groupby('dept')['dept_hours'].sum().to_dict()
            
            # 初始化结果列表
            all_results = []
            
            # 按部门处理工时记录
            revenue_allocator_obj = RevenueAllocator()
            for dept, total_hours in dept_total_hours.items():
                self.textBrowser_4.append(f"\n处理部门：{dept}")
                self.textBrowser_4.append(f"部门总工时：{total_hours}")
                app.processEvents()
                
                # 获取该部门的所有记录
                dept_records = dept_hour_datas[dept_hour_datas['dept'] == dept].to_dict('records')
                
                # 分配该部门的工时
                person_hour_data = revenue_allocator_obj.allocate_person_average_hours(
                    dept_records,
                    max_hours_per_day,
                    start_date,
                    end_date,
                    staff_dict,
                    {dept: total_hours},  # 只传入当前部门的总工时
                    config_content
                )
                all_results.extend(person_hour_data)
                
                # 显示分配结果
                allocated_hours = sum(record['allocated_hours'] for record in person_hour_data)
                allocation_rate = (allocated_hours / total_hours * 100) if total_hours > 0 else 0
                self.textBrowser_4.append(f"已分配工时：{allocated_hours:.2f}")
                self.textBrowser_4.append(f"分配率：{allocation_rate:.2f}%")
                app.processEvents()

            # 创建结果DataFrame
            result_df = pd.DataFrame(all_results)
            
            # 生成输出文件名
            current_time = datetime.datetime.now().strftime('%Y-%m-%d %H.%M.%S')
            person_hour_path = f"{configContent['Hour_Files_Export_URL']}\\3.person hour {current_time}.xlsx"
            
            # 保存结果
            result_df.to_excel(person_hour_path, index=True, index_label='ID')
            
            # 打开结果文件
            os.startfile(person_hour_path)
            
            # 更新UI
            self.lineEdit_31.setText(person_hour_path)
            self.textBrowser_4.append("\n分配人员完成")
            self.textBrowser_4.append(f"文件路径：{person_hour_path}")
            
            # 显示总体分配统计
            total_original = sum(dept_total_hours.values())
            total_allocated = result_df['allocated_hours'].sum()
            allocation_rate = (total_allocated / total_original * 100) if total_original > 0 else 0
            
            self.textBrowser_4.append(f"\n总体分配统计:")
            self.textBrowser_4.append(f"原始总工时: {total_original:.2f}")
            self.textBrowser_4.append(f"已分配工时: {total_allocated:.2f}")
            self.textBrowser_4.append(f"分配率: {allocation_rate:.2f}%")
            
            app.processEvents()
        else:
            self.textBrowser_4.append("请先完成部门工时计算")
            app.processEvents()

    def clear_hour_gui(self):
        self.lineEdit_30.clear()
        self.lineEdit_37.clear()
        self.lineEdit_38.clear()
        self.lineEdit_31.clear()
        self.textBrowser_4.clear()

    def open_file(self, path):
        os.startfile(path)


    def hour_Operate(self):
        """
        处理工时数据并进行SAP操作，流程为：登录成功→录入hour成功→保存。
        只有上一步成功才执行下一步，失败则跳过后续步骤。
        保留log文件逻辑，必要信息显示在textBrowser_4。
        """
        time_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        log_file = os.path.join(configContent['Hour_Files_Export_URL'], f'log_{time_str}.xlsx')
        columns = [
            'ID', 'staff_id', 'week', 'order_no', 'allocated_hours', 'office_time',
            'material_code', 'item', 'allocated_day', 'staff_name', 'status', 'message'
        ]
        log_obj = Logger(log_file=log_file, columns=columns)
        try:
            hour_path = self.lineEdit_31.text()
            if not hour_path:
                QMessageBox.warning(self, "警告", "请先选择工时文件！")
                return

            get_data = Get_Data()
            raw_data = get_data.getFileTableData(hour_path)
            renamed_data = get_data.rename_hour_fields(raw_data, configContent['Hour_Field_Mapping'])
            sap = Sap()

            for idx, row in renamed_data.iterrows():
                log_data = {
                    'ID': row.get('ID', idx+1),
                    'staff_id': row.get('staff_id', ''),
                    'week': row.get('week', ''),
                    'order_no': row.get('order_no', ''),
                    'allocated_hours': row.get('allocated_hours', ''),
                    'office_time': row.get('office_time', ''),
                    'material_code': row.get('material_code', ''),
                    'item': row.get('item', ''),
                    'allocated_day': row.get('allocated_day', ''),
                    'staff_name': row.get('staff_name', ''),
                    'status': '',
                    'message': ''
                }

                # 工时不能为0
                if row.get('allocated_hours', '') == '' or row.get('allocated_hours', '') == 0:
                    msg = f"SAP数据有问题！Item ID: {row.get('ID', idx + 1)}；错误信息：工时不能为0"
                    log_data['status'] = 'Failed'
                    log_data['message'] = msg
                    self.textBrowser_4.append(f"<font color='red'>{msg}</font>")
                    log_obj.log(log_data)
                    app.processEvents()
                    continue

                # 1. 登录
                login_res = sap.login_hour_gui(row)
                if not login_res.get('flag', False):
                    msg = f"登录SAP失败！Item ID: {row.get('ID', idx+1)}；错误信息：{login_res.get('msg', '未知错误')}"
                    log_data['status'] = 'Failed'
                    log_data['message'] = msg
                    self.textBrowser_4.append(f"<font color='red'>{msg}</font>")
                    log_obj.log(log_data)
                    app.processEvents()
                    continue

                # 2. 录入hour（仅在登录成功时执行）
                recording_res = sap.recording_hours(row)
                if not recording_res.get('flag', False):
                    msg = f"记录工时失败！Item ID: {row.get('ID', idx+1)}；错误信息：{recording_res.get('msg', '未知错误')}"
                    log_data['status'] = 'Failed'
                    log_data['message'] = msg
                    self.textBrowser_4.append(f"<font color='red'>{msg}</font>")
                    log_obj.log(log_data)
                    app.processEvents()
                    continue

                # 3. 保存（仅在录入hour成功时执行）
                save_res = sap.save_hours()
                if not save_res.get('flag', False):
                    msg = f"保存工时失败！Item ID: {row.get('ID', idx+1)}；错误信息：{save_res.get('msg', '未知错误')}"
                    log_data['status'] = 'Failed'
                    log_data['message'] = msg
                    self.textBrowser_4.append(f"<font color='red'>{msg}</font>")
                    log_obj.log(log_data)
                    app.processEvents()
                    continue

                # 成功
                msg = f"成功处理 Item ID: {row.get('ID', idx+1)} 的工时数据"
                log_data['status'] = 'Success'
                log_data['message'] = msg
                self.textBrowser_4.append(msg)
                log_obj.log(log_data)
                app.processEvents()

            log_obj.save_log_to_excel()
            self.textBrowser_4.append(f"所有工时数据处理完成！日志文件保存在：{log_file}")
            os.startfile(log_file)
            app.processEvents()
        except Exception as e:
            log_obj.save_log_to_excel()
            self.textBrowser_4.append(f"错误：处理过程中出现错误: {str(e)}\n日志文件保存在：{log_file}")
            os.startfile(log_file)


if __name__ == "__main__":
    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling)
    app = QApplication(sys.argv)
    myWin = MyMainWindow()
    myTable = MyTableWindow()
    myWin.show()
    myWin.getConfig()
    sys.exit(app.exec_())
