def hourOperate(self):
    """
    处理工时数据并进行SAP操作，流程为：登录、录入hour、保存。
    保留log文件逻辑，必要信息显示在textBrowser_4。
    """
    time_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = os.path.join(configContent['Hour_Files_Export_URL'], f'log_{time_str}.xlsx')
    columns = [
        'ID', 'staff_id', 'week', 'order_no', 'allocated_hours', 'office_time',
        'material_code', 'item', 'allocated_day', 'staff_name', 'status', 'message', 'Update'
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
                'ID': row.get('ID', idx + 1),
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
                'message': '',
                'Update': datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            }
            # 1. 登录
            login_res = sap.login_hour_gui(row)
            if not login_res.get('flag', False):
                msg = f"登录SAP失败！Staff ID: {row.get('staff_id', '')}, Week: {row.get('week', '')}"
                log_data['status'] = 'Failed'
                log_data['message'] = msg
                self.textBrowser_4.append(msg)
                log_obj.log(log_data)
                app.processEvents()
                continue
            # 2. 录入hour
            recording_res = sap.recording_hours(row)
            if not recording_res.get('flag', False):
                msg = f"记录工时失败！Staff ID: {row.get('staff_id', '')}, Week: {row.get('week', '')}"
                log_data['status'] = 'Failed'
                log_data['message'] = msg
                self.textBrowser_4.append(msg)
                log_obj.log(log_data)
                app.processEvents()
                continue
            # 3. 保存
            save_res = sap.save_hours()
            if not save_res.get('flag', False):
                msg = f"保存工时失败！Staff ID: {row.get('staff_id', '')}, Week: {row.get('week', '')}"
                log_data['status'] = 'Failed'
                log_data['message'] = msg
                self.textBrowser_4.append(msg)
                log_obj.log(log_data)
                app.processEvents()
                continue
            # 成功
            msg = f"成功处理 Staff ID: {row.get('staff_id', '')}, Week: {row.get('week', '')} 的工时数据"
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