## 项目说明

### 项目目标
该项目旨在提供一个自动操作SAP创建订单的解决方案，同时还能够获取并处理数据，并实现PDF文件重命名的功能。

### 功能特点
* 数据处理：根据特殊开票信息将原始数据分开，并按要求合并数据。
* SAP操作：实现SAP自动创建订单的功能。
* 数据恢复：能够找回数据，确保数据完整。

### 安装指南
1. 克隆项目：
`git clone https://github.com/chen-huai/Sap_Operation.git`
2. 安装依赖库：
`pip install -r requirements.txt`
3. 运行主程序：
直接运行主程序`Sap_Operate.py`即可。
4. 配置文件：
运行成功后，将在桌面生成一个`config`文件夹，并生成`config_sap.csv`的配置文件。可根据`config_sap.csv`文件设置自己需要的参数。
5. 打包成exe程序：
   * 安装第三方库pyinstaller：
   `pip install pyinstaller`
   * 执行打包命令：
   `pyinstaller -F -w 主程序绝对路径`
   
### 源码结构说明
* `Sap_Operate.py`：主程序，包含基础配置、数据处理、SAP逻辑操作、PDF重命名的逻辑操作。
* `Sap_Operate_Ui.py`：UI界面。
* `Get_Data.py`：数据处理模块，包括数据基础处理。
* `Sap_Function.py`：SAP操作模块，实现SAP基础操作。
* `File_Operate.py`：文件处理模块，用于创建文件夹和获取文件名称。
* `Data_Table.py`：表格处理模块，实现表格基础设置。
* `PDF_Operate.py`：PDF文件处理模块，包含PDF文件读取和保存功能。

### 特别感谢
* 特别感谢JetBrains的支持，提供优秀的开发工具，让我们能够更高效地进行编码工作。

## Project Description

### Project Objective
This project aims to provide a solution for automatically operating SAP to create orders, as well as obtaining and processing data, and implementing PDF file renaming functionality.

### Features
* Processing: Separate original data based on special invoicing information and merge data as required.
* SAP Operations: Implement the functionality to automatically create orders in SAP.
* Data Recovery: Ability to retrieve data to ensure data integrity.

### Installation Guide
1. Clone the project:
`git clone https://github.com/chen-huai/Sap_Operation.git`
2. Install dependencies:
`pip install -r requirements.txt`
3. Run the main program: Simply run the main program Sap_Operate.py.
4. Configuration file: Upon successful execution, a `config` folder will be generated on the desktop, along with a `config_sap.csv` configuration file. You can set your own parameters in the `config_sap.csv` file.
5. Package as an exe program:
   * Install the third-party library pyinstaller:
   `pip install pyinstaller`
   * Execute the packaging command:
   `pyinstaller -F -w absolute_path_to_main_program`
   
### Source Code Structure
* `Sap_Operate.py`: Main program, including basic configuration, data processing, SAP logic operations, and PDF renaming logic operations.
* `Sap_Operate_Ui.py`: UI interface.
* `Get_Data.py`: Data processing module, including basic data processing.
* `Sap_Function.py`: SAP operation module, implementing basic SAP operations.
* `File_Operate.py`: File processing module for creating folders and obtaining file names.
* `Data_Table.py`: Table processing module, implementing basic table settings.
* `PDF_Operate.py`: PDF file processing module, including PDF file reading and saving functionality.

### Special Thanks
* Special thanks to JetBrains for their support, providing excellent development tools that allow us to code more efficiently.