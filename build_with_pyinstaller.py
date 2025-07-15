import os
import subprocess

entry_file = 'Sap_Operate_theme.py'

if not os.path.exists(entry_file):
    print(f'未找到主程序入口文件: {entry_file}，请确认文件存在。')
    exit(1)

# 构建pyinstaller命令
cmd = [
    'pyinstaller',
    '--onefile',
    '--windowed',
    '--clean',
    '--noconfirm',
    '--icon=Sap_Operate_Logo.ico',
    entry_file
]

print(f'正在打包 {entry_file} ...')
result = subprocess.run(cmd)
if result.returncode == 0:
    print('打包成功，生成的可执行文件在dist目录下。')
else:
    print('打包失败，请检查错误信息。') 