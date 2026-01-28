# -*- coding: utf-8 -*-
"""
SAP工时操作工具 - 自动打包+数字签名脚本
===========================================

功能：
1. 自动检查环境（文件、依赖）
2. 使用 PyInstaller 打包成单文件 EXE
3. 使用 code_signer 模块自动进行代码签名
4. 验证签名结果

使用方法：
    python build_with_signing.py

要求：
    - 主程序：Sap_Hour_Operate_theme.py
    - 证书配置：signature_config.json
    - 图标文件：ch-4.ico

作者：基于 Temu_PDF_Rename_APP 打包工具改进
版本：2.0 (集成 code_signer 模块)
日期：2026-01-23
"""

import os
import sys
import subprocess
import shutil
import json
import time
import glob
from pathlib import Path

# 尝试导入 code_signer 模块
try:
    sys.path.insert(0, str(Path(__file__).parent))
    from code_signer import CodeSigner
    from code_signer.config_loader import load_signing_config
    CODE_SIGNER_AVAILABLE = True
except ImportError as e:
    CODE_SIGNER_AVAILABLE = False
    print(f"[警告] code_signer 模块不可用: {e}")
    print("[信息] 将使用基础签名方式")

# 导入签名工具（兼容层）
try:
    from signing_tool import SigningTool
    SIGNING_TOOL_AVAILABLE = True
except ImportError:
    SIGNING_TOOL_AVAILABLE = False


# ====================== 配置区域 ======================
CONFIG = {
    # 主程序文件
    'main_script': 'Sap_Hour_Operate_theme.py',

    # 证书配置文件
    'signature_config': 'signature_config.json',

    # 证书文件（备用）
    'certificate_file': '170859-code-signing.cer',

    # 图标文件
    'icon_file': 'ch-4.ico',

    # 输出EXE名称
    'exe_name': 'Sap_Hour_Operate_theme',

    # 项目证书名称（在配置中定义）
    'project_certificate': 'sap_hour_operation',

    # 时间戳服务器
    'timestamp_server': 'http://timestamp.digicert.com',

    # 是否保留控制台窗口
    'console': False,
}


# ====================== 工具函数 ======================

def print_header(text):
    """打印标题"""
    print("\n" + "=" * 60)
    print(f"  {text}")
    print("=" * 60)


def print_step(step_num, text):
    """打印步骤"""
    print(f"\n[步骤 {step_num}] {text}")
    print("-" * 60)


def check_files():
    """检查必要文件是否存在"""
    print_step(1, "检查必要文件")

    required_files = [
        CONFIG['main_script'],
        CONFIG['icon_file'],
        CONFIG['signature_config'],
    ]

    missing_files = []
    for file in required_files:
        if os.path.exists(file):
            print(f"  ✓ 找到: {file}")
        else:
            print(f"  ✗ 缺失: {file}")
            missing_files.append(file)

    # 检查证书文件（非强制）
    if os.path.exists(CONFIG['certificate_file']):
        print(f"  ✓ 找到证书: {CONFIG['certificate_file']}")
    else:
        print(f"  ⚠ 警告: 未找到证书文件 {CONFIG['certificate_file']}")

    # 检查 code_signer 模块
    if CODE_SIGNER_AVAILABLE:
        print(f"  ✓ code_signer 模块可用")
    else:
        print(f"  ⚠ code_signer 模块不可用，将使用基础签名")

    if missing_files:
        print(f"\n❌ 错误: 缺少必要文件: {', '.join(missing_files)}")
        return False

    print("\n✅ 文件检查通过")
    return True


def check_dependencies():
    """检查Python依赖包"""
    print_step(2, "检查Python依赖")

    required_modules = {
        'PyQt5': 'PyQt5',
        'pandas': 'pandas',
        'numpy': 'numpy',
        'win32com': 'pywin32',
    }

    missing = []
    for module_name, package_name in required_modules.items():
        try:
            __import__(module_name)
            print(f"  ✓ {module_name}")
        except ImportError:
            print(f"  ✗ {module_name} (缺失)")
            missing.append(package_name)

    if missing:
        print(f"\n⚠ 警告: 缺少依赖包: {', '.join(missing)}")
        choice = input("是否立即安装? (y/n): ").strip().lower()
        if choice in ['y', 'yes', '是']:
            try:
                subprocess.run(
                    [sys.executable, "-m", "pip", "install"] + missing,
                    check=True
                )
                print("✅ 依赖安装成功")
            except subprocess.CalledProcessError as e:
                print(f"❌ 依赖安装失败: {e}")
                return False
        else:
            print("⚠ 继续打包，但可能因缺少依赖导致失败")
            return False

    print("\n✅ 依赖检查通过")
    return True


def clean_build_artifacts():
    """清理旧的打包文件"""
    print_step(3, "清理旧的打包文件")

    dirs_to_remove = ['build', 'dist']
    removed = []

    for dir_name in dirs_to_remove:
        if os.path.exists(dir_name):
            try:
                shutil.rmtree(dir_name)
                print(f"  ✓ 删除: {dir_name}/")
                removed.append(dir_name)
            except Exception as e:
                print(f"  ✗ 删除失败 {dir_name}: {e}")

    # 清理旧的spec文件
    old_specs = glob.glob(f"{CONFIG['exe_name']}.spec")
    for spec_file in old_specs:
        try:
            os.remove(spec_file)
            print(f"  ✓ 删除: {spec_file}")
            removed.append(spec_file)
        except Exception as e:
            print(f"  ✗ 删除失败 {spec_file}: {e}")

    if removed:
        print(f"\n✅ 已清理: {len(removed)} 项")
    else:
        print("\n✅ 无需清理")


def build_exe():
    """使用PyInstaller打包EXE"""
    print_step(4, "开始打包")

    # 构建打包命令
    cmd = [
        sys.executable, "-m", "PyInstaller",
        '--onefile',           # 单文件
        '--windowed' if not CONFIG['console'] else '--console',  # 窗口模式
        '--clean',             # 清理缓存
        '--noconfirm',         # 覆盖确认
        f"--icon={CONFIG['icon_file']}",

        # 收集所有pandas和numpy依赖
        '--collect-all', 'pandas',
        '--collect-all', 'numpy',
        '--copy-metadata', 'pandas',
        '--copy-metadata', 'numpy',

        # 隐藏导入
        '--hidden-import', 'PyQt5.QtCore',
        '--hidden-import', 'PyQt5.QtGui',
        '--hidden-import', 'PyQt5.QtWidgets',
        '--hidden-import', 'pandas',
        '--hidden-import', 'pandas._libs',
        '--hidden-import', 'pandas._libs.tslibs',
        '--hidden-import', 'numpy',
        '--hidden-import', 'numpy.core',
        '--hidden-import', 'win32com',
        '--hidden-import', 'win32com.client',
        '--hidden-import', 'qt_material',

        CONFIG['main_script']
    ]

    print(f"  执行命令: PyInstaller {CONFIG['main_script']}")
    print(f"  配置: 单文件模式, 无控制台, 收集所有依赖")

    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True
            # 移除 encoding='utf-8' 以避免 Windows 系统工具输出 GBK 编码导致的解码错误
        )

        if result.returncode == 0:
            exe_path = f"dist/{CONFIG['exe_name']}.exe"
            if os.path.exists(exe_path):
                size_mb = os.path.getsize(exe_path) / (1024 * 1024)
                print(f"\n✅ 打包成功!")
                print(f"  文件: {exe_path}")
                print(f"  大小: {size_mb:.1f} MB")
                return True, exe_path
            else:
                print(f"\n❌ 打包完成但找不到EXE文件")
                return False, "找不到生成的EXE文件"
        else:
            print(f"\n❌ 打包失败!")
            if result.stderr:
                error_lines = result.stderr.split('\n')[-20:]
                print("  错误信息:")
                for line in error_lines:
                    if line.strip():
                        print(f"    {line}")
            return False, result.stderr

    except Exception as e:
        print(f"\n❌ 打包异常: {e}")
        return False, str(e)


def initialize_signer():
    """初始化签名工具"""
    print("\n  初始化签名系统...")

    if CODE_SIGNER_AVAILABLE:
        try:
            print("  [优先] 尝试使用 code_signer 模块...")
            config_path = CONFIG['signature_config']

            if os.path.exists(config_path):
                print(f"    使用配置: {config_path}")
                signer = CodeSigner.from_config(config_path)
                print("    ✓ code_signer 初始化成功")
                return signer, 'code_signer'
            else:
                print("    ⚠ 配置文件不存在，使用默认配置")
                signer = CodeSigner()
                print("    ✓ code_signer 初始化成功（默认配置）")
                return signer, 'code_signer'

        except Exception as e:
            print(f"    ✗ code_signer 初始化失败: {e}")
            if SIGNING_TOOL_AVAILABLE:
                print("  [备选] 尝试使用 signing_tool...")
                try:
                    signer = SigningTool(CONFIG['signature_config'])
                    print("    ✓ signing_tool 初始化成功")
                    return signer, 'signing_tool'
                except Exception as e2:
                    print(f"    ✗ signing_tool 初始化失败: {e2}")
            return None, None

    elif SIGNING_TOOL_AVAILABLE:
        try:
            print("  [备选] 使用 signing_tool...")
            signer = SigningTool(CONFIG['signature_config'])
            print("  ✓ signing_tool 初始化成功")
            return signer, 'signing_tool'
        except Exception as e:
            print(f"  ✗ signing_tool 初始化失败: {e}")
            return None, None

    else:
        print("  ✗ 所有签名模块均不可用")
        return None, None


def sign_exe_file_with_module(exe_path, signer, signer_type):
    """使用签名模块对EXE文件进行数字签名"""
    print_step(5, "开始数字签名（使用签名模块）")

    print(f"  签名工具: {signer_type}")
    print(f"  证书名称: {CONFIG['project_certificate']}")
    print(f"  目标文件: {exe_path}")

    try:
        # 显示证书信息
        print("\n  证书信息:")
        if signer_type == 'code_signer':
            try:
                signer.display_certificate_info(CONFIG['project_certificate'])
            except Exception as e:
                print(f"    ⚠ 无法显示证书信息: {e}")

        # 执行签名
        print("\n  执行签名...")
        success, message = signer.sign_file(exe_path, CONFIG['project_certificate'])

        if success:
            print(f"  ✓ 签名成功!")
            print(f"    消息: {message}")
            return True, f"{signer_type}签名: {message}"
        else:
            print(f"  ✗ 签名失败: {message}")
            return False, f"{signer_type}签名失败: {message}"

    except Exception as e:
        error_msg = f"签名异常: {e}"
        print(f"  ✗ {error_msg}")
        return False, error_msg


def sign_exe_file_fallback(exe_path):
    """备用签名方法：使用signtool或PowerShell"""
    print_step(5, "开始数字签名（备用方法）")

    print(f"  目标文件: {exe_path}")
    certificate_path = CONFIG['certificate_file']

    if not os.path.exists(certificate_path):
        print(f"  ✗ 证书文件不存在: {certificate_path}")
        return False, "证书文件不存在"

    # 方法1: signtool
    print("\n  [方法1] 尝试使用 signtool.exe...")
    signtool_paths = glob.glob(r"C:\Program Files*\Windows Kits\10\bin\*\x64\signtool.exe")

    if signtool_paths:
        signtool_path = signtool_paths[0]
        print(f"    找到: {signtool_path}")

        try:
            cmd = [
                signtool_path, "sign",
                "/f", certificate_path,
                "/fd", "SHA256",
                "/td", "SHA256",
                "/tr", CONFIG['timestamp_server'],
                exe_path
            ]

            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True
                # 移除 encoding='utf-8' 以避免 Windows 系统工具输出 GBK 编码导致的解码错误
            )

            if result.returncode == 0:
                print("  ✓ signtool 签名成功!")
                return True, "signtool签名成功"
            else:
                print(f"  ✗ signtool 签名失败")
                if result.stderr:
                    print(f"    错误: {result.stderr.strip()}")

        except Exception as e:
            print(f"  ✗ signtool 异常: {e}")
    else:
        print("  ✗ 未找到 signtool.exe")

    # 方法2: PowerShell
    print("\n  [方法2] 尝试使用 PowerShell...")
    ps_script = f"""
try {{
    $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2('{certificate_path}')
    $file = Get-Content -Path '{exe_path}' -Raw -Encoding Byte
    $contentInfo = New-Object System.Security.Cryptography.Pkcs.ContentInfo $file
    $signedCms = New-Object System.Security.Cryptography.Pkcs.SignedCms $contentInfo
    $signerInfo = New-Object System.Security.Cryptography.Pkcs.CmsSigner $cert
    $signedCms.ComputeSignature($signerInfo)
    $signedBytes = $signedCms.Encode()
    [System.IO.File]::WriteAllBytes('{exe_path}', $signedBytes)
    Write-Host "PowerShell签名成功"
    exit 0
}} catch {{
    Write-Host "PowerShell签名失败: $($_.Exception.Message)"
    exit 1
}}
"""

    try:
        ps_file = "temp_sign.ps1"
        with open(ps_file, "w", encoding="utf-8") as f:
            f.write(ps_script)

        result = subprocess.run(
            ["powershell", "-ExecutionPolicy", "Bypass", "-File", ps_file],
            capture_output=True,
            text=True
            # 移除 encoding='utf-8' 以避免 PowerShell 输出 GBK 编码导致的解码错误
        )

        try:
            os.remove(ps_file)
        except:
            pass

        if result.returncode == 0:
            print("  ✓ PowerShell 签名成功!")
            return True, "PowerShell签名成功"
        else:
            print(f"  ✗ PowerShell 签名失败")
            if result.stderr:
                print(f"    错误: {result.stderr.strip()}")

    except Exception as e:
        print(f"  ✗ PowerShell 异常: {e}")

    return False, "所有签名方法均失败"


def verify_signature(exe_path):
    """验证EXE文件的数字签名"""
    print_step(6, "验证数字签名")

    try:
        # 使用 PowerShell 验证签名
        ps_command = f"Get-AuthenticodeSignature '{exe_path}' | Select-Object Status, SignerCertificate | Format-List"

        result = subprocess.run(
            ["powershell", "-Command", ps_command],
            capture_output=True,
            text=True
            # 移除 encoding='utf-8' 以避免 PowerShell 输出 GBK 编码导致的解码错误
        )

        output = result.stdout.strip()
        if output:
            print("  签名信息:")
            for line in output.split('\n'):
                if line.strip():
                    print(f"    {line}")

            if "Valid" in output or "有效" in output:
                print("\n✅ 数字签名验证通过")
                return True
            else:
                print("\n⚠ 数字签名状态未知")
                return False
        else:
            print("  ⚠ 无法获取签名信息")
            return False

    except Exception as e:
        print(f"  ✗ 验证异常: {e}")
        return False


def main():
    """主函数"""
    print_header("SAP工时操作工具 - 自动打包+数字签名 v2.0")
    print(f"配置:")
    print(f"  主程序: {CONFIG['main_script']}")
    print(f"  证书配置: {CONFIG['signature_config']}")
    print(f"  证书文件: {CONFIG['certificate_file']}")
    print(f"  图标: {CONFIG['icon_file']}")
    print(f"  控制台: {'是' if CONFIG['console'] else '否'}")

    if CODE_SIGNER_AVAILABLE:
        print(f"  签名系统: code_signer 模块")
    elif SIGNING_TOOL_AVAILABLE:
        print(f"  签名系统: signing_tool 模块")
    else:
        print(f"  签名系统: 基础签名方式")

    # 记录开始时间
    start_time = time.time()

    try:
        # 步骤1: 检查文件
        if not check_files():
            input("\n按回车退出...")
            return False

        # 步骤2: 检查依赖
        if not check_dependencies():
            input("\n按回车退出...")
            return False

        # 步骤3: 清理旧文件
        clean_build_artifacts()

        # 步骤4: 打包
        success, result = build_exe()
        if not success:
            print(f"\n❌ 打包失败: {result}")
            input("\n按回车退出...")
            return False

        exe_path = result

        # 步骤5: 签名
        print("\n" + "=" * 60)
        print("  数字签名")
        print("=" * 60)

        # 初始化签名工具
        signer, signer_type = initialize_signer()

        if signer:
            # 使用签名模块
            sign_success, sign_message = sign_exe_file_with_module(exe_path, signer, signer_type)
        else:
            # 使用备用方法
            sign_success, sign_message = sign_exe_file_fallback(exe_path)

        # 步骤6: 验证签名（如果签名成功）
        if sign_success:
            verify_signature(exe_path)

        # 计算耗时
        elapsed_time = time.time() - start_time

        # 显示完成信息
        print_header("✅ 打包完成!")
        print(f"\n生成的文件:")
        print(f"  1. {exe_path}")
        print(f"     大小: {os.path.getsize(exe_path) / (1024*1024):.1f} MB")
        print(f"     签名: {'✓ 已签名' if sign_success else '✗ 未签名'}")

        if sign_success:
            print(f"\n  签名信息:")
            print(f"     状态: {sign_message}")
            print(f"     时间: {time.strftime('%Y-%m-%d %H:%M:%S')}")
            print(f"     证书: {CONFIG['project_certificate']}")
        else:
            print(f"\n  签名失败:")
            print(f"     原因: {sign_message}")
            print(f"     提示: 请检查签名配置和证书")

        print(f"\n  耗时: {elapsed_time:.1f} 秒")
        print("=" * 60)

        # 询问是否打开文件夹
        try:
            choice = input("\n是否打开 dist 目录? (y/n): ").strip().lower()
            if choice in ['y', 'yes', '是']:
                os.startfile(os.path.dirname(os.path.abspath(exe_path)))
        except:
            pass

        input("\n按回车退出...")
        return True

    except KeyboardInterrupt:
        print("\n\n⚠ 用户取消操作")
        input("\n按回车退出...")
        return False
    except Exception as e:
        print(f"\n\n❌ 发生错误: {e}")
        import traceback
        traceback.print_exc()
        input("\n按回车退出...")
        return False


if __name__ == "__main__":
    main()
