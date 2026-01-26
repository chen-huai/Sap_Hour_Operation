# -*- coding: utf-8 -*-
"""
SAP工时操作工具 - 专用数字签名工具
===================================

功能：对已生成的 EXE 文件进行数字签名
使用方法：
    python SAP签名工具.py

证书信息：
- 名称: SAP工时操作工具证书
- SHA1: 144ac4069565211ab67d25a9d6d33af0e18e511e
- 颁发者: TÜV SÜD Certification and Testing (China)
"""

import os
import sys
from pathlib import Path

# 导入签名模块
try:
    from code_signer import CodeSigner
    CODE_SIGNER_AVAILABLE = True
except ImportError:
    CODE_SIGNER_AVAILABLE = False
    print("[警告] code_signer 模块不可用")

try:
    from signing_tool import SigningTool
    SIGNING_TOOL_AVAILABLE = True
except ImportError:
    SIGNING_TOOL_AVAILABLE = False


# ====================== 配置 ======================
PROJECT_CERTIFICATE = "sap_hour_operation"
PROJECT_NAME = "SAP工时操作工具"
CONFIG_FILES = [
    "signature_config.json",
]


def find_exe_file():
    """查找 Sap_Hour_Operate_theme.exe 文件"""
    default_paths = [
        "dist/Sap_Hour_Operate_theme.exe",
        "Sap_Hour_Operate_theme.exe",
    ]

    for path in default_paths:
        if os.path.exists(path):
            return path

    # 如果找不到，询问用户
    print("\n[提示] 未找到 EXE 文件，请输入文件路径：")
    print("例如：dist/Sap_Hour_Operate_theme.exe")
    custom_path = input("路径: ").strip().strip('"')

    if os.path.exists(custom_path):
        return custom_path

    return None


def initialize_signer():
    """初始化签名器"""
    print("\n[初始化] 正在初始化签名系统...")

    # 优先使用 code_signer
    if CODE_SIGNER_AVAILABLE:
        print("  [优先] 尝试使用 code_signer 模块...")

        for config_file in CONFIG_FILES:
            if os.path.exists(config_file):
                try:
                    print(f"    使用配置: {config_file}")
                    signer = CodeSigner.from_config(config_file)
                    print("    ✓ code_signer 初始化成功")
                    return signer, 'code_signer', config_file
                except Exception as e:
                    print(f"    ✗ 配置 {config_file} 加载失败: {e}")
                    continue

        # 使用默认配置
        try:
            print("    使用默认配置")
            signer = CodeSigner()
            print("    ✓ code_signer 初始化成功（默认配置）")
            return signer, 'code_signer', '默认配置'
        except Exception as e:
            print(f"    ✗ code_signer 初始化失败: {e}")

    # 回退到 signing_tool
    if SIGNING_TOOL_AVAILABLE:
        print("  [备选] 尝试使用 signing_tool...")
        for config_file in CONFIG_FILES:
            if os.path.exists(config_file):
                try:
                    print(f"    使用配置: {config_file}")
                    signer = SigningTool(config_file)
                    print("    ✓ signing_tool 初始化成功")
                    return signer, 'signing_tool', config_file
                except Exception as e:
                    print(f"    ✗ 配置 {config_file} 加载失败: {e}")

    print("  ✗ 所有签名模块均不可用")
    return None, None, None


def display_certificate_info(signer, signer_type):
    """显示证书详细信息"""
    print("\n[证书信息]")
    print("-" * 60)

    try:
        if signer_type == 'code_signer':
            # 显示证书信息
            print(f"  证书名称: {PROJECT_CERTIFICATE}")

            # 尝试获取证书配置
            try:
                cert_config = signer.config.get_certificate(PROJECT_CERTIFICATE)
                if cert_config:
                    print(f"  SHA1: {cert_config.sha1}")
                    print(f"  使用者: {cert_config.subject}")
                    print(f"  颁发者: {cert_config.issuer}")
                    if hasattr(cert_config, 'description'):
                        print(f"  描述: {cert_config.description}")
            except:
                print(f"  SHA1: 144ac4069565211ab67d25a9d6d33af0e18e511e")
                print(f"  使用者: SAP工时操作工具")
                print(f"  颁发者: TÜV SÜD Certification and Testing (China)")

        elif signer_type == 'signing_tool':
            cert_config = signer.get_certificate_config(PROJECT_CERTIFICATE)
            if cert_config:
                print(f"  名称: {cert_config.get('name', PROJECT_CERTIFICATE)}")
                print(f"  SHA1: {cert_config.get('sha1', 'N/A')}")
                print(f"  使用者: {cert_config.get('subject', 'N/A')}")
                print(f"  颁发者: {cert_config.get('issuer', 'N/A')}")

    except Exception as e:
        print(f"  获取证书信息失败: {e}")

    print("-" * 60)


def verify_file_signature(signer, file_path, signer_type):
    """验证文件签名"""
    print("\n[验证] 检查文件签名状态...")

    try:
        if signer_type == 'code_signer':
            success, message = signer.verify_signature(file_path)
            return success, message
        elif signer_type == 'signing_tool':
            success, message = signer.verify_signature(file_path)
            return success, message
    except Exception as e:
        return False, f"验证异常: {e}"

    return False, "无法验证"


def sign_file_with_cert(signer, file_path, certificate_name, signer_type):
    """使用指定证书签名文件"""
    print("\n[签名] 开始签名过程...")
    print("-" * 60)

    try:
        if signer_type == 'code_signer':
            success, message = signer.sign_file(file_path, certificate_name)
        elif signer_type == 'signing_tool':
            success, message = signer.sign_file(file_path, certificate_name)
        else:
            return False, "不支持的签名器类型"

        return success, message

    except Exception as e:
        return False, f"签名异常: {e}"


def main():
    """主函数"""
    print("=" * 60)
    print(f"{PROJECT_NAME} - 专用数字签名工具")
    print("=" * 60)

    # 检查签名模块
    if not CODE_SIGNER_AVAILABLE and not SIGNING_TOOL_AVAILABLE:
        print("\n❌ 错误: 没有可用的签名模块")
        print("\n请确保以下文件存在：")
        print("  - code_signer/")
        print("  - signing_tool.py")
        print("  - signature_config.json")
        input("\n按回车退出...")
        return False

    # 初始化签名工具
    signer, signer_type, used_config = initialize_signer()

    if not signer:
        print("\n❌ 错误: 签名工具初始化失败")
        print("\n请检查：")
        print("  1. signature_config.json 文件是否存在")
        print("  2. code_signer 模块是否完整")
        print("  3. 配置文件格式是否正确")
        input("\n按回车退出...")
        return False

    print(f"\n✅ 签名工具初始化成功")
    print(f"   类型: {signer_type}")
    print(f"   配置: {used_config}")

    # 显示证书信息
    display_certificate_info(signer, signer_type)

    # 查找 EXE 文件
    print("\n[检查] 正在查找 Sap_Hour_Operate_theme.exe...")
    exe_file = find_exe_file()

    if not exe_file:
        print("\n❌ 错误: 未找到 EXE 文件")
        print("\n请确保：")
        print("  1. 已运行打包脚本生成 EXE 文件")
        print("  2. EXE 文件在 dist/ 目录或项目根目录")
        print("  3. 文件名为: Sap_Hour_Operate_theme.exe")
        input("\n按回车退出...")
        return False

    file_size = os.path.getsize(exe_file) / (1024 * 1024)
    print(f"\n✅ 找到文件: {exe_file}")
    print(f"   大小: {file_size:.1f} MB")

    # 检查是否已签名
    print("\n" + "=" * 60)
    print("  签名前检查")
    print("=" * 60)

    verify_success, verify_message = verify_file_signature(signer, exe_file, signer_type)

    if verify_success:
        print("  ⚠ 文件已有签名")
        print(f"  信息: {verify_message}")

        choice = input("\n是否重新签名? (y/n): ").strip().lower()
        if choice not in ['y', 'yes', '是']:
            print("\n跳过签名操作")
            input("\n按回车退出...")
            return True
    else:
        print("  ✓ 文件未签名，将进行签名")

    # 执行签名
    print("\n" + "=" * 60)
    print("  执行签名")
    print("=" * 60)

    success, message = sign_file_with_cert(
        signer,
        exe_file,
        PROJECT_CERTIFICATE,
        signer_type
    )

    if success:
        print(f"\n✅ 签名成功!")
        print(f"   消息: {message}")

        # 验证签名
        print("\n" + "=" * 60)
        print("  验证签名")
        print("=" * 60)

        verify_success, verify_message = verify_file_signature(signer, exe_file, signer_type)

        if verify_success:
            print("  ✅ 签名验证成功!")
            print("\n  签名详情:")
            if isinstance(verify_message, str):
                for line in verify_message.split('\n'):
                    if line.strip():
                        print(f"    {line}")
            else:
                print(f"    {verify_message}")
        else:
            print(f"  ⚠ 签名验证失败: {verify_message}")

        print("\n" + "=" * 60)
        print("  签名完成")
        print("=" * 60)
        print(f"\n  文件: {exe_file}")
        print(f"  证书: {PROJECT_CERTIFICATE}")
        print(f"  大小: {os.path.getsize(exe_file) / (1024*1024):.1f} MB")

        # 询问是否打开文件位置
        try:
            choice = input("\n是否打开文件所在目录? (y/n): ").strip().lower()
            if choice in ['y', 'yes', '是']:
                import subprocess
                subprocess.Popen(['explorer', '/select,', exe_file])
        except:
            pass

    else:
        print(f"\n❌ 签名失败: {message}")
        print("\n可能的解决方案:")
        print("  1. 确保 Windows SDK 已安装（包含 signtool.exe）")
        print("  2. 确保证书已正确安装")
        print("  3. 检查网络连接（时间戳服务器）")
        print("  4. 确认文件未被占用")
        print("  5. 检查 signature_config.json 配置")

    input("\n按回车退出...")
    return success


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n⚠ 用户取消操作")
        input("\n按回车退出...")
    except Exception as e:
        print(f"\n\n❌ 发生错误: {e}")
        import traceback
        traceback.print_exc()
        input("\n按回车退出...")
