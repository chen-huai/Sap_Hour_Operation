# -*- coding: utf-8 -*-
"""
通用代码签名工具模块
使用方法：
    from signing_tool import SigningTool

    tool = SigningTool("signature_config.json")
    success, message = tool.sign_file("app.exe", "certificate_name")

    # 或使用默认证书
    success, message = tool.sign_file("app.exe")
"""
import os
import sys
import subprocess
import time
import json
import glob
import hashlib
from pathlib import Path
from code_signer.utils import safe_subprocess_run
from code_signer.config_loader import load_signing_config, get_config_load_info
from typing import Dict, List, Tuple, Optional, Any


class SigningTool:
    """通用代码签名工具类"""

    def __init__(self, config_path: str = "signature_config.json"):
        """
        初始化签名工具
        :param config_path: 配置文件路径
        """
        self.config_path = config_path
        self.signing_records = []

        # 使用新的配置加载器
        try:
            self.config_obj = load_signing_config(config_path)
            load_info = get_config_load_info()
            print(f"[信息] 使用配置源: {load_info['load_source']}")

            # 验证配置
            errors = self.config_obj.validate()
            if errors:
                print("[警告] 配置验证发现问题:")
                for error in errors:
                    print(f"  - {error}")
            else:
                print("[信息] 配置验证通过")

        except Exception as e:
            print(f"[错误] 初始化签名工具失败: {e}")
            print("请确保 code_signer 模块和配置文件存在且格式正确")
            raise

    def get_config(self, key_path: str, default=None) -> Any:
        """
        获取配置值（兼容旧接口）
        :param key_path: 配置路径，如 'signature.certificates.default'
        :param default: 默认值
        :return: 配置值
        """
        try:
            # 支持新的配置对象访问方式
            if hasattr(self.config_obj, key_path):
                return getattr(self.config_obj, key_path)

            # 支持嵌套属性访问
            keys = key_path.split('.')
            value = self.config_obj

            for key in keys:
                if hasattr(value, key):
                    value = getattr(value, key)
                elif hasattr(value, '__dict__') and key in value.__dict__:
                    value = value.__dict__[key]
                else:
                    return default

            return value

        except Exception:
            return default

    
    def find_signing_tool(self, tool_name: str = None) -> Optional[str]:
        """
        查找签名工具
        :param tool_name: 工具名称，如果为None则按优先级查找第一个可用的
        :return: 工具路径，如果未找到返回None
        """
        # 使用新的配置对象
        tools = self.config_obj.signing_tools

        if tool_name:
            if tool_name not in tools or not tools[tool_name].enabled:
                return None
            tool_configs = {tool_name: tools[tool_name]}
        else:
            # 按优先级排序
            tool_configs = dict(sorted(
                tools.items(),
                key=lambda x: x[1].priority
            ))

        for name, config in tool_configs.items():
            if not config.enabled:
                continue

            if name == 'signtool':
                path = self._find_signtool(config.path)
            elif name == 'powershell':
                path = 'powershell'  # PowerShell是内置的
            elif name == 'osslsigncode':
                path = self._find_osslsigncode(config.path)
            else:
                continue

            if path:
                return path

        return None

    def _find_signtool(self, path_config: str) -> Optional[str]:
        """查找signtool.exe"""
        if path_config != 'auto':
            return path_config if os.path.exists(path_config) else None

        try:
            search_paths = [
                r"C:\Program Files (x86)\Windows Kits\10\bin\*\x64\signtool.exe",
                r"C:\Program Files\Windows Kits\10\bin\*\x64\signtool.exe"
            ]

            for pattern in search_paths:
                matches = glob.glob(pattern)
                if matches:
                    return matches[-1]  # 使用最新版本
        except Exception:
            pass

        return None

    def _find_osslsigncode(self, path_config: str) -> Optional[str]:
        """查找osslsigncode"""
        if path_config != 'auto':
            return path_config if os.path.exists(path_config) else None

        # 在PATH中查找
        try:
            result = subprocess.run(['where', 'osslsigncode'],
                                  capture_output=True, text=True)
            if result.returncode == 0:
                return result.stdout.strip().split('\n')[0]
        except Exception:
            pass

        return None

    def get_certificate_config(self, cert_name: str) -> Optional[Dict[str, Any]]:
        """
        获取证书配置
        :param cert_name: 证书名称
        :return: 证书配置字典
        """
        certificates = self.get_config('signature.certificates', {})
        return certificates.get(cert_name)

    def verify_certificate_exists(self, cert_config: Dict[str, Any]) -> bool:
        """
        验证证书是否存在
        :param cert_config: 证书配置
        :return: 是否存在
        """
        if not cert_config or not cert_config.get('sha1'):
            return False

        try:
            cmd = ['certutil', '-user', '-store', 'My', cert_config['sha1']]
            result = safe_subprocess_run(cmd, encoding='utf-8')
            return result.returncode == 0
        except Exception:
            return False

    def find_target_files(self, search_dir: str = ".") -> List[str]:
        """
        查找目标文件
        :param search_dir: 搜索目录
        :return: 文件路径列表
        """
        files = []
        search_patterns = self.get_config('signature.file_paths.search_patterns', ['*.exe'])
        exclude_patterns = self.get_config('signature.file_paths.exclude_patterns', [])

        for pattern in search_patterns:
            pattern_path = os.path.join(search_dir, pattern)
            for file_path in glob.glob(pattern_path):
                # 检查排除模式
                if any(glob.fnmatch.fnmatch(os.path.basename(file_path), exclude_pattern)
                      for exclude_pattern in exclude_patterns):
                    continue
                files.append(file_path)

        return files

    def sign_file_with_signtool(self, file_path: str, cert_config: Dict[str, Any]) -> Tuple[bool, str]:
        """使用signtool签名文件"""
        signtool_path = self._find_signtool('auto')
        if not signtool_path:
            return False, "[错误] 未找到signtool.exe"

        timestamp_server = self.get_config('signature.timestamp_server', 'http://timestamp.digicert.com')
        hash_algorithm = self.get_config('signature.hash_algorithm', 'sha256')

        cmd = [
            signtool_path, "sign",
            "/sha1", cert_config['sha1'],
            "/fd", hash_algorithm,
            "/td", hash_algorithm,
            "/tr", timestamp_server,
            file_path
        ]

        try:
            if self.get_config('signature.output.verbose', False):
                print(f"[执行] {' '.join(cmd)}")

            result = safe_subprocess_run(cmd, encoding='utf-8')

            if result.returncode == 0:
                return True, "[成功] 签名成功"
            else:
                return False, f"[错误] 签名失败: {result.stderr}"

        except Exception as e:
            return False, f"[错误] 签名异常: {e}"

    def sign_file_with_powershell(self, file_path: str, cert_config: Dict[str, Any]) -> Tuple[bool, str]:
        """使用PowerShell签名文件"""
        timestamp_server = self.get_config('signature.timestamp_server', 'http://timestamp.digicert.com')

        script = f'''
        Set-AuthenticodeSignature -FilePath "{file_path}" -Certificate "{cert_config['sha1']}" -TimestampServer "{timestamp_server}"
        $LASTEXITCODE
        '''

        try:
            result = safe_subprocess_run(['powershell', '-Command', script], encoding='utf-8')

            if result.returncode == 0:
                return True, "[成功] PowerShell签名成功"
            else:
                return False, f"[错误] PowerShell签名失败: {result.stderr}"

        except Exception as e:
            return False, f"[错误] PowerShell签名异常: {e}"

    def sign_file(self, file_path: str, cert_name: str = None) -> Tuple[bool, str]:
        """
        签名文件
        :param file_path: 文件路径
        :param cert_name: 证书名称，如果为None则使用默认证书
        :return: (是否成功, 消息)
        """
        if not self.get_config('signature.enabled', False):
            return False, "[错误] 签名功能未启用"

        if not os.path.exists(file_path):
            return False, f"[错误] 文件不存在: {file_path}"

        # 获取证书配置
        if not cert_name:
            cert_name = self.get_config('signature.default_certificate', 'default_template')

        cert_config = self.get_certificate_config(cert_name)
        if not cert_config:
            return False, f"[错误] 未找到证书配置: {cert_name}"

        # 验证证书是否存在
        if self.get_config('signature.policies.verify_before_sign', True):
            if not self.verify_certificate_exists(cert_config):
                return False, f"[错误] 证书不存在或无法访问: {cert_config.get('name', cert_name)}"

        # 验证文件是否已签名
        if self.get_config('signature.policies.verify_before_sign', True):
            if self.verify_signature(file_path)[0]:
                return False, f"[警告] 文件已有签名: {file_path}"

        # 获取签名工具
        tool_path = self.find_signing_tool()
        if not tool_path:
            return False, "[错误] 未找到可用的签名工具"

        # 执行签名
        max_retries = self.get_config('signature.policies.max_retries', 3)
        auto_retry = self.get_config('signature.policies.auto_retry', True)

        for attempt in range(max_retries):
            if tool_path.endswith('signtool.exe'):
                success, message = self.sign_file_with_signtool(file_path, cert_config)
            elif tool_path == 'powershell':
                success, message = self.sign_file_with_powershell(file_path, cert_config)
            else:
                return False, f"[错误] 不支持的签名工具: {tool_path}"

            if success or not auto_retry:
                break

            if attempt < max_retries - 1:
                print(f"[重试] 签名失败，{attempt + 1}秒后重试...")
                time.sleep(attempt + 1)

        # 保存签名记录
        if success and self.get_config('signature.policies.record_signing_history', True):
            self.save_signing_record(file_path, cert_name, success, message)

        return success, message

    def verify_signature(self, file_path: str) -> Tuple[bool, str]:
        """
        验证文件签名
        :param file_path: 文件路径
        :return: (是否成功, 消息)
        """
        signtool_path = self._find_signtool('auto')
        if not signtool_path:
            return False, "[错误] 未找到signtool.exe"

        try:
            cmd = [signtool_path, "verify", "/pa", file_path]
            result = safe_subprocess_run(cmd, encoding='utf-8')

            if result.returncode == 0:
                return True, result.stdout
            else:
                return False, result.stderr

        except Exception as e:
            return False, f"[错误] 验证异常: {e}"

    def save_signing_record(self, file_path: str, cert_name: str, success: bool, message: str):
        """
        保存签名记录
        :param file_path: 文件路径
        :param cert_name: 证书名称
        :param success: 是否成功
        :param message: 消息
        """
        cert_config = self.get_certificate_config(cert_name)

        record = {
            "file_path": file_path,
            "file_hash": self._calculate_file_hash(file_path),
            "certificate_name": cert_name,
            "certificate_sha1": cert_config.get('sha1', '') if cert_config else '',
            "signing_time": time.strftime("%Y-%m-%d %H:%M:%S"),
            "success": success,
            "message": message,
            "tool": "signing_tool.py",
            "config_file": self.config_path
        }

        self.signing_records.append(record)

        # 保存到文件
        if self.get_config('signature.output.save_records', True):
            record_dir = self.get_config('signature.file_paths.record_directory', './signature_records')
            os.makedirs(record_dir, exist_ok=True)

            record_file = os.path.join(record_dir,
                f"{os.path.basename(file_path)}_signing_record.json")

            try:
                with open(record_file, "w", encoding="utf-8") as f:
                    json.dump(record, f, indent=2, ensure_ascii=False)

                if self.get_config('signature.output.verbose', False):
                    print(f"[记录] 签名记录已保存到: {record_file}")
            except Exception as e:
                print(f"[警告] 保存签名记录失败: {e}")

    def _calculate_file_hash(self, file_path: str) -> str:
        """计算文件哈希"""
        try:
            with open(file_path, 'rb') as f:
                return hashlib.sha256(f.read()).hexdigest()
        except Exception:
            return ""

    def display_certificate_info(self, cert_name: str = None):
        """
        显示证书信息
        :param cert_name: 证书名称，如果为None则显示默认证书
        """
        if not cert_name:
            cert_name = self.get_config('signature.default_certificate', 'default_template')

        cert_config = self.get_certificate_config(cert_name)
        if not cert_config:
            print(f"[错误] 未找到证书配置: {cert_name}")
            return

        print("\n" + "="*60)
        print(f"证书信息: {cert_config.get('name', cert_name)}")
        print("="*60)
        print(f"证书名称: {cert_config.get('name', cert_name)}")
        print(f"SHA1哈希: {cert_config.get('sha1', 'N/A')}")
        print(f"使用者: {cert_config.get('subject', 'N/A')}")
        print(f"颁发者: {cert_config.get('issuer', 'N/A')}")
        if cert_config.get('valid_from') and cert_config.get('valid_to'):
            print(f"有效期: {cert_config['valid_from']} - {cert_config['valid_to']}")
        print(f"描述: {cert_config.get('description', 'N/A')}")
        print("="*60)

    def batch_sign(self, search_dir: str = ".", cert_name: str = None) -> Dict[str, Tuple[bool, str]]:
        """
        批量签名文件
        :param search_dir: 搜索目录
        :param cert_name: 证书名称
        :return: 文件路径到结果的映射
        """
        files = self.find_target_files(search_dir)
        if not files:
            return {}

        results = {}
        for file_path in files:
            print(f"\n[处理] 签名文件: {file_path}")
            success, message = self.sign_file(file_path, cert_name)
            results[file_path] = (success, message)

            if success:
                print(f"[成功] {message}")
            else:
                print(f"[失败] {message}")

        return results


# 便捷函数
def sign_file(file_path: str, config_path: str = "signature_config.json",
              cert_name: str = None) -> Tuple[bool, str]:
    """
    便捷的文件签名函数
    :param file_path: 文件路径
    :param config_path: 配置文件路径
    :param cert_name: 证书名称
    :return: (是否成功, 消息)
    """
    tool = SigningTool(config_path)
    return tool.sign_file(file_path, cert_name)


def verify_file_signature(file_path: str) -> Tuple[bool, str]:
    """
    便捷的签名验证函数
    :param file_path: 文件路径
    :return: (是否成功, 消息)
    """
    tool = SigningTool()
    return tool.verify_signature(file_path)