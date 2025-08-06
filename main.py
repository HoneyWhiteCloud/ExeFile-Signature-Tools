import os
import sys
import glob
import time
import shutil
import winreg
from pathlib import Path
from typing import Optional, Tuple, List
from enum import Enum
from dataclasses import dataclass


class Colors(Enum):
    """终端颜色枚举"""
    RED = '\033[1;31m'
    GREEN = '\033[1;32m'
    YELLOW = '\033[1;33m'
    BLUE = '\033[1;34m'
    CYAN = '\033[1;36m'
    BLINK_GREEN = '\033[5;32;40m'
    RESET = '\033[0m'


class SignatureStatus(Enum):
    """签名状态枚举"""
    TRUSTED = "trusted"          # 受信任的签名（微软认证）
    SELF_SIGNED = "self_signed"  # 自签名
    UNSIGNED = "unsigned"        # 未签名
    INVALID = "invalid"          # 无效签名
    UNKNOWN = "unknown"          # 未知状态


class FileFormats(Enum):
    """支持的文件格式枚举"""
    EXE = '.exe'
    DLL = '.dll'
    SYS = '.sys'
    MSI = '.msi'
    CAB = '.cab'
    CAT = '.cat'
    OCX = '.ocx'
    PS1 = '.ps1'
    PSM1 = '.psm1'
    PSD1 = '.psd1'
    JS = '.js'
    VBS = '.vbs'
    WSF = '.wsf'
    
    @classmethod
    def get_all_extensions(cls) -> List[str]:
        """获取所有支持的扩展名"""
        return [fmt.value for fmt in cls]
    
    @classmethod
    def is_supported(cls, file_path: str) -> bool:
        """检查文件是否为支持的格式"""
        return any(file_path.lower().endswith(ext) for ext in cls.get_all_extensions())
    
    @classmethod
    def get_format_description(cls) -> str:
        """获取格式描述字符串"""
        formats = cls.get_all_extensions()
        return f"({', '.join(formats)})"


@dataclass
class SigningConfig:
    """签名配置数据类"""
    name: str
    email: Optional[str] = None
    password: Optional[str] = None
    pfx_path: Optional[str] = None


@dataclass
class SignatureInfo:
    """签名信息数据类"""
    status: SignatureStatus
    signer_name: Optional[str] = None
    issuer: Optional[str] = None
    timestamp: Optional[str] = None
    is_microsoft_signed: bool = False
    error_message: Optional[str] = None


class DigitalSignatureTool:
    """数字签名工具主类"""
    
    VERSION = "v0.0.0.3"
    TITLE = f"数字签名生成/签名程序(非认证) {VERSION}"
    TIMESTAMP_URLS = [
        "http://timestamp.comodoca.com/authenticode",
        "http://timestamp.digicert.com",
        "http://timestamp.sectigo.com",
        "http://tsa.starfieldtech.com"
    ]
    
    def __init__(self):
        self.tools_path = self._get_resource_path("tools")
        self.tools = {
            'cert2spc': self._get_resource_path(os.path.join("tools", "cert2spc.exe")),
            'makecert': self._get_resource_path(os.path.join("tools", "makecert.exe")),
            'pvk2pfx': self._get_resource_path(os.path.join("tools", "pvk2pfx.exe")),
            'signtool': self._get_resource_path(os.path.join("tools", "signtool.exe"))
        }
        self.desktop_path = self._get_desktop_path()
        self.current_timestamp_url = self.TIMESTAMP_URLS[0]
        
    @staticmethod
    def _get_resource_path(relative_path: str) -> str:
        """获取资源路径（支持打包后的路径）"""
        if getattr(sys, 'frozen', False):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)
    
    @staticmethod
    def _get_desktop_path() -> str:
        """获取桌面路径"""
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                             r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders')
        return winreg.QueryValueEx(key, "Desktop")[0]
    
    def _clear_screen(self):
        """清屏"""
        os.system('cls' if os.name == 'nt' else 'clear')
    
    def _print_colored(self, text: str, color: Colors):
        """打印带颜色的文本"""
        print(f"{color.value}{text}{Colors.RESET.value}")
    
    def _check_tools(self) -> bool:
        """检查必要的工具文件是否存在"""
        if not os.path.exists(self.tools_path):
            self._print_colored(f"工具文件夹丢失:\nPATH: {{{self.tools_path}}}", Colors.RED)
            return False
            
        for tool_name, tool_path in self.tools.items():
            if not os.path.exists(tool_path):
                self._print_colored(f"{tool_name}.exe文件丢失:\nPATH: {tool_path}", Colors.RED)
                return False
                
        os.chdir(self.tools_path)
        return True
    
    def _execute_command(self, command: str) -> str:
        """执行系统命令并返回输出"""
        return os.popen(command).read()
    
    def _get_user_input(self, prompt: str, allow_empty: bool = True) -> str:
        """获取用户输入"""
        value = input(prompt).strip()
        if not allow_empty and not value:
            self._print_colored("输入不能为空！", Colors.RED)
            return self._get_user_input(prompt, allow_empty)
        return value
    
    def _get_file_path(self, prompt: str, check_format: bool = True) -> str:
        """获取文件路径并验证"""
        while True:
            path = self._get_user_input(prompt).strip('"')
            
            if check_format:
                if FileFormats.is_supported(path):
                    if os.path.exists(path):
                        return path
                    else:
                        self._print_colored(f"文件不存在: {path}", Colors.RED)
                else:
                    self._print_colored(
                        f"不支持的文件格式！\n支持的格式: {FileFormats.get_format_description()}", 
                        Colors.RED
                    )
            else:
                # 用于PFX文件等特定格式
                if os.path.exists(path):
                    return path
                else:
                    self._print_colored(f"文件不存在: {path}", Colors.RED)
                    
            input("回车以重试：")
            self._clear_screen()
    
    def _get_multiple_file_paths(self, prompt: str) -> List[str]:
        """获取多个文件路径（批量处理）"""
        self._clear_screen()
        print("批量处理模式")
        print("请逐个输入文件路径，输入空行结束")
        print(f"支持的格式: {FileFormats.get_format_description()}\n")
        
        files = []
        count = 1
        while True:
            file_path = input(f"文件 {count} (直接回车结束): ").strip().strip('"')
            if not file_path:
                break
                
            if FileFormats.is_supported(file_path) and os.path.exists(file_path):
                files.append(file_path)
                count += 1
                self._print_colored(f"✓ 已添加: {os.path.basename(file_path)}", Colors.GREEN)
            else:
                self._print_colored("✗ 文件无效或不支持的格式", Colors.RED)
                
        return files
    
    def _parse_signature_info(self, verify_output: str) -> SignatureInfo:
        """解析签名验证输出，获取详细信息"""
        info = SignatureInfo(status=SignatureStatus.UNKNOWN)
        
        # 检查是否未签名
        if "SignTool Error: No signature found" in verify_output or "未找到签名" in verify_output:
            info.status = SignatureStatus.UNSIGNED
            return info
        
        # 检查是否为自签名证书（根证书不受信任）
        if "terminated in a root certificate which is not trusted" in verify_output:
            info.status = SignatureStatus.SELF_SIGNED
            # 不将此作为错误
        elif "SignTool Error" in verify_output and "Number of errors" not in verify_output:
            # 其他类型的SignTool错误
            info.status = SignatureStatus.INVALID
            info.error_message = "签名验证失败"
        
        # 提取签名者信息
        for line in verify_output.split('\n'):
            line = line.strip()
            if "Issued to:" in line or "颁发给:" in line:
                info.signer_name = line.split(':', 1)[1].strip()
            elif "Issued by:" in line or "颁发者:" in line:
                info.issuer = line.split(':', 1)[1].strip()
            elif "Timestamp:" in line or "时间戳:" in line or "Timestamped:" in line:
                info.timestamp = line.split(':', 1)[1].strip()
            elif "Signing Certificate Chain:" in line:
                # 开始解析证书链
                continue
        
        # 如果已经确定是自签名，直接返回
        if info.status == SignatureStatus.SELF_SIGNED:
            return info
        
        # 判断是否为微软认证的签名
        microsoft_keywords = [
            "Microsoft Corporation",
            "Microsoft Code Signing",
            "Microsoft Windows",
            "Windows Verified Publisher",
            "DigiCert",
            "VeriSign",
            "Symantec",
            "GlobalSign",
            "Sectigo",
            "Comodo",
            "Thawte",
            "GeoTrust"
        ]
        
        # 检查是否为受信任的证书颁发机构
        if info.issuer:
            for keyword in microsoft_keywords:
                if keyword.lower() in info.issuer.lower():
                    info.is_microsoft_signed = True
                    break
        
        # 判断最终签名状态
        if info.status == SignatureStatus.UNKNOWN:
            if "Successfully verified" in verify_output:
                # 签名验证成功
                if info.is_microsoft_signed:
                    info.status = SignatureStatus.TRUSTED
                else:
                    info.status = SignatureStatus.TRUSTED
            elif "Number of files successfully Verified: 1" in verify_output:
                # 另一种成功验证的标志
                if info.is_microsoft_signed:
                    info.status = SignatureStatus.TRUSTED
                else:
                    info.status = SignatureStatus.TRUSTED
            else:
                # 检查是否为自签名
                if info.signer_name and info.issuer:
                    if info.signer_name == info.issuer:
                        info.status = SignatureStatus.SELF_SIGNED
                    elif "Root Certificate" in verify_output or "root certificate" in verify_output:
                        info.status = SignatureStatus.SELF_SIGNED
                    else:
                        info.status = SignatureStatus.SELF_SIGNED
                else:
                    info.status = SignatureStatus.INVALID
                
        return info
    
    def _display_signature_status(self, info: SignatureInfo, file_name: str):
        """显示签名状态信息"""
        print(f"\n文件: {file_name}")
        print("-" * 60)
        
        # 根据状态显示不同颜色的状态信息
        if info.status == SignatureStatus.TRUSTED:
            self._print_colored("✓ 签名状态: 受信任的签名（已通过认证）", Colors.GREEN)
        elif info.status == SignatureStatus.SELF_SIGNED:
            self._print_colored("⚠ 签名状态: 自签名证书（未经认证）", Colors.YELLOW)
        elif info.status == SignatureStatus.UNSIGNED:
            self._print_colored("✗ 签名状态: 未签名", Colors.RED)
        elif info.status == SignatureStatus.INVALID:
            self._print_colored("✗ 签名状态: 签名无效或损坏", Colors.RED)
        else:
            self._print_colored("? 签名状态: 未知", Colors.CYAN)
        
        # 显示详细信息
        if info.signer_name:
            print(f"\n签名者: {info.signer_name}")
        if info.issuer:
            print(f"颁发者: {info.issuer}")
        if info.timestamp:
            print(f"时间戳: {info.timestamp}")
        if info.error_message:
            self._print_colored(f"\n错误信息: {info.error_message}", Colors.RED)
        
        # 显示建议
        if info.status == SignatureStatus.SELF_SIGNED:
            self._print_colored(
                "\n提示: 此文件使用自签名证书。虽然文件已签名，但证书未经过权威机构认证。",
                Colors.YELLOW
            )
        elif info.status == SignatureStatus.UNSIGNED:
            self._print_colored(
                "\n提示: 此文件未进行数字签名。建议对可执行文件进行签名以确保完整性。",
                Colors.RED
            )
    
    def option_verify_signature(self):
        """选项7: 验证文件签名"""
        self._clear_screen()
        print("验证文件的数字签名状态\n")
        print(f"支持的文件格式: {FileFormats.get_format_description()}\n")
        
        # 询问批量还是单个
        mode = self._get_user_input("验证模式:\n[1] 单个文件\n[2] 批量验证\n选择 (默认为1): ") or "1"
        
        if mode == '2':
            files = self._get_multiple_file_paths("批量验证签名")
            if files:
                self._clear_screen()
                print(f"准备验证 {len(files)} 个文件的签名...\n")
                
                # 统计信息
                stats = {
                    SignatureStatus.TRUSTED: 0,
                    SignatureStatus.SELF_SIGNED: 0,
                    SignatureStatus.UNSIGNED: 0,
                    SignatureStatus.INVALID: 0,
                    SignatureStatus.UNKNOWN: 0
                }
                
                for i, file_path in enumerate(files, 1):
                    print(f"\n[{i}/{len(files)}] 正在验证: {os.path.basename(file_path)}")
                    
                    # 使用signtool verify命令，使用 /pa 参数允许任何证书
                    verify_cmd = f'signtool verify /pa /v "{file_path}"'
                    result = self._execute_command(verify_cmd)
                    
                    # 解析签名信息
                    info = self._parse_signature_info(result)
                    stats[info.status] += 1
                    
                    # 显示简要状态
                    if info.status == SignatureStatus.TRUSTED:
                        self._print_colored("✓ 受信任的签名", Colors.GREEN)
                    elif info.status == SignatureStatus.SELF_SIGNED:
                        self._print_colored("⚠ 自签名", Colors.YELLOW)
                    elif info.status == SignatureStatus.UNSIGNED:
                        self._print_colored("✗ 未签名", Colors.RED)
                    elif info.status == SignatureStatus.INVALID:
                        self._print_colored("✗ 签名无效", Colors.RED)
                    else:
                        self._print_colored("? 未知状态", Colors.CYAN)
                
                # 显示统计信息
                print("\n" + "=" * 60)
                print("验证结果统计：")
                if stats[SignatureStatus.TRUSTED] > 0:
                    self._print_colored(f"  受信任的签名: {stats[SignatureStatus.TRUSTED]} 个", Colors.GREEN)
                if stats[SignatureStatus.SELF_SIGNED] > 0:
                    self._print_colored(f"  自签名证书: {stats[SignatureStatus.SELF_SIGNED]} 个", Colors.YELLOW)
                if stats[SignatureStatus.UNSIGNED] > 0:
                    self._print_colored(f"  未签名: {stats[SignatureStatus.UNSIGNED]} 个", Colors.RED)
                if stats[SignatureStatus.INVALID] > 0:
                    self._print_colored(f"  签名无效: {stats[SignatureStatus.INVALID]} 个", Colors.RED)
                if stats[SignatureStatus.UNKNOWN] > 0:
                    self._print_colored(f"  未知状态: {stats[SignatureStatus.UNKNOWN]} 个", Colors.CYAN)
        else:
            file_path = self._get_file_path("请将您要验证签名的文件拖入窗口内：")
            
            self._clear_screen()
            print(f"正在验证: {os.path.basename(file_path)}")
            
            # 使用signtool verify命令（使用 /pa 参数允许任何证书）
            verify_cmd = f'signtool verify /pa /v "{file_path}"'
            result = self._execute_command(verify_cmd)
            
            # 解析并显示签名信息
            info = self._parse_signature_info(result)
            self._display_signature_status(info, os.path.basename(file_path))
            
            # 询问是否显示原始输出
            show_raw = self._get_user_input("\n是否显示详细的原始输出？(y/N): ").lower()
            if show_raw == 'y':
                print("\n" + "=" * 60)
                print("原始输出：")
                print(result)
        
        input("\n回车以返回主界面……")
    
    def _select_timestamp_server(self):
        """选择时间戳服务器"""
        self._clear_screen()
        print("选择时间戳服务器：")
        for i, url in enumerate(self.TIMESTAMP_URLS, 1):
            print(f"[{i}] {url}")
        
        choice = self._get_user_input("\n选择服务器 (默认为1): ") or "1"
        try:
            index = int(choice) - 1
            if 0 <= index < len(self.TIMESTAMP_URLS):
                self.current_timestamp_url = self.TIMESTAMP_URLS[index]
                print(f"已选择: {self.current_timestamp_url}")
                time.sleep(1)
        except:
            pass
    
    def _create_certificate(self, config: SigningConfig) -> bool:
        """创建证书文件"""
        # 构建makecert命令
        cn_part = f'CN={config.name}'
        email_part = f'EMAIL={config.email}' if config.email else ''
        name_spec = f'"{cn_part}+{email_part}"' if email_part else f'"{cn_part}"'
        
        cmd = f'makecert -sv name.pvk -r -n {name_spec} name.cer'
        self._execute_command(cmd)
        
        # 转换为spc格式
        self._execute_command("cert2spc name.cer name.spc")
        return True
    
    def _create_pfx(self, password: Optional[str] = None) -> Tuple[bool, Optional[str]]:
        """创建PFX文件"""
        while True:
            self._clear_screen()
            if password is None:
                password = self._get_user_input("请输入您刚才输入的密钥(如无则空着)：")
            
            if password:
                cmd = f"pvk2pfx -pvk name.pvk -pi {password} -spc name.spc -pfx Key.pfx -f"
            else:
                cmd = "pvk2pfx -pvk name.pvk -spc name.spc -pfx Key.pfx -f"
                
            result = self._execute_command(cmd)
            
            if "ERROR: Password incorrect or PVK file corrupted." in result:
                self._print_colored("密码错误！", Colors.RED)
                input("回车以重试：")
                password = None
            else:
                return True, password
    
    def _cleanup_temp_files(self):
        """清理临时文件"""
        temp_files = glob.glob("name.*")
        for file in temp_files:
            try:
                os.remove(file)
            except:
                pass
    
    def _copy_to_desktop(self, filename: str):
        """复制文件到桌面"""
        desktop_file = os.path.join(self.desktop_path, filename)
        if os.path.exists(desktop_file):
            os.remove(desktop_file)
        shutil.copy(filename, self.desktop_path)
        print(f"{filename}文件已保存至桌面……")
    
    def _sign_file(self, file_path: str, pfx_path: str, password: Optional[str] = None, add_timestamp: bool = True):
        """执行签名操作"""
        file_path = Path(file_path)
        
        # 处理路径中包含空格的情况
        if ' ' in str(file_path):
            original_cwd = os.getcwd()
            os.chdir(file_path.parent)
            temp_name = f"temp_sign{file_path.suffix}"
            os.rename(file_path.name, temp_name)
            
            self._execute_sign_command(str(file_path.parent / temp_name), pfx_path, password, add_timestamp)
            
            os.rename(temp_name, file_path.name)
            os.chdir(original_cwd)
        else:
            self._execute_sign_command(str(file_path), pfx_path, password, add_timestamp)
    
    def _execute_sign_command(self, file_path: str, pfx_path: str, password: Optional[str] = None, add_timestamp: bool = True):
        """执行实际的签名命令"""
        # 签名
        if password:
            sign_cmd = f'signtool sign /f "{pfx_path}" /p {password} "{file_path}"'
        else:
            sign_cmd = f'signtool sign /f "{pfx_path}" "{file_path}"'
            
        result = self._execute_command(sign_cmd)
        
        # 添加时间戳
        if add_timestamp:
            timestamp_cmd = f'signtool timestamp /t {self.current_timestamp_url} "{file_path}"'
            self._execute_command(timestamp_cmd)
        
        return "Successfully" in result or "成功" in result
    
    def option_create_and_sign(self):
        """选项1: 一键生成.pfx文件并签名"""
        self._clear_screen()
        print("本功能可一键生成.pfx文件并为您所要签名的程序签名+添加时间戳")
        print("添加时间戳时需要联网，请关注您的网络状态！\n")
        print(f"支持的文件格式: {FileFormats.get_format_description()}\n")
        self._print_colored("\n.pfx文件将保存至桌面，请知悉", Colors.RED)
        self._print_colored("生成过程中会需要您生成密钥，为了安全起见，建议输入一个可靠性牢固的密码并牢记", Colors.RED)
        self._print_colored("注意：生成的证书为自签名证书，未经权威机构认证", Colors.YELLOW)
        
        # 获取用户信息
        config = SigningConfig(
            name=self._get_user_input("\n您的签名者名称：", allow_empty=False),
            email=self._get_user_input("\n您的电子邮箱地址（可选，如无需则空着）：")
        )
        
        self._clear_screen()
        print("请在下一个窗口创建您的密钥")
        input("回车以继续：")
        
        # 创建证书和PFX
        self._create_certificate(config)
        success, password = self._create_pfx()
        
        if success:
            self._copy_to_desktop("Key.pfx")
            self._cleanup_temp_files()
            
            # 询问批量还是单个
            self._clear_screen()
            mode = self._get_user_input("签名模式:\n[1] 单个文件\n[2] 批量签名\n选择: ")
            
            if mode == '2':
                files = self._get_multiple_file_paths("批量签名模式")
                if files:
                    self._clear_screen()
                    print(f"准备签名 {len(files)} 个文件...")
                    for i, file_path in enumerate(files, 1):
                        print(f"\n[{i}/{len(files)}] 正在签名: {os.path.basename(file_path)}")
                        self._sign_file(file_path, "Key.pfx", password)
                        self._print_colored("✓ 完成", Colors.GREEN)
            else:
                file_path = self._get_file_path(f"请将您所要签名的文件拖入窗口内 {FileFormats.get_format_description()}：")
                self._sign_file(file_path, "Key.pfx", password)
            
            os.remove("Key.pfx")
            self._clear_screen()
            print("程序签名+添加时间戳完成")
            self._print_colored("提示：您使用的是自签名证书，签名后的文件可能会被系统警告", Colors.YELLOW)
            input("回车以返回主界面……")
    
    def option_sign_with_pfx(self):
        """选项2: 使用现有PFX文件签名"""
        self._clear_screen()
        print("本功能可将您的.pfx文件签名您欲签名的文件，并添加时间戳\n")
        print("添加时间戳时需要联网，请关注您的网络状态！\n")
        print(f"支持的文件格式: {FileFormats.get_format_description()}\n")
        
        # 首先获取PFX文件路径，但不使用格式检查
        while True:
            pfx_path = self._get_user_input("您的.pfx文件路径：").strip('"')
            if pfx_path.lower().endswith('.pfx') and os.path.exists(pfx_path):
                break
            else:
                self._print_colored("请确保文件是.pfx格式且存在！", Colors.RED)
                input("回车以重试：")
                self._clear_screen()
        
        password = self._get_user_input("请输入您数字证书的密钥(如无则空着)：")
        
        # 询问批量还是单个
        self._clear_screen()
        mode = self._get_user_input("签名模式:\n[1] 单个文件\n[2] 批量签名\n选择: ")
        
        if mode == '2':
            files = self._get_multiple_file_paths("批量签名模式")
            if files:
                self._clear_screen()
                print(f"准备签名 {len(files)} 个文件...")
                for i, file_path in enumerate(files, 1):
                    print(f"\n[{i}/{len(files)}] 正在签名: {os.path.basename(file_path)}")
                    self._sign_file(file_path, pfx_path, password)
                    self._print_colored("✓ 完成", Colors.GREEN)
        else:
            file_path = self._get_file_path(f"请将您所要签名的文件拖入窗口内 {FileFormats.get_format_description()}：")
            self._sign_file(file_path, pfx_path, password)
        
        self._clear_screen()
        print("程序签名+添加时间戳完成")
        input("\n回车以返回主界面……")
    
    def option_add_timestamp(self):
        """选项3: 仅添加时间戳"""
        self._clear_screen()
        print("添加时间戳时需要联网，请关注您的网络状态！\n")
        print(f"支持的文件格式: {FileFormats.get_format_description()}\n")
        
        # 选择时间戳服务器
        self._select_timestamp_server()
        
        # 询问批量还是单个
        self._clear_screen()
        mode = self._get_user_input("时间戳模式:\n[1] 单个文件\n[2] 批量处理\n选择: ")
        
        if mode == '2':
            files = self._get_multiple_file_paths("批量添加时间戳")
            if files:
                self._clear_screen()
                print(f"准备为 {len(files)} 个文件添加时间戳...")
                for i, file_path in enumerate(files, 1):
                    print(f"\n[{i}/{len(files)}] 正在处理: {os.path.basename(file_path)}")
                    file_path_obj = Path(file_path)
                    
                    if ' ' in str(file_path_obj):
                        original_cwd = os.getcwd()
                        os.chdir(file_path_obj.parent)
                        temp_name = f"temp_timestamp{file_path_obj.suffix}"
                        os.rename(file_path_obj.name, temp_name)
                        
                        timestamp_cmd = f'signtool timestamp /t {self.current_timestamp_url} "{temp_name}"'
                        self._execute_command(timestamp_cmd)
                        
                        os.rename(temp_name, file_path_obj.name)
                        os.chdir(original_cwd)
                    else:
                        timestamp_cmd = f'signtool timestamp /t {self.current_timestamp_url} "{file_path}"'
                        self._execute_command(timestamp_cmd)
                    
                    self._print_colored("✓ 完成", Colors.GREEN)
        else:
            file_path = self._get_file_path(f"请将您所要添加时间戳的文件拖入窗口内 {FileFormats.get_format_description()}：")
            
            file_path_obj = Path(file_path)
            if ' ' in str(file_path_obj):
                original_cwd = os.getcwd()
                os.chdir(file_path_obj.parent)
                temp_name = f"temp_timestamp{file_path_obj.suffix}"
                os.rename(file_path_obj.name, temp_name)
                
                timestamp_cmd = f'signtool timestamp /t {self.current_timestamp_url} "{temp_name}"'
                self._execute_command(timestamp_cmd)
                
                os.rename(temp_name, file_path_obj.name)
                os.chdir(original_cwd)
            else:
                timestamp_cmd = f'signtool timestamp /t {self.current_timestamp_url} "{file_path}"'
                self._execute_command(timestamp_cmd)
        
        self._clear_screen()
        print("程序添加时间戳完成")
        input("\n回车以返回主界面……")
    
    def option_create_pfx_only(self):
        """选项4: 仅生成PFX文件"""
        self._clear_screen()
        
        config = SigningConfig(
            name=self._get_user_input("您的签名者名称：", allow_empty=False),
            email=self._get_user_input("\n您的电子邮箱地址（可选，如无需则空着）：")
        )
        
        self._clear_screen()
        print("请在下一个窗口创建您的密钥")
        input("回车以继续：")
        
        self._create_certificate(config)
        success, _ = self._create_pfx()
        
        if success:
            self._copy_to_desktop("Key.pfx")
            self._cleanup_temp_files()
            os.remove("Key.pfx")
            
            self._clear_screen()
            self._print_colored("注意：生成的证书为自签名证书，未经权威机构认证", Colors.YELLOW)
            input("\n回车以返回主界面……")
    
    def option_create_cer_only(self):
        """选项5: 仅创建CER证书"""
        self._clear_screen()
        
        config = SigningConfig(
            name=self._get_user_input("您的签名者名称：", allow_empty=False),
            email=self._get_user_input("\n您的电子邮箱地址（可选，如无需则空着）：")
        )
        
        self._clear_screen()
        print("请在下一个窗口创建您的密钥")
        input("回车以继续：")
        
        self._create_certificate(config)
        os.rename("name.cer", "Key.cer")
        self._copy_to_desktop("Key.cer")
        self._cleanup_temp_files()
        os.remove("Key.cer")
        
        self._clear_screen()
        self._print_colored("注意：生成的证书为自签名证书，未经权威机构认证", Colors.YELLOW)
        input("\n回车以返回主界面……")
    
    def option_exit(self):
        """选项8: 退出程序"""
        self._cleanup_temp_files()
        self._clear_screen()
        print("临时文件已清除！")
        sys.exit(0)
    
    def show_menu(self):
        """显示主菜单"""
        self._clear_screen()
        self._print_colored(f"{self.TITLE}\n", Colors.BLINK_GREEN)
        self._print_colored("######主界面######\n", Colors.RED)
        
        menu_options = [
            "1、一键生成.pfx文件并为您所要签名的程序签名+添加时间戳",
            "2、将.pfx文件签名到您所要签名的程序+添加时间戳",
            "3、仅添加时间戳",
            "4、仅生成.pfx文件",
            "5、仅创建安全证书 (.cer 文件)",
            "6、选择时间戳服务器",
            "7、验证文件签名",
            "8、清理临时文件并退出"
        ]
        
        self._print_colored("输入各选项前的数字以选择↓", Colors.GREEN)
        for option in menu_options:
            self._print_colored(option, Colors.GREEN)
        
        print(f"\n当前时间戳服务器: {self.current_timestamp_url}")
        print(f"支持的文件格式: {FileFormats.get_format_description()}")
        
        # 显示签名状态说明
        print("\n签名状态说明：")
        self._print_colored("  ✓ 绿色 = 受信任的签名（权威机构认证）", Colors.GREEN)
        self._print_colored("  ⚠ 黄色 = 自签名证书（未经认证）", Colors.YELLOW)
        self._print_colored("  ✗ 红色 = 未签名或签名无效", Colors.RED)
        
        print(Colors.RESET.value)
        
        return self._get_user_input(f'\n{Colors.BLUE.value}[光标]:{Colors.RESET.value}')
    
    def run(self):
        """运行主程序"""
        os.system(f"title {self.TITLE}")
        
        # 检查工具文件
        if not self._check_tools():
            print("\n请您在杀毒软件内信任本软件或关闭杀毒软件！本程序内无任何恶意代码！")
            input("\n回车以退出……")
            sys.exit(1)
        
        # 菜单映射
        menu_map = {
            '1': self.option_create_and_sign,
            '2': self.option_sign_with_pfx,
            '3': self.option_add_timestamp,
            '4': self.option_create_pfx_only,
            '5': self.option_create_cer_only,
            '6': self._select_timestamp_server,
            '7': self.option_verify_signature,
            '8': self.option_exit
        }
        
        # 主循环
        while True:
            try:
                choice = self.show_menu()
                if choice in menu_map:
                    menu_map[choice]()
                else:
                    continue
            except KeyboardInterrupt:
                self.option_exit()
            except Exception as e:
                self._print_colored(f"发生错误: {str(e)}", Colors.RED)
                input("\n回车以继续...")


def main():
    """程序入口"""
    tool = DigitalSignatureTool()
    tool.run()


if __name__ == '__main__':
    main()