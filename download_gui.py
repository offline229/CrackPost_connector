import imaplib
import babel.numbers
import email
import os
import re
import logging
import time
import configparser
import threading
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from datetime import datetime, timedelta
from pathlib import Path
import email.header
import email.utils
import sys
from PIL import Image, ImageTk  # 需要安装: pip install pillow
import subprocess
import webbrowser
from tkcalendar import Calendar  # 需要安装: pip install tkcalendar

# 登录页「帮助」文案（可直接改这里；留空则用下方默认说明）
LOGIN_EXTRA_HELP = """
"""
LOGIN_HELP_DEFAULT = (
    "1. 感谢您使用[地缝可视化]程序！\n"
    "2. 本程序需要进行以下操作：\n"
    "   - 开启 IMAP 服务（允许本程序登录邮箱并找到地缝相关信件）\n"
    "   - 注册你的 OC 并下载来往信件\n"
    "   - 如有必要，对一些错误识别的信件手动纠正\n"
    "   - 打开可视化观赏\n"
    "   - 你可以在信件库存内对已有的信件进行改动\n"
    "3. 如遇到任何问题，请在企划群内询问。\n"
    "4. 本程序不涉及任何对隐私信息的暴露问题。"
)
import imaplib, ssl, sys

# 神秘bug已解决: 用 ENVELOPE 解析出的 subject 和用 RFC822 头解析出的 subject 不同
# 兼容一天发多封信给同一个人的情况
# 收	2023-09-05	【4o4 n0t fOUnd	yamuko,猫,咪	A*164	.\email\收\20230905_Fw_A_164_来自【4o4_n0t_fOUnd】的信
# 待办：
# 1、附件汇总
# 	统计所有收信
# 	第一步，对所有有明确的发件人 收件人的收信，检索其是否在contenttxt中带有附件 的字样	
# 	若有，则找到
#  2、增加未匹配手动加入功能
# 打包后任务栏/标题栏 icon：需把 icon 目录打进包，或放在 exe 同级；
# 运行时用 window_icon.apply_tk_window_icon（参考 TablePet）。

# python -m venv venv
# venv\Scripts\activate  # Windows
# pip install --upgrade pip
# pip install pyinstaller pillow tkcalendar babel imapclient

# pyinstaller -F -w -i "icon\CrackPost_v1.ico" `
#     --add-data "icon;icon" `
#     --hidden-import babel.numbers `
#     --hidden-import imapclient `
#     --hidden-import PIL.Image `
#     --hidden-import PIL.ImageTk `
#     --hidden-import tkcalendar `
#     --hidden-import window_icon `
#     --onefile --windowed --icon=icon\CrackPost_v1.ico `
#     download_gui.py

# ==================== 日志配置（全局，仅一次）====================
# 兼容 py 脚本和 PyInstaller 打包后的 exe
# 全局只获取一次
if getattr(sys, 'frozen', False):
    script_dir = Path(sys.executable).parent
else:
    script_dir = Path(os.path.dirname(os.path.abspath(__file__)))


def find_app_icons():
    """定位 ico/png：开发态看源码旁；打包后优先 exe 同级，再回退 _MEIPASS。"""
    ico_names = ("CrackPost_v1.ico", "CrackPost.ico")
    png_names = ("CrackPost_v1.png", "CrackPost.png")
    roots = [script_dir / "icon"]
    if getattr(sys, "frozen", False):
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            roots.append(Path(meipass) / "icon")
    ico_path = png_path = None
    for root in roots:
        if ico_path is None:
            for name in ico_names:
                p = root / name
                if p.is_file():
                    ico_path = p
                    break
        if png_path is None:
            for name in png_names:
                p = root / name
                if p.is_file():
                    png_path = p
                    break
        if ico_path and png_path:
            break
    return ico_path, png_path


def apply_app_window_icon(root):
    """开发与打包后统一设置窗口/任务栏图标。"""
    ico_path, png_path = find_app_icons()
    if not ico_path:
        searched = [script_dir / "icon"]
        if getattr(sys, "frozen", False) and getattr(sys, "_MEIPASS", None):
            searched.append(Path(sys._MEIPASS) / "icon")
        print(f"[icon] 未找到 ico（已查: {[str(p) for p in searched]}）")
        return False
    try:
        # 顶层同目录模块；打包时靠 --hidden-import window_icon 打进包
        from window_icon import apply_tk_window_icon
        return apply_tk_window_icon(
            root,
            ico_path,
            png_path=png_path,
        )
    except Exception as e:
        print(f"[icon] apply_tk_window_icon 失败: {e}")
        # 回退：仅 iconbitmap / iconphoto
        try:
            if sys.platform.startswith("win") and ico_path.is_file():
                root.iconbitmap(str(ico_path))
            if png_path and png_path.is_file():
                pil_img = Image.open(str(png_path)).convert("RGBA")
                photo = ImageTk.PhotoImage(pil_img.resize((32, 32), Image.LANCZOS))
                root.iconphoto(True, photo)
                root._app_icon_fallback = photo
            return True
        except Exception as e2:
            print(f"[icon] 回退也失败: {e2}")
            return False

log_dir = script_dir / "log"
log_dir.mkdir(exist_ok=True, parents=True)

import logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_dir / "email_downloader.log", encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

class EmailDownloader:
    def __init__(self, config_file=None):
        """初始化邮件下载器"""
        # 直接用全局的 script_dir
        self.base_dir = script_dir / "email"
        self.base_dir.mkdir(exist_ok=True, parents=True)
        # 日志目录（已在文件顶部初始化，此处仅备份）
        log_dir = script_dir / "log"
        log_dir.mkdir(exist_ok=True, parents=True)
        
        if config_file is None:
            config_file = script_dir / "email_config.ini"
        else:
            config_file = script_dir / config_file
            
        self.config_file = config_file
        self.config = self._load_config(config_file)
        self.email_addr = self.config.get('Credentials', 'email', fallback='')
        self.password = self.config.get('Credentials', 'password', fallback='')
        self.imap_server = self.config.get('Servers', 'imap_server', fallback='')
        self.imap_port = int(self.config.get('Servers', 'imap_port', fallback='993'))
        
        
        # 创建基本目录 - 现在指向系统根目录下的email文件夹
        self.base_dir = script_dir / "email"
        self.base_dir.mkdir(exist_ok=True, parents=True)
        
        # 邮箱连接和其他设置
        self.mail = None
        self.mailbox_mapping = {}
        self.client = None
        self.search_results = []
        self.oc_registerdata = []
        self.oc_aliases = {}  # {主名: [别名, ...]}
        self.last_unmatched = []
        
        # 添加超时设置
        self.timeout = int(self.config.get('Connection', 'timeout', fallback='30'))
        self.batch_size = int(self.config.get('Connection', 'batch_size', fallback='100'))
        self.crackpost_addrs = {
            'crackpost2@126.com',
            'crackpost@126.com',
        }

    def direction_from_addresses(self, sender, recipients=None):
        """根据与 crackpost 中转站的关系判定收/发。发出去=发，收进来=收。"""
        sender_l = (sender or '').lower()
        recipients = recipients or []
        crack = {a.lower() for a in self.crackpost_addrs}
        if sender_l in crack:
            return "收"
        for r in recipients:
            if (r or '').lower() in crack:
                return "发"
        # 回退：相对本邮箱
        me = (self.email_addr or '').lower()
        if sender_l == me:
            return "发"
        return "收"

    def extract_letter_type(self, subject):
        """
        从主题字符串中提取信件类型（如A123、B456、C789、A*123、A+B456等）。
        兼容下划线和空格分隔（如 FW:C_2485、FW:C 2485）
        """
        import re
        if not subject:
            return ""
        
        # 支持 Fw: 或 FW_ 后面的类型（允许下划线和空格）
        m = re.search(
            r'(?i)Fw[:：_\s-]*\s*(A[\s_]*\*?[\s_]*\d+|B[\s_]*\d+|C[\s_]*\d+|A[\s_]*\+[\s_]*B[\s_]*\d+|A[\s_]*\+[\s_]*C[\s_]*\d+)',
            subject
        )
        if m:
            # 去掉类型内部的空格和下划线
            return re.sub(r'[\s_]+', '', m.group(1))
        
        # 兜底：主题开头直接是类型
        m2 = re.match(
            r'(?i)^(A[\s_]*\*?[\s_]*\d+|B[\s_]*\d+|C[\s_]*\d+|A[\s_]*\+[\s_]*B[\s_]*\d+|A[\s_]*\+[\s_]*C[\s_]*\d+)',
            subject
        )
        if m2:
            return re.sub(r'[\s_]+', '', m2.group(1))
        
        # 最后兜底：匹配单字母+数字（如 C_2485 或 C2485）
        m3 = re.match(r'(?i)(A|B|C.*|A\+B.*|A\+C.*)', subject)
        if m3:
            return re.sub(r'[\s_]+', '', m3.group(1))
        
        return ""

    def _safe_filename(self, name):
        """生成安全的文件名"""
        return re.sub(r'[\\/:*?"<>|]', '_', name)

    def _expected_email_dir(self, email_info):
        """根据邮件的发件人、日期和主题，生成本地存储的标准文件夹路径（以便比较是否保存过）"""
        current_email = self.email_addr.lower()
        sender = (email_info.get('sender') or '').strip()
        send_type = "发" if sender.lower() == current_email else "收"

        # 日期部分
        date_folder = "unknown_date"
        date_val = email_info.get('date_obj') or email_info.get('date')
        try:
            if hasattr(date_val, 'strftime'):
                date_folder = date_val.strftime("%Y%m%d")
            elif isinstance(date_val, str) and date_val:
                try:
                    dobj = email.utils.parsedate_to_datetime(date_val)
                    date_folder = dobj.strftime("%Y%m%d")
                except Exception:
                    m = re.search(r'(\d{4})[-/]?(\d{2})[-/]?(\d{2})', date_val)
                    if m:
                        date_folder = f"{m.group(1)}{m.group(2)}{m.group(3)}"
        except Exception:
            pass

        subject = (email_info.get('subject') or '')[:30]
        safe = self._safe_filename(subject)
        email_dir = self.base_dir / send_type / f"{date_folder}_{safe}"
        return email_dir

    def _email_already_downloaded(self, email_info):
        """判断邮件是否已完整下载（仅检查本地文件）
        
        Args:
            email_info: 邮件信息字典
            
        Returns:
            tuple: (是否已下载, 目录路径, 原因)
        """
        # 1. 计算预期存储路径
        email_dir = self._expected_email_dir(email_info)
        
        # 2. 检查目录是否存在
        if not email_dir.exists():
            return False, str(email_dir), "目录不存在"
        
        # 3. 检查 content.txt 是否存在
        content_file = email_dir / "content.txt"
        if not content_file.exists():
            return False, str(email_dir), "content.txt 不存在"
        
        # 4. 检查 content.txt 文件大小
        try:
            file_size = content_file.stat().st_size
            if file_size <= 100:
                return False, str(email_dir), "content.txt 文件过小"
        except Exception as e:
            return False, str(email_dir), f"检查文件大小失败: {e}"
        
        # 5. 检查必要字段
        try:
            with open(content_file, 'r', encoding='utf-8') as f:
                content = f.read()
                required_fields = ["主题:", "发件邮箱:", "收件邮箱:", "日期:"]
                missing_fields = [field for field in required_fields if field not in content]
                if missing_fields:
                    return False, str(email_dir), f"content.txt 缺少字段: {', '.join(missing_fields)}"
        except Exception as e:
            return False, str(email_dir), f"读取 content.txt 失败: {e}"
        
        # 6. 检查附件（如果标记有附件）
        has_attachments = email_info.get('has_attachments', False)
        if has_attachments:
            # 统计非 content.txt 的文件数
            try:
                files_in_dir = [f for f in email_dir.iterdir() if f.is_file() and f.name != 'content.txt']
                if not files_in_dir:
                    return False, str(email_dir), "附件缺失"
            except Exception as e:
                return False, str(email_dir), f"检查附件失败: {e}"
        
        # 7. 所有检查通过
        logger.info(f"邮件已完整下载（本地验证）: {email_dir}")
        return True, str(email_dir), "已完整（本地验证）"

    # 这里的规则负责搜索，只要搜到即可
    def get_default_rule(self, rule_num=1):
        """获取默认规则
        
        Args:
            rule_num: 规则编号，1或2
            
        Returns:
            包含规则的字典
        """
        rule = {}
        section = 'DefaultRules'
        
        if rule_num == 1:
            rule['subject_pattern'] = self.config.get(section, 'rule1_subject_pattern', 
                                                fallback=r'^Fw[:：]?(A\d+|A\*\d+|B\d+|C\d+|A\+B\d+|A\+C\d+|A\+B\*\d+|A\+C\*\d+)')
            rule['from'] = self.config.get(section, 'rule1_from', fallback='crackpost2@126.com,crackpost@126.com')
            rule['to'] = self.config.get(section, 'rule1_to', fallback=self.email_addr)
        else:
            rule['subject_pattern'] = self.config.get(section, 'rule2_subject_pattern', fallback=r' ^(A|B|C.*|A\+B.*|A\+C.*)$')
            rule['from'] = self.config.get(section, 'rule2_from', fallback=self.email_addr)
            rule['to'] = self.config.get(section, 'rule2_to', fallback='crackpost2@126.com,crackpost@126.com')
            rule['body_contains'] = self.config.get(section, 'rule2_body_contains', fallback='发件人=【')
        
        # 替换变量
        for key in rule:
            if isinstance(rule[key], str):
                rule[key] = rule[key].replace('{self_email}', self.email_addr)
        
        return rule

    def decode_mime_header(self, header):
        """更可靠地解码MIME格式的邮件头"""
        if not header or not isinstance(header, str):
            return ""
            
        # 已经是正常文本，直接返回
        if not header.startswith('=?'):
            return header
            
        try:
            # 针对QQ邮箱特殊处理Base64编码
            import base64
            import re
            
            # 查找所有编码段落
            pattern = r'=\?([^?]+)\?([B|Q])\?([^?]*)\?='
            matches = re.findall(pattern, header)
            
            if matches:
                result = ""
                for charset, encoding, encoded_text in matches:
                    if encoding.upper() == 'B':  # Base64
                        decoded_bytes = base64.b64decode(encoded_text)
                        result += decoded_bytes.decode(charset, errors='replace')
                    elif encoding.upper() == 'Q':  # Quoted-printable
                        import quopri
                        decoded_bytes = quopri.decodestring(encoded_text)
                        result += decoded_bytes.decode(charset, errors='replace')
                return result
            else:
                # 使用email.header标准模块解码
                parts = email.header.decode_header(header)
                result = ""
                for part, charset in parts:
                    if isinstance(part, bytes):
                        result += part.decode(charset or 'utf-8', errors='replace')
                    else:
                        result += str(part)
                return result
        except Exception as e:
            logger.error(f"解码标题失败: {header}, 错误: {str(e)}")
            return header

    def _load_config(self, config_file):
        """加载或创建配置文件"""
        config = configparser.ConfigParser()
        
        if not os.path.exists(config_file):
            # 创建默认配置
            config['Credentials'] = {
                'email': '',
                'password': ''
            }
            
            config['Servers'] = {
                'imap_server': '',
                'imap_port': '993',
            }
            
            config['Storage'] = {
                # 与 exe/源码同级的 email，不要落到用户主目录
                'base_directory': str(script_dir / 'email')
            }
            
            config['Connection'] = {
                'timeout': '30',
                'batch_size': '100'
            }
            
            config['DefaultRules'] = {
                'rule1_subject_pattern': r'^Fw[:：]?(A\d+|A\*\d+|B\d+|C\d+|A\+B\d+|A\+C\d+|A\+B\*\d+|A\+C\*\d+)',
                'rule1_from': 'crackpost2@126.com,crackpost@126.com',
                'rule1_to': '{self_email}',
                
                'rule2_subject_pattern': r'^(A|B|C.*|A\+B.*|A\+C.*)$',
                'rule2_from': '{self_email}',
                'rule2_to': 'crackpost2@126.com,crackpost@126.com',
                'rule2_body_contains': '发件人=【'
            }
            # 创建配置文件目录（如果不存在）
            os.makedirs(os.path.dirname(config_file), exist_ok=True)
            
            # 保存配置文件
            with open(config_file, 'w') as f:
                config.write(f)
            
            logger.info(f"配置文件已创建: {config_file}")
        else:
            config.read(config_file)
            
        return config
    
    def save_config(self):
        """保存配置到文件"""
        # 更新配置对象
        self.config['Credentials'] = {
            'email': self.email_addr,
            'password': self.password
        }
        
        self.config['Servers'] = {
            'imap_server': self.imap_server,
            'imap_port': str(getattr(self, 'imap_port', 993)),
        }
        
        self.config['Storage'] = {
            'base_directory': str(self.base_dir)
        }
        
        self.config['Connection'] = {
            'timeout': str(self.timeout),
            'batch_size': str(self.batch_size)
        }
        
        # 保存到文件
        with open(self.config_file, 'w') as f:
            self.config.write(f)
        
        logger.info(f"配置已保存到: {self.config_file}")
    
    def connect_to_email(self):
        """连接 IMAP 并登录（优先 AUTHENTICATE LOGIN，失败回退 LOGIN），登录后立即发送 IMAP ID"""
        try:
            from imapclient import IMAPClient
            import ssl as _ssl
            if not self.imap_server or not self.email_addr or not self.password:
                raise RuntimeError("邮箱地址/授权码/IMAP服务器未配置完整")
            ssl_context = _ssl.create_default_context()
            self.client = IMAPClient(
                self.imap_server,
                port=int(getattr(self, 'imap_port', 993)),
                ssl=True,
                ssl_context=ssl_context,
                timeout=int(getattr(self, 'timeout', 30))
            )
            self.client.normalise_times = False

            # 优先 AUTHENTICATE LOGIN（对部分 126/163 可用）
            def _auth_login(resp):
                prompt = resp.decode(errors='ignore') if isinstance(resp, (bytes, bytearray)) else str(resp or '')
                return (self.email_addr if ('user' in prompt.lower() or 'username' in prompt.lower()) else self.password).encode()

            try:
                # 使用底层 imap 的 AUTHENTICATE LOGIN（更接近 test.py 的成功路径）
                self.client._imap.authenticate('LOGIN', _auth_login)
                logger.info(f"使用 AUTHENTICATE LOGIN 成功登录 {self.email_addr}")
            except Exception as e:
                logger.warning(f"AUTHENTICATE LOGIN 失败，回退普通 LOGIN: {e}")
                self.client.login(self.email_addr, self.password)
                logger.info(f"成功登录到 {self.email_addr} 的邮箱")

            # 关键：发送 IMAP ID，避免被视为 Unsafe Login（优先使用 IMAPClient.id_）
            try:
                id_kv = {
                    'name': 'Crackpost',
                    'version': '1.0',
                    'vendor': 'Cyclop',
                    'os': sys.platform,
                    'support-url': 'https://github.com/offline229/'
                }
                if hasattr(self.client, 'id_'):
                    try:
                        self.client.id_(id_kv)
                        logger.info("已发送 IMAP ID via IMAPClient.id_()")
                    except Exception as e:
                        logger.debug(f"IMAPClient.id_() 失败: {e}")
                        raise
                else:
                    raise RuntimeError("IMAPClient 无 id_ 方法，使用底层发送 ID")
            except Exception:
                # 回退：通过底层 imap 发送原始 ID 命令（兼容性更高）
                try:
                    # 构造 ID 参数 (\"k\" \"v\" ...)
                    pairs = ' '.join([f'"{k}" "{v}"' for k, v in id_kv.items()])
                    tag = self.client._imap._new_tag()
                    cmd = f"{tag} ID ({pairs})\r\n"
                    self.client._imap.send(cmd.encode())
                    typ, resp = self.client._imap._get_tagged_response(tag)
                    logger.info(f"通过底层 IMAP 发送 ID 返回: {typ} {resp}")
                except Exception as e2:
                    logger.debug(f"底层发送 ID 失败: {e2}")

            return True
        except Exception as e:
            logger.error(f"连接或登录失败: {e}")
            try:
                if self.client:
                    try:
                        self.client.logout()
                    except Exception:
                        pass
            except Exception:
                pass
            self.client = None
            return False
        
    def close_connection(self):
        """关闭邮箱连接"""
        if self.client:
            try:
                self.client.logout()
                logger.info("已关闭邮箱连接")
            except Exception as e:
                logger.error(f"关闭连接时发生错误: {str(e)}")
    
    def get_all_mailboxes(self):
        """返回所有邮箱文件夹名称列表"""
        if not self.client:
            if not self.connect_to_email():
                return []
        
        try:
            # 获取文件夹列表（自动解码）
            folders = self.client.list_folders()
            
            # 创建映射字典 - 保存原始名称用于select操作
            self.mailbox_mapping = {}
            mailbox_names = []
            
            for flags, delimiter, name in folders:
                # IMAPClient已自动解码
                raw_name = name
                display_name = name
                
                # 保存映射
                self.mailbox_mapping[display_name] = raw_name
                mailbox_names.append(display_name)
            
            return mailbox_names
        except Exception as e:
            logger.error(f"获取邮箱文件夹失败: {str(e)}")
            return []
    
    def search_emails_advanced(self, folder, subject_pattern=None, from_address=None,
                            to_address=None, start_date=None, end_date=None, 
                            callback=None, max_emails=None):
        """高级邮件搜索，支持多种条件和分页"""
        
        # ✅ 新增：检查连接状态，如果连接失效则重新连接
        if not self.client:
            logger.warning("客户端为空，正在重新连接...")
            if not self.connect_to_email():
                return []
        
        # ✅ 新增：尝试发送 NOOP 命令检查连接是否仍然有效
        try:
            self.client.noop()  # 发送 NOOP 保活命令
            logger.debug("连接检查: OK")
        except Exception as e:
            logger.warning(f"连接已失效，正在重新连接: {e}")
            try:
                self.client.logout()
            except Exception:
                pass
            self.client = None
            if not self.connect_to_email():
                return []
                
        try:
            # 选择文件夹（EXAMINE -> SELECT 回退）
            try:
                self.client.select_folder(folder, readonly=True)
            except Exception as e:
                logger.warning(f"EXAMINE 失败，回退为 SELECT: {e}")
                try:
                    self.client.select_folder(folder, readonly=False)
                except Exception as e2:
                    logger.error(f"选择邮箱文件夹失败(尝试 SELECT): {e2}")
                    return []
            logger.info(f"已选择文件夹: {folder}")

            # 一进检索就通知前台，避免长时间停在「0 封」
            if callback:
                try:
                    callback(0, 0)
                except Exception:
                    pass
            
            # 构建搜索条件
            search_criteria = []
            
            if start_date:
                start_date_str = start_date.strftime('%d-%b-%Y')
                search_criteria += ['SINCE', start_date_str]

            if end_date:
                next_day = (end_date + timedelta(days=1)).strftime('%d-%b-%Y')
                search_criteria += ['BEFORE', next_day]
            
            # 如果没有设置任何条件，搜索全部
            if not search_criteria:
                search_criteria.append('ALL')
            
            # 执行初始搜索，获取邮件ID
            logger.info(f"搜索条件: {search_criteria}")
            msg_ids = self.client.search(search_criteria)
            
            if not msg_ids:
                logger.info("未找到匹配的邮件")
                if callback:
                    try:
                        callback(100, 0)
                    except Exception:
                        pass
                return []
            
            logger.info(f"找到 {len(msg_ids)} 封邮件，开始处理...")
            if callback:
                try:
                    callback(1, 0)
                except Exception:
                    pass
            
            # 如果设置了最大数量，限制处理的邮件数
            if max_emails and len(msg_ids) > max_emails:
                msg_ids = msg_ids[-max_emails:]  # 取最新的N封
                
            # 分批处理
            matched_emails = []
            batch_size = min(self.batch_size, 100)  # 每批处理的邮件数
            
            for i in range(0, len(msg_ids), batch_size):
                batch_ids = msg_ids[i:i+batch_size]
                try:
                    # 获取这批邮件的信息
                    response = self.client.fetch(batch_ids, ['ENVELOPE', 'FLAGS'])
                    
                    for msg_id, data in response.items():
                        envelope = data[b'ENVELOPE']
                        subject = self.decode_mime_header(envelope.subject.decode()) if envelope.subject else ""
                        subject = re.sub(r'\s+', '', subject)  # 入口处去空格
                        # 进一步筛选
                        if subject_pattern and not re.search(subject_pattern, subject, re.IGNORECASE):
                            continue
                            
                        # 检查发件人
                        sender = ""
                        if envelope.from_ and len(envelope.from_) > 0:
                            sender = envelope.from_[0].mailbox.decode() + '@' + envelope.from_[0].host.decode()

                        if from_address:
                            # 支持多个邮箱，逗号分隔
                            from_list = [x.strip() for x in from_address.split(",") if x.strip()]
                            if not any(f in sender for f in from_list):
                                continue

                        # 检查收件人
                        recipients = []
                        if envelope.to:
                            for recipient in envelope.to:
                                if recipient.mailbox and recipient.host:
                                    email_addr = recipient.mailbox.decode() + '@' + recipient.host.decode()
                                    recipients.append(email_addr)

                        if to_address:
                            to_list = [x.strip() for x in to_address.split(",") if x.strip()]
                            if not any(any(t in r for t in to_list) for r in recipients):
                                continue
                            
                        # 处理日期
                        if envelope.date:
                            if isinstance(envelope.date, datetime):
                                date_str = envelope.date.strftime("%Y-%m-%d %H:%M:%S")
                                date_obj = envelope.date
                            elif isinstance(envelope.date, bytes):
                                date_str = envelope.date.decode()
                                try:
                                    date_obj = email.utils.parsedate_to_datetime(date_str)
                                except:
                                    date_obj = None
                            else:
                                date_str = str(envelope.date)
                                date_obj = None
                        else:
                            date_str = ""
                            date_obj = None
                            
                        # 匹配成功，添加到结果
                        matched_emails.append({
                            'id': msg_id,
                            'subject': subject,
                            'sender': sender,
                            'date': date_str,
                            'date_obj': date_obj,
                            'folder': folder,
                        })
                        
                    # 更新进度回调
                    if callback:
                        progress = min(100, int((i + len(batch_ids)) / len(msg_ids) * 100))
                        callback(progress, len(matched_emails))
                        
                except Exception as e:
                    logger.error(f"处理批次 {i//batch_size + 1} 时出错: {str(e)}")
                    continue
                    
            # 按日期排序（最新的在前）
            matched_emails.sort(key=lambda x: x['date_obj'] if x['date_obj'] else datetime.min, reverse=True)
            
            logger.info(f"成功找到 {len(matched_emails)} 封匹配邮件")
            # 筛主题不合规时也可能较慢，保持「正在检索」观感
            if callback:
                try:
                    callback(99, len(matched_emails))
                except Exception:
                    pass
            # 来往地址对得上但主题不合规 → 结构化未匹配，供界面手工补类型/收发件人
            unmatched = []
            if subject_pattern:
                matched_ids = {e['id'] for e in matched_emails}
                for msg_id in msg_ids:
                    if msg_id in matched_ids:
                        continue
                    try:
                        response = self.client.fetch([msg_id], ['ENVELOPE'])
                        envelope = response[msg_id][b'ENVELOPE']
                        subject = self.decode_mime_header(envelope.subject.decode()) if envelope.subject else ""
                        subject = re.sub(r'\s+', '', subject)

                        sender = ""
                        if envelope.from_ and len(envelope.from_) > 0:
                            sender = envelope.from_[0].mailbox.decode() + '@' + envelope.from_[0].host.decode()

                        recipients = []
                        if envelope.to:
                            for recipient in envelope.to:
                                if recipient.mailbox and recipient.host:
                                    recipients.append(
                                        recipient.mailbox.decode() + '@' + recipient.host.decode()
                                    )

                        sender_ok = True
                        recipient_ok = True
                        if from_address:
                            from_list = [x.strip() for x in from_address.split(",") if x.strip()]
                            sender_ok = any(f in sender for f in from_list)
                        if to_address:
                            to_list = [x.strip() for x in to_address.split(",") if x.strip()]
                            recipient_ok = any(any(t in r for t in to_list) for r in recipients)

                        if not (sender_ok and recipient_ok):
                            continue

                        if envelope.date:
                            if isinstance(envelope.date, datetime):
                                date_str = envelope.date.strftime("%Y-%m-%d %H:%M:%S")
                                date_obj = envelope.date
                            else:
                                date_str = str(envelope.date)
                                date_obj = None
                        else:
                            date_str = ""
                            date_obj = None

                        unmatched.append({
                            'id': msg_id,
                            'subject': subject,
                            'sender': sender,
                            'recipients': recipients,
                            'date': date_str,
                            'date_obj': date_obj,
                            'folder': folder,
                            'direction': self.direction_from_addresses(sender, recipients),
                        })
                    except Exception as e:
                        logger.debug(f"处理未匹配邮件时出错 id={msg_id}: {e}")

                logger.info(f"总邮件数: {len(msg_ids)}")
                logger.info(f"匹配当前规则邮件数: {len(matched_emails)}")
                logger.info(f"未匹配当前规则邮件数: {len(unmatched)}")
                if unmatched:
                    logger.info(f"未匹配邮件数量: {len(unmatched)}，前{min(300, len(unmatched))}个主题如下：")
                    for i, item in enumerate(unmatched[:300]):
                        print(f"[未匹配{i+1}] {item.get('subject', '')}")

            self.last_unmatched = unmatched
            return matched_emails
            
        except Exception as e:
            logger.error(f"搜索邮件时出错: {str(e)}")
            return []


    # ========== 1. 获取原始邮件数据 ==========

    def ensure_connected(self, max_retries=2, backoff=1):
        """确保 IMAP 连接可用；不可用时自动重连并重试（返回 True/False）"""
        for attempt in range(1, max_retries + 1):
            try:
                if not self.client:
                    logger.info("IMAP 客户端为空，尝试连接...")
                    if self.connect_to_email():
                        return True
                    else:
                        logger.warning("connect_to_email 失败")
                else:
                    try:
                        # 发送 NOOP 保活检测
                        self.client.noop()
                        return True
                    except Exception as e:
                        logger.warning(f"连接检测失败（尝试 {attempt}/{max_retries}）：{e}")
                        try:
                            self.client.logout()
                        except Exception:
                            pass
                        self.client = None
                        if self.connect_to_email():
                            return True
            except Exception as e:
                logger.warning(f"ensure_connected 异常: {e}")
            # 指数回退
            time.sleep(backoff * attempt)
        logger.error("无法建立有效的 IMAP 连接（ensure_connected 失败）")
        return False

    def _fetch_raw_email(self, email_id, folder):
        """获取原始邮件数据（带重连与重试逻辑）

        Returns:
            tuple: (email.message.Message对象, 错误信息)
        """
        try:
            # 先确保连接可用
            if not self.ensure_connected():
                return None, "无法连接到邮箱（ensure_connected 失败）"

            # 选择文件夹（EXAMINE -> SELECT 回退），遇到错误尝试重连一次
            try:
                self.client.select_folder(folder, readonly=True)
            except Exception as e:
                logger.warning(f"EXAMINE 失败，尝试重连并回退为 SELECT: {e}")
                try:
                    self.client.logout()
                except Exception:
                    pass
                self.client = None
                if not self.ensure_connected():
                    return None, f"选择文件夹失败: 无法重连 ({e})"
                try:
                    self.client.select_folder(folder, readonly=False)
                except Exception as e2:
                    return None, f"选择文件夹失败: {str(e2)}"

            # 获取邮件（fetch），fetch 失败时重连、重新选文件夹再重试
            try:
                fetch_data = self.client.fetch([email_id], ['RFC822'])
                if not fetch_data or email_id not in fetch_data:
                    return None, f"获取邮件 {email_id} 失败"
            except Exception as e:
                logger.warning(f"fetch 失败，尝试重连并重试: {e}")
                try:
                    self.client.logout()
                except Exception:
                    pass
                self.client = None
                if not self.ensure_connected():
                    return None, f"获取邮件失败: 无法重连 ({e})"
                try:
                    self.client.select_folder(folder, readonly=True)
                except Exception:
                    try:
                        self.client.select_folder(folder, readonly=False)
                    except Exception as e_sel:
                        return None, f"获取邮件失败: 重连后无法选择文件夹 ({e_sel})"
                try:
                    fetch_data = self.client.fetch([email_id], ['RFC822'])
                    if not fetch_data or email_id not in fetch_data:
                        return None, f"获取邮件 {email_id} 失败（重试）"
                except Exception as e2:
                    return None, f"获取邮件失败: {str(e2)}"

            raw_email = fetch_data[email_id][b'RFC822']
            msg = email.message_from_bytes(raw_email)
            return msg, None

        except Exception as e:
            return None, f"获取邮件时出错: {str(e)}"

    # ========== 2. 解析基本元数据 ==========
    def _parse_email_metadata(self, msg):
        """解析邮件的基本元数据
        
        Returns:
            dict: {'subject', 'sender', 'date_str', 'date_folder'}
        """
        metadata = {}
        
        # 主题
        raw_subject = msg['subject'] or ''
        subject = self.decode_mime_header(raw_subject)

        # 如果原始内容是 Q-encoding，则应用 _ => space
        if '?Q?' in raw_subject.upper():
            subject = subject.replace('_', ' ')
        metadata['subject'] = re.sub(r'\s+', '', subject)
        
        # 发件人
        metadata['sender'] = email.utils.parseaddr(msg['from'])[1]
        
        # 日期
        date_str = msg['date'] or ''
        metadata['date_str'] = date_str
        
        try:
            date_obj = email.utils.parsedate_to_datetime(date_str)
            metadata['date_folder'] = date_obj.strftime("%Y%m%d")
        except:
            metadata['date_folder'] = "unknown_date"
        
        return metadata
    
    # ========== 3. 确定存储目录 ==========
    def _determine_email_directory(self, metadata):
        """确定邮件存储目录
        
        Returns:
            Path: 邮件存储目录路径
        """
        current_email = self.email_addr.lower()
        sender = metadata['sender']
        send_type = "发" if sender.lower() == current_email else "收"
        
        subject = metadata['subject'][:30]
        safe_subject = self._safe_filename(subject)
        
        email_dir = self.base_dir / send_type / f"{metadata['date_folder']}_{safe_subject}"
        email_dir.mkdir(parents=True, exist_ok=True)
        
        return email_dir
    
    # ========== 4. 解析收件人 ==========
    def _parse_recipients(self, msg):
        """解析收件人列表
        
        Returns:
            list: 收件人名称列表
        """
        to_addrs = []
        try:
            addrs = email.utils.getaddresses(
                msg.get_all('To', []) + msg.get_all('Cc', [])
            )
            for name, addr in addrs:
                if name and name.strip():
                    decoded_name = self.decode_mime_header(name.strip())
                    to_addrs.append(decoded_name)
                elif addr:
                    local = addr.split('@')[0]
                    to_addrs.append(local)
        except Exception as e:
            logger.warning(f"解析收件人失败: {e}")
        
        return to_addrs
    
    # ========== 5. 提取邮件正文 ==========
    def _extract_email_body(self, msg):
        """提取邮件正文(优先纯文本,备选HTML)
        
        Returns:
            str: 邮件正文文本
        """
        body_text = ""
        found_body = False
        
        for part in msg.walk():
            content_type = part.get_content_type()
            content_disp = str(part.get("Content-Disposition"))
            
            # 优先纯文本
            if content_type == "text/plain" and "attachment" not in content_disp:
                charset = part.get_content_charset()
                try:
                    if charset:
                        body_text = part.get_payload(decode=True).decode(charset, errors='replace')
                    else:
                        body_text = part.get_payload(decode=True).decode(errors='replace')
                    logger.debug(f"原始纯文本正文: {repr(body_text[:100])}")
                except:
                    body_text = "无法解码邮件内容"
                found_body = True
                break
            
            # 备选 HTML
            elif content_type == "text/html" and "attachment" not in content_disp and not found_body:
                charset = part.get_content_charset()
                try:
                    if charset:
                        html_text = part.get_payload(decode=True).decode(charset, errors='replace')
                    else:
                        html_text = part.get_payload(decode=True).decode(errors='replace')
                    
                    logger.debug(f"原始HTML正文: {repr(html_text[:100])}")
                    
                    # HTML转文本
                    html_text = re.sub(r'(<br\s*/?>|<p>|</p>|<div[^>]*>|</div>)', '\n', html_text, flags=re.IGNORECASE)
                    html_text = html_text.replace('&nbsp;', ' ')
                    body_text = re.sub('<[^<]+?>', '', html_text)
                    body_text = re.sub(r'\n+', '\n', body_text)
                    
                    logger.debug(f"最终处理后的正文: {repr(body_text[:100])}")
                    found_body = True
                except:
                    body_text = "无法解码 HTML 邮件内容"
                    found_body = True
        
        return body_text
    
    # ========== 6. 下载附件 ==========
    def _download_attachments(self, msg, email_dir, email_id):
        """下载所有附件
        
        Returns:
            tuple: (附件路径列表, 错误信息列表)
        """
        attachment_paths = []
        attachment_errors = []
        
        for part in msg.walk():
            content_disp = str(part.get("Content-Disposition"))
            
            if "attachment" in content_disp or part.get_filename():
                filename = part.get_filename()
                if filename:
                    if isinstance(filename, bytes):
                        filename = filename.decode(errors='replace')
                    filename = self.decode_mime_header(filename)
                    filename = re.sub(r'[\\/:*?"<>|]', '_', filename)
                    
                    if not filename or filename.strip() == "":
                        filename = f"attachment_{email_id}_{int(time.time())}.bin"
                    
                    file_path = email_dir / filename
                    try:
                        with open(file_path, 'wb') as f:
                            payload = part.get_payload(decode=True)
                            if payload:
                                f.write(payload)
                        attachment_paths.append(str(file_path))
                        logger.info(f"保存附件: {file_path}")
                    except Exception as e:
                        error_msg = f"保存附件 {filename} 失败: {str(e)}"
                        logger.error(error_msg)
                        attachment_errors.append(error_msg)
        
        return attachment_paths, attachment_errors
    
    # ========== 7. 保存 content.txt ==========
    def _save_content_file(self, email_dir, metadata, to_addrs, body_text):
        """保存 content.txt 文件
        
        Returns:
            tuple: (成功标志, 错误信息)
        """
        content_file = email_dir / "content.txt"
        try:
            with open(content_file, "w", encoding="utf-8") as f:
                f.write(f"主题: {metadata['subject']}\n")
                f.write(f"发件邮箱: {metadata['sender']}\n")
                f.write(f"收件邮箱: {', '.join(to_addrs)}\n")
                f.write(f"日期: {metadata['date_str']}\n")
                f.write("-" * 50 + "\n\n")
                f.write(body_text if body_text.strip() else "[邮件内容为空或无法解析]")
            logger.info(f"已保存 content.txt: {content_file}")
            return True, None
        except Exception as e:
            error_msg = f"保存 content.txt 失败: {str(e)}"
            logger.error(error_msg)
            return False, error_msg
    
    # ========== 8. 验证下载完整性 ==========
    def _validate_download(self, email_dir, attachment_errors):
        """验证下载完整性
        
        Returns:
            tuple: (成功标志, 错误信息)
        """
        content_file = email_dir / "content.txt"
        
        if not content_file.exists():
            return False, "content.txt 不存在"
        
        if content_file.stat().st_size <= 100:
            return False, "content.txt 文件过小"
        
        if attachment_errors:
            return False, "; ".join(attachment_errors)
        
        return True, None
    
    # ========== 9. 主流程协调器 ==========
    def download_email(self, email_id, folder):
        """下载单个邮件(主流程)
        
        Returns:
            tuple: (result字典, 错误信息)
        """
        if not self.client:
            if not self.connect_to_email():
                return None, "连接失败"

        try:
            # 1. 获取原始邮件
            msg, error = self._fetch_raw_email(email_id, folder)
            if error:
                return None, error
            
            # 2. 解析元数据
            metadata = self._parse_email_metadata(msg)
            
            # 3. 确定存储目录
            email_dir = self._determine_email_directory(metadata)

            # 4. 解析收件人
            to_addrs = self._parse_recipients(msg)
            
            # 5. 提取正文
            body_text = self._extract_email_body(msg)
            
            # 6. 下载附件
            attachment_paths, attachment_errors = self._download_attachments(msg, email_dir, email_id)
            
            # 7. 保存 content.txt
            success, error = self._save_content_file(email_dir, metadata, to_addrs, body_text)
            if not success:
                return None, error
            
            # 8. 验证完整性
            success, error = self._validate_download(email_dir, attachment_errors)
            if not success:
                return None, error
            
            # 9. 构建结果
            result = {
                'subject': metadata['subject'],
                'sender': metadata['sender'],
                'date': metadata['date_str'],
                'has_attachments': len(attachment_paths) > 0,
                'directory': str(email_dir),
                'attachment_count': len(attachment_paths),
                'attachments': attachment_paths,
                'to_addrs': to_addrs
            }
            
            logger.info(f"处理完成: {metadata['subject']}")
            return result, None

        except Exception as e:
            error_msg = f"下载邮件 {email_id} 时出错: {str(e)}"
            logger.error(error_msg)
            return None, error_msg

    def download_multiple_emails(self, email_list, progress_callback=None, max_retries=2):
        """下载多封邮件，支持智能跳过和详细错误报告"""
        results = []
        total = len(email_list)
        succeeded_ids = set()  # 真正下载成功的
        skipped_ids = set()    # 跳过的
        failed_details = []

        # 第一轮：检查+下载
        for i, email_info in enumerate(email_list):
            email_id = email_info['id']
            folder = email_info['folder']

            def _attach_user_overrides(res):
                for key in ('user_letter_type', 'user_sender_name', 'user_receiver_name', 'user_direction'):
                    if email_info.get(key):
                        res[key] = email_info[key]
                return res

            # 判断是否已完整下载
            already, email_dir, reason = self._email_already_downloaded(email_info)
            if already:
                logger.info(f"✓ 跳过已下载邮件 id={email_id} 主题={email_info.get('subject')}")
                results.append(_attach_user_overrides({
                    'subject': email_info.get('subject', ''),
                    'sender': email_info.get('sender', ''),
                    'date': email_info.get('date', ''),
                    'has_attachments': False,
                    'directory': email_dir,
                    'attachment_count': 0,
                    'attachments': [],
                    'to_addrs': [],
                    'status': 'skipped',
                    'reason': reason
                }))
                skipped_ids.add(email_id)  # ✅ 修改：单独记录跳过的
            else:
                # 需要下载
                logger.info(f"→ 开始下载 id={email_id} 主题={email_info.get('subject')} (原因: {reason})")
                result, error = self.download_email(email_id, folder)
                if result:
                    result['status'] = 'success'
                    result['reason'] = '下载成功'
                    results.append(_attach_user_overrides(result))
                    succeeded_ids.add(email_id)  # ✅ 只记录真正下载成功的
                    logger.info(f"✓ 下载成功 id={email_id}")
                else:
                    logger.warning(f"✗ 下载失败 id={email_id}: {error}")
                    failed_details.append({
                        'id': email_id,
                        'subject': email_info.get('subject', ''),
                        'error': error
                    })

            if progress_callback:
                progress = int((i + 1) / total * 100)
                # ✅ 修改：更新进度时传递当前邮件的状态
                email_info['status'] = 'skipped' if email_id in skipped_ids else ('success' if email_id in succeeded_ids else '')
                email_info['reason'] = results[-1].get('reason', '') if results else ''
                progress_callback(progress, i + 1, total)

            if i < total - 1:
                time.sleep(0.3)

        # 重试失败的项
        failed_list = [e for e in email_list if e['id'] not in succeeded_ids and e['id'] not in skipped_ids]
        attempt = 0
        while failed_list and attempt < max_retries:
            attempt += 1
            logger.info(f"\n【第 {attempt} 轮重试】待重试数: {len(failed_list)}")
            
            for email_info in list(failed_list):
                email_id = email_info['id']
                folder = email_info['folder']
                
                # 再次检查
                already, email_dir, reason = self._email_already_downloaded(email_info)
                def _attach_ov(res):
                    for key in ('user_letter_type', 'user_sender_name', 'user_receiver_name', 'user_direction'):
                        if email_info.get(key):
                            res[key] = email_info[key]
                    return res

                if already:
                    logger.info(f"[重试{attempt}] ✓ 已有有效内容 id={email_id}")
                    results.append(_attach_ov({
                        'subject': email_info.get('subject', ''),
                        'sender': email_info.get('sender', ''),
                        'date': email_info.get('date', ''),
                        'has_attachments': False,
                        'directory': email_dir,
                        'attachment_count': 0,
                        'attachments': [],
                        'to_addrs': [],
                        'status': 'skipped',
                        'reason': reason
                    }))
                    failed_list.remove(email_info)
                    skipped_ids.add(email_id)
                    # 从失败列表中移除
                    failed_details = [f for f in failed_details if f['id'] != email_id]
                    continue

                # 重新下载
                res, error = self.download_email(email_id, folder)
                if res:
                    res['status'] = 'success'
                    res['reason'] = f'重试{attempt}成功'
                    results.append(_attach_ov(res))
                    failed_list.remove(email_info)
                    succeeded_ids.add(email_id)
                    # 更新失败列表
                    for fd in failed_details:
                        if fd['id'] == email_id:
                            failed_details.remove(fd)
                            break
                    logger.info(f"[重试{attempt}] ✓ 成功 id={email_id}")
                else:
                    logger.warning(f"[重试{attempt}] ✗ 仍失败 id={email_id}: {error}")
                    # 更新错误信息
                    for fd in failed_details:
                        if fd['id'] == email_id:
                            fd['error'] = f"重试{attempt}后仍失败: {error}"
                            break
                time.sleep(0.5)

        # 标记最终失败的项
        if failed_list:
            logger.error(f"\n【最终失败】{len(failed_list)} 封邮件无法下载")
            for email_info in failed_list:
                results.append({
                    'subject': email_info.get('subject', ''),
                    'sender': email_info.get('sender', ''),
                    'date': email_info.get('date', ''),
                    'has_attachments': False,
                    'directory': 'FAILED',
                    'attachment_count': 0,
                    'attachments': [],
                    'to_addrs': [],
                    'status': 'failed',
                    'reason': '多次重试后失败'
                })

        # ✅ 修改：统计更清晰
        success_count = len(succeeded_ids)
        skipped_count = len(skipped_ids)
        failed_count = len(failed_list)

        # 生成详细报告
        summary = f"\n\n{'='*60}\n下载汇总:\n✓ 成功: {success_count} 封\n⊙ 跳过: {skipped_count} 封\n✗ 失败: {failed_count} 封\n{'='*60}\n"
        logger.info(summary)

        # 失败详情
        if failed_details:
            logger.error("\n失败详情:")
            for fd in failed_details:
                logger.error(f"  - [{fd['id']}] {fd['subject']}: {fd['error']}")

        # 生成报告
        self.generate_download_report(results, email_list, failed_details)

        return results
    

    # 所有涉及收发件人推断的操作在此步进行
    def generate_download_report(self, results, email_list, failed_details=None):
        """生成报告，TSV 去重，记录失败详情"""
        report_path = self.base_dir / "download_report.txt"
        global_result_path = self.base_dir / "global_result.tsv"
        current_email = self.email_addr.lower()

        # 1. 读取已有 TSV
        existing_oc = set()
        existing_mails = set()
        existing_rows = []
        existing_aliases_line = "" # 【新增】：保存可能的别名行
        
        if global_result_path.exists():
            with open(global_result_path, "r", encoding="utf-8") as f:
                lines = f.readlines()
                if lines:
                    data_start_idx = 0
                    for i, line in enumerate(lines):
                        if line.startswith("oc_registerdata"):
                            oc_line = line.strip().split('\t')
                            if len(oc_line) > 1 and oc_line[1]:
                                existing_oc = set([x.strip() for x in oc_line[1].split(',') if x.strip()])
                        elif line.startswith("oc_aliases"):
                            # 【新增】：提取并保护旧别名数据
                            existing_aliases_line = line.strip()
                        elif line.startswith("收/发类型") or line.startswith("收/发"):
                            data_start_idx = i
                            break
                            
                    # 解析邮件（关键：用更多字段构建唯一键）
                    for line in lines[data_start_idx+1:]:
                        cols = line.strip().split('\t')
                        if len(cols) >= 6:
                            # 唯一标识：发件人 + 日期 + 主题/类型 + 路径
                            key = (cols[2], cols[1], cols[4], cols[5])
                            existing_mails.add(key)
                            existing_rows.append(line.strip())

        # 2. 合并 OC
        all_oc = set(existing_oc)
        all_oc.update(self.oc_registerdata)
        all_oc = sorted(all_oc)
        only_oc_name = None
        if len(all_oc) == 1:
            only_oc_name = all_oc[0].split('_')[0]

        # 3. 生成新邮件数据（去重）
        new_rows = []
        project_root = Path(os.path.dirname(os.path.abspath(__file__))).resolve()
        duplicate_count = 0

        # 合并内存中的别名（若报告生成时尚未写过别名行）
        if self.oc_aliases and not existing_aliases_line:
            groups = []
            for main, subs in self.oc_aliases.items():
                if subs:
                    groups.append(main + "=" + "=".join(subs))
            if groups:
                existing_aliases_line = "oc_aliases\t" + ",".join(groups)

        for result in results:
            # 跳过失败的
            if result.get('status') == 'failed':
                continue

            mail_sender = result.get('sender', '').strip()
            send_type = "发" if mail_sender.lower() == current_email else "收"
            date_str = result.get('date', '')
            try:
                date_obj = email.utils.parsedate_to_datetime(date_str)
                date_short = date_obj.strftime("%Y-%m-%d")
            except:
                date_short = date_str[:10]

            # 自动补全发件人
            content_path = Path(result['directory']) / "content.txt"
            sender_name = ""
            if content_path.exists():
                with open(content_path, "r", encoding="utf-8") as cf:
                    content = cf.read()
                    m = re.search(r"发件[^=]*=【(.+?)】", content)
                    if m:
                        sender_name = m.group(1)
            if not sender_name:
                m = re.search(r"来自(.+?)的信", str(result['directory']))
                if m:
                    sender_name = m.group(1)
            if not sender_name:
                sender_name = mail_sender
            if not sender_name:
                sender_name = "未知"

            receiver_name = ""
            subject = result.get('subject', '')
            subject = re.sub(r'\s+', '_', subject)
            letter_type = self.extract_letter_type(subject)
            if not letter_type:
                print(f"[DEBUG] letter_type解析失败，输入subject: '{subject}' 路径: {result.get('directory','')}")

            # 发件时的收件人逻辑
            if send_type == "发":
                if letter_type and letter_type.lower().startswith("c"):
                    receiver_raw = letter_type[1:]
                    receiver_name = re.sub(r'[\s=\+\[\]]+', '', receiver_raw)
                else:
                    receiver_name = ""
            elif send_type == "收" and not receiver_name and only_oc_name:
                receiver_name = only_oc_name

            # 用户指定覆盖（未匹配手工入库，字段由 download_multiple 抄到 result）
            if result.get('user_letter_type'):
                letter_type = result['user_letter_type']
            if result.get('user_sender_name'):
                sender_name = result['user_sender_name']
            if result.get('user_receiver_name'):
                receiver_name = result['user_receiver_name']
            if result.get('user_direction') in ("收", "发"):
                send_type = result['user_direction']

            save_path = result.get('directory', '')
            try:
                save_dir = Path(save_path).resolve()
                rel = save_dir.relative_to(project_root)
                save_path_formatted = '.\\' + str(rel).replace('/', '\\')
            except Exception:
                try:
                    save_dir = Path(save_path).resolve()
                    rel2 = save_dir.relative_to(self.base_dir.parent.resolve())
                    save_path_formatted = '.\\' + str(rel2).replace('/', '\\')
                except Exception:
                    save_path_formatted = save_path

            # 新增：用更强唯一键（发件人、收件人、日期、类型、路径）
            key = (sender_name, receiver_name, date_short, letter_type, save_path_formatted)
            if key not in existing_mails:
                row = f"{send_type}\t{date_short}\t{sender_name}\t{receiver_name}\t{letter_type}\t{save_path_formatted}"
                new_rows.append(row)
                existing_mails.add(key)
            else:
                duplicate_count += 1
                logger.info(f"跳过重复项: {sender_name} {receiver_name} {date_short} {letter_type}")

        # 4. 合并 old+new 后全局去重
        all_rows = existing_rows + new_rows
        unique_row_keys = set()
        final_rows = []
        for row in all_rows:
            cols = row.strip().split('\t')
            if len(cols) < 6:
                continue
            row_key = tuple(cols[:6])  # 用全部6列做唯一性
            if row_key not in unique_row_keys:
                unique_row_keys.add(row_key)
                final_rows.append(row)

        # 5. 写入 TSV
        with open(global_result_path, "w", encoding="utf-8") as f:
            f.write(f"oc_registerdata\t{','.join(all_oc)}\n")
            if existing_aliases_line: # 【新增】：写回别名行
                f.write(existing_aliases_line + "\n")
            f.write("收/发类型\t发信日期\t发件人\t收件人\t信件类型\t信件下载位置\n")
            for row in final_rows:
                f.write(row + "\n")

        # 6. 写入下载报告
        with open(report_path, "w", encoding="utf-8") as f:
            f.write("邮件下载报告\n")
            f.write("=" * 50 + "\n\n")
            f.write(f"下载时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"总邮件数: {len(email_list)}\n")
            f.write(f"成功下载: {len([r for r in results if r.get('status') == 'success'])}\n")
            f.write(f"跳过已有: {len([r for r in results if r.get('status') == 'skipped'])}\n")
            f.write(f"下载失败: {len([r for r in results if r.get('status') == 'failed'])}\n")
            f.write(f"跳过重复: {duplicate_count}\n")
            f.write(f"存储目录: {self.base_dir}\n\n")
            
            # 成功列表
            f.write("成功下载的邮件:\n")
            f.write("-" * 50 + "\n")
            for i, result in enumerate([r for r in results if r.get('status') == 'success']):
                f.write(f"{i+1}. {result['subject']}\n")
                f.write(f"   发件邮箱: {result['sender']}\n")
                f.write(f"   日期: {result['date']}\n")
                f.write(f"   保存位置: {result['directory']}\n")
                if result['has_attachments']:
                    f.write(f"   附件数量: {result['attachment_count']}\n")
                f.write(f"   原因: {result.get('reason', '')}\n\n")
            
            # 跳过列表
            f.write("\n跳过的邮件:\n")
            f.write("-" * 50 + "\n")
            for i, result in enumerate([r for r in results if r.get('status') == 'skipped']):
                f.write(f"{i+1}. {result['subject']}\n")
                f.write(f"   原因: {result.get('reason', '')}\n\n")
            
            # 失败详情
            if failed_details:
                f.write("\n失败的邮件（详细）:\n")
                f.write("-" * 50 + "\n")
                for i, fd in enumerate(failed_details):
                    f.write(f"{i+1}. [{fd['id']}] {fd['subject']}\n")
                    f.write(f"   错误原因: {fd['error']}\n\n")

        logger.info(f"全局结果已生成: {global_result_path}")
        logger.info(f"下载报告已生成: {report_path}")
        return report_path


class EmailDownloaderGUI:
    MAIL_PRESETS = {
        "QQ邮箱": "imap.qq.com",
        "163邮箱": "imap.163.com",
        "126邮箱": "imap.126.com",
        "Gmail": "imap.gmail.com",
        "Outlook": "outlook.office365.com",
        "其他": "",
    }

    def __init__(self, root):
        self.root = root
        apply_app_window_icon(self.root)

        self.root.title("地缝邮件下载工具")
        self.root.geometry("720x640")
        self.root.minsize(560, 480)

        self.downloader = EmailDownloader()
        self.selected_folder = ""
        self.matched_results = []
        self.unmatched_results = []
        self.manual_queue = []  # 用户从「未匹配」补录进来的信
        self.busy = False
        self.oc_edit_index = None

        self.create_widgets()
        self.initialize_app()

    # ---------- 智能猜测（入库后自动调用；notify=False 不弹窗）----------
    def smart_guess_receivers(self, notify=True):
        global_result_path = self.downloader.base_dir / "global_result.tsv"
        if not global_result_path.exists():
            msg = "还没有邮件记录，无法识别收件人。"
            if notify:
                messagebox.showerror("错误", "未找到 global_result.tsv")
            return msg

        with open(global_result_path, "r", encoding="utf-8") as f:
            lines = f.readlines()
        if not lines or not lines[0].startswith("oc_registerdata"):
            msg = "邮件记录格式不对，无法识别收件人。"
            if notify:
                messagebox.showerror("错误", "global_result.tsv 格式错误")
            return msg

        oc_line = lines[0].strip().split('\t')
        oc_names = []
        oc_date_map = {}
        if len(oc_line) > 1 and oc_line[1]:
            for x in oc_line[1].split(','):
                if x.strip():
                    parts = x.strip().split('_')
                    if len(parts) == 2:
                        oc_names.append(parts[0])
                        oc_date_map[parts[0]] = parts[1]

        # 主名 → 别名列表；正文匹配时别名也算到主名上
        aliases_by_main = {oc: list(self.downloader.oc_aliases.get(oc, [])) for oc in oc_names}
        data_start = 1
        for i, line in enumerate(lines):
            if line.startswith("oc_aliases"):
                cols = line.strip().split('\t')
                if len(cols) > 1 and cols[1]:
                    for group in cols[1].split(','):
                        parts = [p.strip() for p in group.split('=') if p.strip()]
                        if len(parts) > 1 and parts[0] in aliases_by_main:
                            merged = list(aliases_by_main[parts[0]])
                            for a in parts[1:]:
                                if a not in merged:
                                    merged.append(a)
                            aliases_by_main[parts[0]] = merged
            elif line.startswith("收/发类型") or line.startswith("收/发"):
                data_start = i
                break

        only_one_oc = len(oc_names) == 1
        only_oc_name = oc_names[0] if only_one_oc else None

        def _count_oc_in_text(content_text):
            """主名与别名都计入该角色；较长的名字优先匹配，减少短名误伤。"""
            oc_count = {oc: 0 for oc in oc_names}
            remaining = content_text
            # 按名字长度降序替换计数，避免「甲」吃掉「甲二」之类
            tokens = []
            for oc in oc_names:
                tokens.append((oc, oc))
                for alias in aliases_by_main.get(oc, []):
                    tokens.append((oc, alias))
            tokens.sort(key=lambda t: len(t[1]), reverse=True)
            for main, token in tokens:
                if not token:
                    continue
                n = remaining.count(token)
                if n:
                    oc_count[main] += n
                    remaining = remaining.replace(token, "\0" * len(token))
            return oc_count

        changed = 0
        already_had = 0
        empty_inbound = 0
        context_logs = []
        new_lines = lines[: data_start + 1]
        oc_guess_count = {oc: 0 for oc in oc_names}
        unguessed_count = 0

        for idx, row in enumerate(lines[data_start + 1:]):
            cols = row.strip().split('\t')
            if len(cols) < 6:
                new_lines.append(row)
                continue

            send_type = cols[0]
            subject = cols[4]
            send_date = cols[1]
            receiver = cols[3]
            directory = cols[5]

            if send_type == "收":
                if receiver:
                    already_had += 1
                else:
                    empty_inbound += 1
                    if only_one_oc:
                        cols[3] = only_oc_name
                        changed += 1
                        oc_guess_count[only_oc_name] += 1
                        context_logs.append(f"[单OC] 行{data_start + idx + 2} → {only_oc_name}")
                    elif subject.startswith("A") or subject.startswith("a") or subject.startswith("B") or subject.startswith("b"):
                        oc_to_fill = []
                        for oc in oc_names:
                            reg_date = oc_date_map.get(oc, "")
                            try:
                                if reg_date and send_date and reg_date <= send_date:
                                    oc_to_fill.append(oc)
                            except Exception:
                                pass
                        if oc_to_fill:
                            cols[3] = ",".join(oc_to_fill)
                            changed += 1
                            for oc in oc_to_fill:
                                oc_guess_count[oc] += 1
                            context_logs.append(f"[A/B] 行{data_start + idx + 2} → {cols[3]}")
                        else:
                            unguessed_count += 1
                    elif subject.startswith("C") or subject.startswith("A+C"):
                        project_root = script_dir.resolve()
                        rel_path = directory[2:] if directory.startswith('.\\') else directory
                        content_path = project_root / rel_path.replace('\\', os.sep) / "content.txt"
                        if content_path.exists():
                            try:
                                content_text = content_path.read_text(encoding="utf-8")
                                oc_count = _count_oc_in_text(content_text)
                                nonzero = [oc for oc, c in oc_count.items() if c > 0]
                                unique = [oc for oc, c in oc_count.items() if c == 1]
                                chosen = None
                                if len(nonzero) == 1:
                                    chosen = nonzero[0]
                                elif len(unique) == 1:
                                    chosen = unique[0]
                                if chosen:
                                    cols[3] = chosen
                                    changed += 1
                                    oc_guess_count[chosen] += 1
                                    context_logs.append(
                                        f"[C] 行{data_start + idx + 2} → {chosen}（含别名匹配）"
                                    )
                                else:
                                    unguessed_count += 1
                            except Exception:
                                unguessed_count += 1
                        else:
                            unguessed_count += 1
                    else:
                        unguessed_count += 1

            new_lines.append('\t'.join(cols) + '\n')

        with open(global_result_path, "w", encoding="utf-8") as f:
            for line in new_lines:
                f.write(line if line.endswith('\n') else line + '\n')

        if empty_inbound == 0:
            summary = (
                f"收件人：没有需要补全的空缺"
                f"（已检查 {already_had} 封收信，收件人栏都已有值；"
            )
        elif changed == 0 and unguessed_count == 0:
            summary = "收件人：没有改动。"
        else:
            detail = "\n".join([f"{oc}：补全 {oc_guess_count[oc]} 封" for oc in oc_names])
            summary = (
                f"收件人：在 {empty_inbound} 封空缺里补全了 {changed} 封"
                + (f"。\n{detail}" if detail else "")
                + (f"\n还有 {unguessed_count} 封空着、认不出来，可到「库存」里手工改。" if unguessed_count else "。")
            )
        log_text = summary + "\n" + "\n".join(context_logs)
        self.add_work_log(log_text)
        if notify:
            messagebox.showinfo("智能猜测完成", summary)
        return summary

    # ---------- 壳 ----------
    def create_widgets(self):
        main = ttk.Frame(self.root, padding="10")
        main.pack(fill=tk.BOTH, expand=True)

        self.notebook = ttk.Notebook(main)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        self.login_frame = ttk.Frame(self.notebook, padding="8")
        self.import_frame = ttk.Frame(self.notebook, padding="8")
        self.inventory_frame = ttk.Frame(self.notebook, padding="8")
        self.advanced_frame = ttk.Frame(self.notebook, padding="8")
        self.notebook.add(self.login_frame, text="登录")
        self.notebook.add(self.import_frame, text="检索入库")
        self.notebook.add(self.inventory_frame, text="信件库存")
        self.notebook.add(self.advanced_frame, text="高级设置")

        self._init_vars()
        self._build_login_tab()
        self._build_import_tab()
        self._build_inventory_tab()
        self._build_advanced_tab()

        status = ttk.Frame(self.root, relief=tk.SUNKEN, padding=(2, 2))
        status.pack(fill=tk.X, side=tk.BOTTOM)
        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(status, textvariable=self.status_var, anchor=tk.W).pack(side=tk.LEFT, fill=tk.X)
        ttk.Label(status, text="v1.3", anchor=tk.E).pack(side=tk.RIGHT)

    def _init_vars(self):
        self.email_var = tk.StringVar(value=self.downloader.email_addr)
        self.password_var = tk.StringVar(value=self.downloader.password)
        self.server_var = tk.StringVar(value=self.downloader.imap_server)
        preset = "其他"
        for name, host in self.MAIL_PRESETS.items():
            if host and host == self.downloader.imap_server:
                preset = name
                break
        self.mailbox_type_var = tk.StringVar(value=preset)

        # 默认始终为程序所在目录（开发=项目根，打包=exe 同级）
        default_root = script_dir.resolve()
        default_email = default_root / "email"
        self.root_dir_var = tk.StringVar(value=str(default_root))
        self.downloader.base_dir = default_email

        stored = self.downloader.config.get("Storage", "base_directory", fallback="")
        if stored:
            stored_path = Path(stored).expanduser()
            # 旧版错误默认：~/email；视为无效，继续用程序目录
            legacy_home = (Path.home() / "email").resolve()
            try:
                is_legacy = stored_path.resolve() == legacy_home
            except Exception:
                is_legacy = False
            if not is_legacy and stored_path.name.lower() == "email":
                self.root_dir_var.set(str(stored_path.parent.resolve()))
                self.downloader.base_dir = stored_path
            elif not is_legacy and stored_path.name.lower() != "email":
                # 配置里直接写了根目录时也兼容
                self.root_dir_var.set(str(stored_path.resolve()))
                self.downloader.base_dir = stored_path / "email"
        self.timeout_var = tk.StringVar(value=str(self.downloader.timeout))
        self.batch_var = tk.StringVar(value=str(self.downloader.batch_size))
        self.imap_port_var = tk.StringVar(value=str(getattr(self.downloader, "imap_port", 993)))

        rule1 = self.downloader.get_default_rule(1)
        rule2 = self.downloader.get_default_rule(2)
        self.adv_rule1_from = tk.StringVar(value=rule1.get("from", ""))
        self.adv_rule1_to = tk.StringVar(value=rule1.get("to", "") or self.downloader.email_addr)
        self.adv_rule1_subject = tk.StringVar(value=rule1.get("subject_pattern", ""))
        self.adv_rule2_from = tk.StringVar(value=rule2.get("from", "") or self.downloader.email_addr)
        self.adv_rule2_to = tk.StringVar(value=rule2.get("to", ""))
        self.adv_rule2_subject = tk.StringVar(value=rule2.get("subject_pattern", ""))

        self.oc_name_var = tk.StringVar()
        self.oc_date_var = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        self.alias_input_var = tk.StringVar()
        self.folder_var = tk.StringVar()
        self.start_date_var = tk.StringVar()
        self.end_date_var = tk.StringVar()
        self.work_status_var = tk.StringVar(value="登录后，在这里登记角色并拉取来往信件。")
        self.summary_var = tk.StringVar(value="")
        self.inv_filter_var = tk.StringVar()

    def _scrollable(self, parent):
        holder = ttk.Frame(parent)
        holder.pack(fill=tk.BOTH, expand=True)
        canvas = tk.Canvas(holder, highlightthickness=0, borderwidth=0)
        vsb = ttk.Scrollbar(holder, orient="vertical", command=canvas.yview)
        inner = ttk.Frame(canvas, padding=(2, 2))
        iid = canvas.create_window((0, 0), window=inner, anchor="nw")

        def on_inner(_e=None):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def on_canvas(event):
            canvas.itemconfigure(iid, width=event.width)

        inner.bind("<Configure>", on_inner)
        canvas.bind("<Configure>", on_canvas)
        canvas.configure(yscrollcommand=vsb.set)

        def wheel(event):
            w = event.widget
            try:
                cls = w.winfo_class()
            except Exception:
                cls = ""
            # 这些控件自己处理滚轮，不要滚外层
            if cls in ("Text", "Listbox", "Treeview"):
                return
            if canvas.winfo_exists():
                canvas.yview_scroll(int(-event.delta / 120), "units")
            return "break"

        def _wheel_enter(_event=None):
            holder.unbind_all("<MouseWheel>")
            holder.bind_all("<MouseWheel>", wheel)

        def _wheel_leave(event=None):
            # 鼠标仍在 holder 内（含嵌入区子控件）时不解绑
            if event is not None:
                try:
                    x, y = holder.winfo_pointerxy()
                    w = holder.winfo_containing(x, y)
                    if w is not None:
                        cur = w
                        while cur:
                            if cur == holder:
                                return
                            cur = cur.master
                except Exception:
                    pass
            holder.unbind_all("<MouseWheel>")

        holder.bind("<Enter>", _wheel_enter)
        holder.bind("<Leave>", _wheel_leave)
        canvas.bind("<MouseWheel>", wheel)
        inner.bind("<MouseWheel>", wheel)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        return inner

    def _make_date_row(self, parent, textvariable, width=12):
        """输入框 + 日历按钮（弹独立窗口，避免 DateEntry 在滚动区一点就关）。"""
        row = ttk.Frame(parent)
        ttk.Entry(row, textvariable=textvariable, width=width).pack(side=tk.LEFT)
        ttk.Button(row, text="日历", width=4, command=lambda: self._pick_date_dialog(textvariable)).pack(
            side=tk.LEFT, padx=4
        )
        return row

    def _pick_date_dialog(self, textvariable):
        top = tk.Toplevel(self.root)
        top.title("选择日期")
        top.transient(self.root)
        top.resizable(False, False)
        try:
            top.grab_set()
        except Exception:
            pass
        cal_kw = {"selectmode": "day", "date_pattern": "yyyy-mm-dd"}
        raw = (textvariable.get() or "").strip()
        if raw:
            try:
                d = datetime.strptime(raw, "%Y-%m-%d")
                cal_kw.update(year=d.year, month=d.month, day=d.day)
            except Exception:
                pass
        cal = Calendar(top, **cal_kw)
        cal.pack(padx=10, pady=10)

        def confirm():
            textvariable.set(cal.get_date())
            top.destroy()

        bar = ttk.Frame(top)
        bar.pack(pady=(0, 10))
        ttk.Button(bar, text="确定", command=confirm).pack(side=tk.LEFT, padx=6)
        ttk.Button(bar, text="取消", command=top.destroy).pack(side=tk.LEFT)
        top.wait_window()

    def initialize_app(self):
        self.load_oc_from_tsv()
        self.refresh_inventory()
        if self.downloader.email_addr and self.downloader.password and self.downloader.imap_server:
            self.status_var.set("已填入上次的邮箱，确认授权码后可以登录")

    # ==================== 登录 ====================
    def _build_login_tab(self):
        inner = self._scrollable(self.login_frame)

        guide = ttk.LabelFrame(inner, text="如何获取 IMAP 授权码", padding=(10, 8))
        guide.pack(fill=tk.X, pady=4)
        guide_text = (
            "1. 打开你的邮箱网页版 → 设置，搜索「IMAP」或「授权码」。\n"
            "2. 开启 IMAP 服务，按提示用短信等方式生成授权码（为了信息安全，请不尽量避免暴露授权码！）。\n"
            "3. 如果授权码有日期限制，请选「全部」，否则授权码到期后需要重复上述操作。\n"
            "4. 把邮箱地址和授权码填到下面，选好邮箱类型，点「登录」。"
        )
        ttk.Label(guide, text=guide_text, justify=tk.LEFT, wraplength=500).pack(anchor=tk.W)

        help_frame = ttk.LabelFrame(inner, text="帮助", padding=(10, 8))
        help_frame.pack(fill=tk.X, pady=4)
        help_text = LOGIN_EXTRA_HELP.strip() or LOGIN_HELP_DEFAULT
        ttk.Label(help_frame, text=help_text, justify=tk.LEFT, wraplength=500).pack(anchor=tk.W)

        form = ttk.LabelFrame(inner, text="登录", padding=(10, 8))
        form.pack(fill=tk.X, pady=8)
        form.columnconfigure(1, weight=1)

        ttk.Label(form, text="邮箱地址").grid(row=0, column=0, sticky=tk.W, pady=6, padx=(0, 8))
        ttk.Entry(form, textvariable=self.email_var).grid(row=0, column=1, sticky=tk.EW, pady=6)

        ttk.Label(form, text="授权码").grid(row=1, column=0, sticky=tk.W, pady=6, padx=(0, 8))
        ttk.Entry(form, textvariable=self.password_var, show="*").grid(row=1, column=1, sticky=tk.EW, pady=6)

        ttk.Label(form, text="邮箱类型").grid(row=2, column=0, sticky=tk.W, pady=6, padx=(0, 8))
        self.mailbox_combo = ttk.Combobox(
            form, textvariable=self.mailbox_type_var,
            values=list(self.MAIL_PRESETS.keys()), state="readonly"
        )
        self.mailbox_combo.grid(row=2, column=1, sticky=tk.EW, pady=6)
        self.mailbox_combo.bind("<<ComboboxSelected>>", self._on_mailbox_type)

        self.server_row = ttk.Frame(form)
        ttk.Label(self.server_row, text="IMAP 服务器").pack(side=tk.LEFT, padx=(0, 8))
        ttk.Entry(self.server_row, textvariable=self.server_var, width=36).pack(side=tk.LEFT, fill=tk.X, expand=True)
        self._on_mailbox_type()

        footer = ttk.Frame(self.login_frame)
        footer.pack(fill=tk.X, pady=(8, 0))
        ttk.Button(footer, text="登录", command=self.do_login).pack(side=tk.RIGHT)

    def _on_mailbox_type(self, _event=None):
        name = self.mailbox_type_var.get()
        host = self.MAIL_PRESETS.get(name, "")
        if host:
            self.server_var.set(host)
            self.server_row.grid_remove()
        else:
            self.server_row.grid(row=3, column=0, columnspan=2, sticky=tk.EW, pady=6)

    def do_login(self):
        if not self.email_var.get().strip() or not self.password_var.get():
            messagebox.showwarning("提示", "请填写邮箱地址和授权码。")
            return
        self._on_mailbox_type()
        if not self.server_var.get().strip():
            messagebox.showwarning("提示", "请选择邮箱类型，或在「其他」里填写 IMAP 服务器。")
            return
        self.downloader.email_addr = self.email_var.get().strip()
        self.downloader.password = self.password_var.get()
        self.downloader.imap_server = self.server_var.get().strip()
        try:
            self.downloader.timeout = int(self.timeout_var.get())
            self.downloader.batch_size = int(self.batch_var.get())
            self.downloader.imap_port = int(self.imap_port_var.get())
            root_path = Path(self.root_dir_var.get())
            root_path.mkdir(parents=True, exist_ok=True)
            self.downloader.base_dir = root_path / "email"
            self.downloader.base_dir.mkdir(parents=True, exist_ok=True)
            self.downloader.save_config()
        except Exception as e:
            messagebox.showerror("错误", f"保存设置失败: {e}")
            return
        self.status_var.set("正在登录…")
        self.root.update()
        if not self.downloader.connect_to_email():
            messagebox.showerror("错误", "登录失败，请检查邮箱、授权码和邮箱类型。")
            self.status_var.set("登录失败")
            return
        folders = self.downloader.get_all_mailboxes()
        if folders:
            self.folder_combo["values"] = folders
            preferred = "INBOX" if "INBOX" in folders else folders[0]
            self.folder_var.set(preferred)
            self.selected_folder = preferred
        else:
            messagebox.showerror("错误", "登录成功，但没有找到收件箱。")
            return
        self.status_var.set(f"已登录，默认文件夹 {self.selected_folder}")
        self.notebook.select(self.import_frame)

    # ==================== 检索入库（子页：OC 管理 / 信件下载）====================
    def _build_import_tab(self):
        self.import_notebook = ttk.Notebook(self.import_frame)
        self.import_notebook.pack(fill=tk.BOTH, expand=True)

        self.oc_page = ttk.Frame(self.import_notebook, padding="4")
        self.download_page = ttk.Frame(self.import_notebook, padding="4")
        self.import_notebook.add(self.oc_page, text="OC 管理")
        self.import_notebook.add(self.download_page, text="信件下载")

        self._build_oc_page()
        self._build_download_page()

    def _build_oc_page(self):
        inner = self._scrollable(self.oc_page)

        oc_box = ttk.LabelFrame(inner, text="角色（OC）", padding=(8, 6))
        oc_box.pack(fill=tk.BOTH, expand=True, pady=4)
        ttk.Label(
            oc_box,
            text="登记你的OC名和OC第一次使用地缝的日期。如果一个OC有多个称呼[别名]，请在此注册",
            wraplength=500,
        ).pack(anchor=tk.W, pady=(0, 4))

        cols = ("名字", "首次使用日期", "别名")
        self.oc_tree = ttk.Treeview(oc_box, columns=cols, show="headings", height=8, selectmode="browse")
        for c, w in zip(cols, (120, 120, 240)):
            self.oc_tree.heading(c, text=c)
            self.oc_tree.column(c, width=w)
        self.oc_tree.pack(fill=tk.BOTH, expand=True, pady=4)
        self.oc_tree.bind("<<TreeviewSelect>>", self._on_oc_select)

        form = ttk.Frame(oc_box)
        form.pack(fill=tk.X, pady=4)
        ttk.Label(form, text="名字").grid(row=0, column=0, sticky=tk.W, padx=(0, 4))
        ttk.Entry(form, textvariable=self.oc_name_var, width=16).grid(row=0, column=1, padx=4)
        ttk.Label(form, text="日期").grid(row=0, column=2, sticky=tk.W, padx=(8, 4))
        self._make_date_row(form, self.oc_date_var, width=12).grid(row=0, column=3, padx=4, sticky=tk.W)

        alias_box = ttk.LabelFrame(oc_box, text="别名", padding=(6, 4))
        alias_box.pack(fill=tk.X, pady=4)
        alias_row = ttk.Frame(alias_box)
        alias_row.pack(fill=tk.X)
        ttk.Entry(alias_row, textvariable=self.alias_input_var, width=24).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(alias_row, text="添加别名", command=self.add_alias_item).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(alias_row, text="删除所选别名", command=self.remove_alias_item).pack(side=tk.LEFT)
        self.alias_listbox = tk.Listbox(alias_box, height=4, exportselection=False)
        self.alias_listbox.pack(fill=tk.X, pady=4)

        btns = ttk.Frame(oc_box)
        btns.pack(fill=tk.X, pady=2)
        ttk.Button(btns, text="添加 / 更新角色", command=self.save_oc_row).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btns, text="删除所选角色", command=self.delete_oc_row).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btns, text="清空表单", command=self.clear_oc_form).pack(side=tk.LEFT)

        foot = ttk.Frame(self.oc_page)
        foot.pack(fill=tk.X, pady=(8, 0))
        ttk.Button(foot, text="下一步：信件下载", command=self._goto_download_page).pack(side=tk.RIGHT)

    def _build_download_page(self):
        inner = self._scrollable(self.download_page)

        pull = ttk.LabelFrame(inner, text="拉取来往信件", padding=(8, 6))
        pull.pack(fill=tk.X, pady=6)
        ttk.Label(
            pull,
            text="在选定邮箱文件夹里搜与地缝的来往信。默认收件箱；若找不到信件，请确保文件夹正确。",
            wraplength=500,
        ).pack(anchor=tk.W)
        fr = ttk.Frame(pull)
        fr.pack(fill=tk.X, pady=4)
        ttk.Label(fr, text="邮箱文件夹").pack(side=tk.LEFT)
        self.folder_combo = ttk.Combobox(fr, textvariable=self.folder_var, width=36, state="readonly")
        self.folder_combo.pack(side=tk.LEFT, padx=6)
        dr = ttk.Frame(pull)
        dr.pack(fill=tk.X, pady=6)
        ttk.Label(dr, text="起始").pack(side=tk.LEFT)
        self._make_date_row(dr, self.start_date_var, width=12).pack(side=tk.LEFT, padx=6)
        ttk.Label(dr, text="结束").pack(side=tk.LEFT, padx=(8, 0))
        self._make_date_row(dr, self.end_date_var, width=12).pack(side=tk.LEFT, padx=6)
        ttk.Button(dr, text="开始拉取", command=self.start_fetch).pack(side=tk.RIGHT)

        ttk.Label(pull, textvariable=self.work_status_var, wraplength=500).pack(anchor=tk.W, pady=2)
        self.work_progress = ttk.Progressbar(pull, orient=tk.HORIZONTAL, mode="determinate")
        self.work_progress.pack(fill=tk.X, pady=4)

        res_box = ttk.LabelFrame(inner, text="将要被下载的来往信件", padding=(6, 6))
        res_box.pack(fill=tk.X, pady=4)
        mcols = ("方向", "日期", "发件邮箱", "主题")
        self.match_tree = ttk.Treeview(res_box, columns=mcols, show="headings", height=8, selectmode="extended")
        for c, w in zip(mcols, (50, 130, 180, 280)):
            self.match_tree.heading(c, text=c)
            self.match_tree.column(c, width=w)
        self.match_tree.pack(fill=tk.X)

        act = ttk.Frame(res_box)
        act.pack(fill=tk.X, pady=6)
        ttk.Button(act, text="处理主题不合规的信…", command=self.open_unmatched_dialog).pack(side=tk.LEFT)
        ttk.Button(act, text="下载入库并打开[地缝可视化]", command=self.start_download_pipeline).pack(side=tk.RIGHT)

        sum_box = ttk.LabelFrame(inner, text="结果摘要", padding=(8, 6))
        sum_box.pack(fill=tk.X, pady=4)
        ttk.Label(sum_box, textvariable=self.summary_var, wraplength=500, justify=tk.LEFT).pack(anchor=tk.W)

        log_box = ttk.LabelFrame(inner, text="详细记录", padding=(6, 6))
        log_box.pack(fill=tk.X, pady=4)
        self.work_log = scrolledtext.ScrolledText(log_box, wrap=tk.WORD, height=8)
        self.work_log.pack(fill=tk.BOTH, expand=True)
        self.work_log.configure(state="disabled")

        foot = ttk.Frame(self.download_page)
        foot.pack(fill=tk.X, pady=(8, 0))
        ttk.Button(foot, text="打开日志文件夹", command=self.open_log_folder).pack(side=tk.LEFT)
        ttk.Button(foot, text="查看下载报告", command=self.view_download_report).pack(side=tk.LEFT, padx=6)
        ttk.Button(foot, text="再次打开[地缝可视化]", command=self.open_visualization_html).pack(side=tk.RIGHT)

    def _goto_download_page(self):
        self.import_notebook.select(self.download_page)

    # ----- OC helpers -----
    def load_oc_from_tsv(self):
        path = self.downloader.base_dir / "global_result.tsv"
        self.downloader.oc_registerdata = []
        self.downloader.oc_aliases = {}
        if path.exists():
            try:
                lines = path.read_text(encoding="utf-8").splitlines()
            except Exception:
                lines = []
            for line in lines:
                if line.startswith("oc_registerdata"):
                    cols = line.split("\t")
                    if len(cols) > 1 and cols[1]:
                        self.downloader.oc_registerdata = [
                            x.strip() for x in cols[1].split(",") if x.strip()
                        ]
                elif line.startswith("oc_aliases"):
                    cols = line.split("\t")
                    if len(cols) > 1 and cols[1]:
                        for group in cols[1].split(","):
                            parts = [p.strip() for p in group.split("=") if p.strip()]
                            if len(parts) > 1:
                                self.downloader.oc_aliases[parts[0]] = parts[1:]
                elif line.startswith("收/发"):
                    break
        self.refresh_oc_tree()

    def refresh_oc_tree(self):
        for item in self.oc_tree.get_children():
            self.oc_tree.delete(item)
        for entry in self.downloader.oc_registerdata:
            if "_" in entry:
                name, date_str = entry.rsplit("_", 1)
            else:
                name, date_str = entry, ""
            aliases = "，".join(self.downloader.oc_aliases.get(name, []))
            self.oc_tree.insert("", tk.END, values=(name, date_str, aliases))

    def _alias_list_get(self):
        return [self.alias_listbox.get(i) for i in range(self.alias_listbox.size())]

    def _alias_list_set(self, aliases):
        self.alias_listbox.delete(0, tk.END)
        for a in aliases or []:
            if a:
                self.alias_listbox.insert(tk.END, a)

    def add_alias_item(self):
        name = self.alias_input_var.get().strip()
        if not name:
            return
        if name in self._alias_list_get():
            self.alias_input_var.set("")
            return
        self.alias_listbox.insert(tk.END, name)
        self.alias_input_var.set("")

    def remove_alias_item(self):
        sel = list(self.alias_listbox.curselection())
        for i in reversed(sel):
            self.alias_listbox.delete(i)

    def _on_oc_select(self, _event=None):
        sel = self.oc_tree.selection()
        if not sel:
            return
        vals = self.oc_tree.item(sel[0], "values")
        self.oc_edit_index = self.oc_tree.index(sel[0])
        self.oc_name_var.set(vals[0])
        self._alias_list_set(self.downloader.oc_aliases.get(vals[0], []))
        self.alias_input_var.set("")
        if vals[1]:
            self.oc_date_var.set(vals[1])

    def clear_oc_form(self):
        self.oc_edit_index = None
        self.oc_name_var.set("")
        self.oc_date_var.set(datetime.now().strftime("%Y-%m-%d"))
        self.alias_input_var.set("")
        self._alias_list_set([])
        self.oc_tree.selection_remove(self.oc_tree.selection())

    def save_oc_row(self):
        name = self.oc_name_var.get().strip()
        if not name:
            messagebox.showwarning("提示", "请填写角色名字。")
            return
        if "_" in name:
            messagebox.showwarning("提示", "角色名里不能有下划线。")
            return
        date_str = self.oc_date_var.get().strip()
        try:
            datetime.strptime(date_str, "%Y-%m-%d")
        except Exception:
            messagebox.showwarning("提示", "请填写或用日历选择日期（格式 YYYY-MM-DD）。")
            return
        # 若输入框里还有未点「添加」的别名，一并收下
        pending = self.alias_input_var.get().strip()
        if pending:
            self.add_alias_item()
        entry = f"{name}_{date_str}"
        aliases = self._alias_list_get()

        # 更新列表：同名覆盖
        new_list = []
        replaced = False
        for old in self.downloader.oc_registerdata:
            old_name = old.split("_")[0]
            if old_name == name:
                new_list.append(entry)
                replaced = True
            else:
                new_list.append(old)
        if not replaced:
            if self.oc_edit_index is not None and 0 <= self.oc_edit_index < len(new_list):
                old_entry = new_list[self.oc_edit_index]
                old_name = old_entry.split("_")[0]
                new_list[self.oc_edit_index] = entry
                if old_name in self.downloader.oc_aliases and old_name != name:
                    self.downloader.oc_aliases.pop(old_name, None)
            else:
                new_list.append(entry)
        self.downloader.oc_registerdata = new_list
        if aliases:
            self.downloader.oc_aliases[name] = aliases
        elif name in self.downloader.oc_aliases:
            del self.downloader.oc_aliases[name]
        self._persist_oc_header()
        self.refresh_oc_tree()
        self.clear_oc_form()
        self.status_var.set(f"已保存角色 {name}")

    def delete_oc_row(self):
        sel = self.oc_tree.selection()
        if not sel:
            messagebox.showinfo("提示", "请先在列表里选一个角色。")
            return
        idx = self.oc_tree.index(sel[0])
        if 0 <= idx < len(self.downloader.oc_registerdata):
            entry = self.downloader.oc_registerdata.pop(idx)
            name = entry.split("_")[0]
            self.downloader.oc_aliases.pop(name, None)
        self._persist_oc_header()
        self.refresh_oc_tree()
        self.clear_oc_form()

    def _persist_oc_header(self):
        path = self.downloader.base_dir / "global_result.tsv"
        if not path.exists():
            return
        lines = path.read_text(encoding="utf-8").splitlines(keepends=True)
        oc_line = "oc_registerdata\t" + ",".join(self.downloader.oc_registerdata) + "\n"
        groups = []
        for main, subs in self.downloader.oc_aliases.items():
            if subs:
                groups.append(main + "=" + "=".join(subs))
        alias_line = ("oc_aliases\t" + ",".join(groups) + "\n") if groups else None

        out = []
        wrote_oc = False
        wrote_alias = False
        for line in lines:
            if line.startswith("oc_registerdata"):
                out.append(oc_line)
                wrote_oc = True
                if alias_line and not wrote_alias:
                    out.append(alias_line)
                    wrote_alias = True
            elif line.startswith("oc_aliases"):
                if alias_line and not wrote_alias:
                    out.append(alias_line)
                    wrote_alias = True
                # 跳过旧别名行
            else:
                out.append(line)
        if not wrote_oc:
            out.insert(0, oc_line)
            if alias_line:
                out.insert(1, alias_line)
        path.write_text("".join(out), encoding="utf-8")

    # ----- 拉取 -----
    def _parse_optional_date(self, var):
        raw = var.get().strip() if hasattr(var, "get") else ""
        if not raw:
            return None
        try:
            return datetime.strptime(raw, "%Y-%m-%d")
        except Exception:
            return None

    def start_fetch(self):
        if self.busy:
            return
        if not self.downloader.client and not self.downloader.connect_to_email():
            messagebox.showerror("错误", "请先在「登录」页登录。")
            self.notebook.select(self.login_frame)
            return
        folder = (self.folder_var.get() or self.selected_folder or "").strip()
        if not folder:
            folders = self.downloader.get_all_mailboxes() or []
            if folders:
                self.folder_combo["values"] = folders
                folder = "INBOX" if "INBOX" in folders else folders[0]
                self.folder_var.set(folder)
        if not folder:
            messagebox.showerror("错误", "没有找到邮箱文件夹。")
            return
        self.selected_folder = folder
        if not self.downloader.oc_registerdata:
            if not messagebox.askyesno("提示", "还没有登记角色。仍要继续拉取吗？"):
                return
        self.busy = True
        self.work_progress["value"] = 0
        self.summary_var.set("")
        self.clear_work_log()
        for item in self.match_tree.get_children():
            self.match_tree.delete(item)
        self.work_status_var.set("正在检索往来信件…")
        self.status_var.set("正在检索往来信件…")
        self.matched_results = []
        self.unmatched_results = []
        self.manual_queue = []
        threading.Thread(target=self._run_fetch, daemon=True).start()

    def _run_fetch(self):
        try:
            folder = self.folder_var.get() or self.selected_folder
            start_date = self._parse_optional_date(self.start_date_var)
            end_date = self._parse_optional_date(self.end_date_var)
            merged = []
            seen = set()
            unmatched_all = []
            unmatched_seen = set()

            for index, (rule_num, label) in enumerate(((1, "收到的信"), (2, "发出的信"))):
                rule = self.downloader.get_default_rule(rule_num)
                base = index * 50

                def cb(progress, found, _base=base, _label=label):
                    def _ui(p=progress, f=found, b=_base, lab=_label):
                        # 检索中始终提示「正在检索」；有命中再附带数量，避免长时间显示「0 封」
                        if f > 0:
                            self.work_status_var.set(
                                f"正在检索往来信件（{lab}）… 已找到 {f} 封"
                            )
                        else:
                            self.work_status_var.set(f"正在检索往来信件（{lab}）…")
                        self.status_var.set("正在检索往来信件…")
                        self.work_progress["value"] = b + p * 0.5
                    self.root.after(0, _ui)

                self.root.after(
                    0,
                    lambda lab=label: self.work_status_var.set(
                        f"正在检索往来信件（{lab}）…"
                    ),
                )

                self.downloader.last_unmatched = []
                results = self.downloader.search_emails_advanced(
                    folder=folder,
                    subject_pattern=rule.get("subject_pattern", ""),
                    from_address=rule.get("from", ""),
                    to_address=rule.get("to", ""),
                    start_date=start_date,
                    end_date=end_date,
                    callback=cb,
                    max_emails=1500,
                )
                for email_info in results:
                    eid = email_info.get("id")
                    if eid not in seen:
                        seen.add(eid)
                        direction = self.downloader.direction_from_addresses(
                            email_info.get("sender", ""),
                            email_info.get("recipients") or [],
                        )
                        # 规则1收、规则2发；再以地址方向覆盖
                        if rule_num == 1:
                            direction = "收"
                        elif rule_num == 2:
                            direction = "发"
                        email_info["direction"] = direction
                        merged.append(email_info)
                for item in getattr(self.downloader, "last_unmatched", []) or []:
                    eid = item.get("id")
                    if eid not in unmatched_seen and eid not in seen:
                        unmatched_seen.add(eid)
                        unmatched_all.append(item)

                # 每一轮规则结束后先刷新列表，避免整段检索期间一直空白
                snap = list(merged)

                def _mid_refresh(s=snap, lab=label, n=len(snap), last=(index == 1)):
                    self.matched_results = s
                    self._fill_match_tree()
                    if last:
                        self.work_status_var.set("正在检索往来信件… 汇总中")
                    else:
                        self.work_status_var.set(
                            f"正在检索往来信件…（{lab}已纳入 {n} 封，继续检索中）"
                        )

                self.root.after(0, _mid_refresh)

            self.matched_results = merged
            self.unmatched_results = unmatched_all
            self.downloader.search_results = list(merged)

            def done():
                self._fill_match_tree()
                self.work_progress["value"] = 100
                self.work_status_var.set(
                    f"找到 {len(merged)} 封合规来往信；另有 {len(unmatched_all)} 封主题不合规。"
                )
                self.status_var.set(
                    f"检索完成：{len(merged)} 封合规 / {len(unmatched_all)} 封待处理"
                )
                self.add_work_log(
                    f"拉取完成：匹配 {len(merged)}，未匹配 {len(unmatched_all)}\n"
                )
                if unmatched_all:
                    if messagebox.askyesno(
                        "主题不合规",
                        f"有 {len(unmatched_all)} 封来往信主题不符合命名规则。\n是否现在手工指定为 A/B/C 信？",
                    ):
                        self.open_unmatched_dialog()
                self.busy = False

            self.root.after(0, done)
        except Exception as e:
            logger.error(f"拉取失败: {e}")
            self.root.after(0, lambda err=e: messagebox.showerror("错误", f"拉取失败: {err}"))
            self.root.after(0, lambda: self.work_status_var.set("拉取失败"))
            self.root.after(0, lambda: setattr(self, "busy", False))

    def _fill_match_tree(self):
        for item in self.match_tree.get_children():
            self.match_tree.delete(item)
        for email_info in self.matched_results + self.manual_queue:
            try:
                if email_info.get("date_obj"):
                    date_str = email_info["date_obj"].strftime("%Y-%m-%d %H:%M")
                else:
                    date_str = (email_info.get("date") or "")[:19] or "未知"
            except Exception:
                date_str = str(email_info.get("date", "未知"))
            subject = email_info.get("subject", "")
            if len(subject) > 48:
                subject = subject[:45] + "..."
            tag = "manual" if email_info.get("user_letter_type") else "match"
            self.match_tree.insert(
                "",
                tk.END,
                values=(
                    email_info.get("direction") or email_info.get("user_direction") or "?",
                    date_str,
                    email_info.get("sender", ""),
                    subject,
                ),
                tags=(tag,),
            )

    # ----- 未匹配手工补录 -----
    def open_unmatched_dialog(self):
        if not self.unmatched_results:
            messagebox.showinfo("提示", "当前没有主题不合规的来往信。")
            return

        win = tk.Toplevel(self.root)
        win.title("主题不合规 — 手工指定")
        win.geometry("720x480")
        win.transient(self.root)

        ttk.Label(
            win,
            text="这些信的收发地址像地缝来往信，但主题不合规。收/发由邮箱方向自动决定；请补信件类型和收发件人。",
            wraplength=680,
        ).pack(anchor=tk.W, padx=10, pady=8)

        cols = ("方向", "日期", "发件邮箱", "主题")
        tree = ttk.Treeview(win, columns=cols, show="headings", height=10, selectmode="browse")
        for c, w in zip(cols, (50, 130, 180, 300)):
            tree.heading(c, text=c)
            tree.column(c, width=w)
        tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=4)

        id_map = {}
        for i, item in enumerate(self.unmatched_results):
            try:
                if item.get("date_obj"):
                    date_str = item["date_obj"].strftime("%Y-%m-%d %H:%M")
                else:
                    date_str = (item.get("date") or "")[:19]
            except Exception:
                date_str = ""
            iid = tree.insert(
                "",
                tk.END,
                values=(item.get("direction", "?"), date_str, item.get("sender", ""), item.get("subject", "")),
            )
            id_map[iid] = i

        form = ttk.LabelFrame(win, text="为所选信件填写", padding=(8, 6))
        form.pack(fill=tk.X, padx=10, pady=6)
        type_var = tk.StringVar()
        sender_var = tk.StringVar()
        receiver_var = tk.StringVar()
        dir_var = tk.StringVar(value="")
        ttk.Label(form, text="方向（自动）").grid(row=0, column=0, sticky=tk.W, pady=3)
        ttk.Label(form, textvariable=dir_var).grid(row=0, column=1, sticky=tk.W, pady=3)
        ttk.Label(form, text="信件类型").grid(row=1, column=0, sticky=tk.W, pady=3)
        ttk.Entry(form, textvariable=type_var, width=20).grid(row=1, column=1, sticky=tk.W, pady=3)
        ttk.Label(form, text="例：A123 / B45 / C某某").grid(row=1, column=2, sticky=tk.W, padx=6)
        ttk.Label(form, text="发件人").grid(row=2, column=0, sticky=tk.W, pady=3)
        ttk.Entry(form, textvariable=sender_var, width=20).grid(row=2, column=1, sticky=tk.W, pady=3)
        ttk.Label(form, text="收件人").grid(row=3, column=0, sticky=tk.W, pady=3)
        ttk.Entry(form, textvariable=receiver_var, width=20).grid(row=3, column=1, sticky=tk.W, pady=3)

        def on_sel(_e=None):
            sel = tree.selection()
            if not sel:
                return
            item = self.unmatched_results[id_map[sel[0]]]
            dir_var.set(item.get("direction", ""))
            # 收：收件人常为自己的 OC；发：发件人常为自己的 OC
            if item.get("direction") == "收" and self.downloader.oc_registerdata:
                receiver_var.set(self.downloader.oc_registerdata[0].split("_")[0])
            elif item.get("direction") == "发" and self.downloader.oc_registerdata:
                sender_var.set(self.downloader.oc_registerdata[0].split("_")[0])

        tree.bind("<<TreeviewSelect>>", on_sel)

        def add_selected():
            sel = tree.selection()
            if not sel:
                messagebox.showwarning("提示", "请先选一封信。", parent=win)
                return
            letter_type = type_var.get().strip()
            sname = sender_var.get().strip()
            rname = receiver_var.get().strip()
            if not letter_type or not sname or not rname:
                messagebox.showwarning("提示", "请填写信件类型、发件人和收件人。", parent=win)
                return
            idx = id_map[sel[0]]
            item = dict(self.unmatched_results[idx])
            item["user_letter_type"] = letter_type
            item["user_sender_name"] = sname
            item["user_receiver_name"] = rname
            item["user_direction"] = item.get("direction") or dir_var.get()
            # 避免重复加入
            if any(x.get("id") == item.get("id") for x in self.manual_queue):
                messagebox.showinfo("提示", "这封已经在待下载列表里了。", parent=win)
                return
            self.manual_queue.append(item)
            # 从待处理列表移除
            self.unmatched_results.pop(idx)
            tree.delete(sel[0])
            # 重建 id_map
            id_map.clear()
            for iid in tree.get_children():
                # values 无法还原 index；按剩余 unmatched 顺序重绑
                pass
            # 简单重建：清空重填
            for iid in tree.get_children():
                tree.delete(iid)
            id_map.clear()
            for i, u in enumerate(self.unmatched_results):
                try:
                    if u.get("date_obj"):
                        ds = u["date_obj"].strftime("%Y-%m-%d %H:%M")
                    else:
                        ds = (u.get("date") or "")[:19]
                except Exception:
                    ds = ""
                iid = tree.insert(
                    "", tk.END,
                    values=(u.get("direction", "?"), ds, u.get("sender", ""), u.get("subject", "")),
                )
                id_map[iid] = i
            self._fill_match_tree()
            self.work_status_var.set(
                f"已手工加入 {len(self.manual_queue)} 封；待处理不合规 {len(self.unmatched_results)} 封。"
            )
            type_var.set("")
            messagebox.showinfo("已加入", "已加入待下载列表。", parent=win)

        bar = ttk.Frame(win)
        bar.pack(fill=tk.X, padx=10, pady=8)
        ttk.Button(bar, text="加入待下载", command=add_selected).pack(side=tk.LEFT)
        ttk.Button(bar, text="关闭", command=win.destroy).pack(side=tk.RIGHT)

    # ----- 下载入库 -----
    def start_download_pipeline(self):
        if self.busy:
            return
        queue = list(self.matched_results) + list(self.manual_queue)
        if not queue:
            messagebox.showinfo("提示", "还没有可下载的信件，请先拉取。")
            return
        if not self.downloader.client and not self.downloader.connect_to_email():
            messagebox.showerror("错误", "请先登录。")
            return
        self.busy = True
        self.work_progress["value"] = 0
        self.work_status_var.set(f"准备下载 {len(queue)} 封…")
        self.clear_work_log()
        threading.Thread(target=self._run_download_pipeline, args=(queue,), daemon=True).start()

    def _run_download_pipeline(self, queue):
        try:
            def progress(p, current, total):
                def _ui(pv=p, c=current, t=total):
                    self.work_progress["value"] = pv * 0.85
                    self.work_status_var.set(f"下载中… {c}/{t}")
                self.root.after(0, _ui)
                if current <= len(queue):
                    info = queue[current - 1]
                    status = info.get("status", "")
                    subject = info.get("subject", "")
                    if status == "skipped":
                        msg = f"[{current}/{total}] 跳过: {subject}\n"
                    elif status == "failed":
                        msg = f"[{current}/{total}] 失败: {subject}\n"
                    elif status == "success":
                        msg = f"[{current}/{total}] 已下载: {subject}\n"
                    else:
                        msg = f"[{current}/{total}] 处理中: {subject}\n"
                    self.root.after(0, lambda m=msg: self.add_work_log(m))

            results = self.downloader.download_multiple_emails(queue, progress)
            success = len([r for r in results if r.get("status") == "success"])
            skipped = len([r for r in results if r.get("status") == "skipped"])
            failed = len([r for r in results if r.get("status") == "failed"])

            def finish():
                try:
                    self._persist_oc_header()
                    # 确保别名进 downloader，供报告写入
                    self.work_status_var.set("正在识别收件人…")
                    self.work_progress["value"] = 90
                    guess = self.smart_guess_receivers(notify=False) or ""
                    self.summary_var.set(
                        f"找到并处理 {len(queue)} 封。新下载 {success}，跳过 {skipped}，失败 {failed}。\n\n{guess}"
                    )
                    self.work_progress["value"] = 100
                    self.work_status_var.set("入库完成，正在打开[地缝可视化]")
                    self.add_work_log("\n入库完成。\n")
                    self.refresh_inventory()
                    if self.open_visualization_html():
                        self.work_status_var.set("入库完成，[地缝可视化]已打开")
                finally:
                    self.busy = False

            self.root.after(0, finish)
        except Exception as e:
            logger.error(f"入库失败: {e}")
            self.root.after(0, lambda err=e: messagebox.showerror("错误", f"入库失败: {err}"))
            self.root.after(0, lambda: self.work_status_var.set("入库失败"))
            self.root.after(0, lambda: setattr(self, "busy", False))

    def add_work_log(self, text):
        self.work_log.configure(state="normal")
        self.work_log.insert(tk.END, text)
        self.work_log.see(tk.END)
        self.work_log.configure(state="disabled")

    def clear_work_log(self):
        self.work_log.configure(state="normal")
        self.work_log.delete(1.0, tk.END)
        self.work_log.configure(state="disabled")

    def open_log_folder(self):
        try:
            path = str(log_dir)
            if sys.platform.startswith("win"):
                os.startfile(path)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", path])
            else:
                subprocess.Popen(["xdg-open", path])
        except Exception as e:
            messagebox.showerror("错误", f"无法打开日志文件夹: {e}")

    def view_download_report(self):
        report_path = self.downloader.base_dir / "download_report.txt"
        if not report_path.exists():
            messagebox.showinfo("提示", "还没有下载报告。")
            return
        try:
            text = report_path.read_text(encoding="utf-8")
            dlg = tk.Toplevel(self.root)
            dlg.title("下载报告")
            dlg.geometry("700x500")
            box = scrolledtext.ScrolledText(dlg, wrap=tk.WORD)
            box.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            box.insert(tk.END, text)
            box.configure(state="disabled")
            ttk.Button(dlg, text="关闭", command=dlg.destroy).pack(pady=8)
            dlg.transient(self.root)
        except Exception as e:
            messagebox.showerror("错误", f"无法打开下载报告: {e}")

    def open_visualization_html(self):
        html_path = script_dir / "visualization_private.html"
        if not html_path.exists():
            messagebox.showerror("错误", f"未找到可视化文件:\n{html_path}")
            return False
        webbrowser.open_new_tab("file:///" + str(html_path).replace("\\", "/"))
        return True

    # ==================== 高级设置 ====================
    def _build_advanced_tab(self):
        inner = self._scrollable(self.advanced_frame)

        ttk.Label(
            inner,
            text="存储、收信解析规则相关的设置",
            wraplength=500,
        ).pack(anchor=tk.W, pady=(0, 8))

        path_box = ttk.LabelFrame(inner, text="存储与连接", padding=(8, 6))
        path_box.pack(fill=tk.X, pady=4)
        path_box.columnconfigure(1, weight=1)
        ttk.Label(path_box, text="根目录").grid(row=0, column=0, sticky=tk.W, pady=4, padx=(0, 8))
        ttk.Entry(path_box, textvariable=self.root_dir_var).grid(row=0, column=1, sticky=tk.EW, pady=4)
        dir_btns = ttk.Frame(path_box)
        dir_btns.grid(row=0, column=2, padx=4)
        ttk.Button(dir_btns, text="浏览…", command=self._browse_root_dir).pack(side=tk.LEFT, padx=2)
        ttk.Button(dir_btns, text="恢复默认", command=lambda: self.root_dir_var.set(str(script_dir))).pack(side=tk.LEFT)
        ttk.Label(path_box, text="信件保存在：根目录\\email\\", foreground="gray").grid(
            row=1, column=0, columnspan=3, sticky=tk.W
        )

        ttk.Label(path_box, text="IMAP 服务器").grid(row=2, column=0, sticky=tk.W, pady=4, padx=(0, 8))
        ttk.Entry(path_box, textvariable=self.server_var).grid(row=2, column=1, sticky=tk.EW, pady=4)
        ttk.Label(path_box, text="端口").grid(row=3, column=0, sticky=tk.W, pady=4, padx=(0, 8))
        ttk.Entry(path_box, textvariable=self.imap_port_var, width=8).grid(row=3, column=1, sticky=tk.W, pady=4)
        ttk.Label(path_box, text="连接超时（秒）").grid(row=4, column=0, sticky=tk.W, pady=4, padx=(0, 8))
        ttk.Entry(path_box, textvariable=self.timeout_var, width=8).grid(row=4, column=1, sticky=tk.W, pady=4)
        ttk.Label(path_box, text="批处理大小").grid(row=5, column=0, sticky=tk.W, pady=4, padx=(0, 8))
        ttk.Entry(path_box, textvariable=self.batch_var, width=8).grid(row=5, column=1, sticky=tk.W, pady=4)

        rule_box = ttk.LabelFrame(inner, text="检索规则（正则 / 地址）", padding=(8, 6))
        rule_box.pack(fill=tk.X, pady=6)
        rule_box.columnconfigure(1, weight=1)
        ttk.Label(rule_box, text="规则1·收信 发件").grid(row=0, column=0, sticky=tk.W, pady=3)
        ttk.Entry(rule_box, textvariable=self.adv_rule1_from).grid(row=0, column=1, sticky=tk.EW, pady=3)
        ttk.Label(rule_box, text="规则1·收信 收件").grid(row=1, column=0, sticky=tk.W, pady=3)
        ttk.Entry(rule_box, textvariable=self.adv_rule1_to).grid(row=1, column=1, sticky=tk.EW, pady=3)
        ttk.Label(rule_box, text="规则1·主题正则").grid(row=2, column=0, sticky=tk.W, pady=3)
        ttk.Entry(rule_box, textvariable=self.adv_rule1_subject).grid(row=2, column=1, sticky=tk.EW, pady=3)
        ttk.Label(rule_box, text="规则2·发信 发件").grid(row=3, column=0, sticky=tk.W, pady=3)
        ttk.Entry(rule_box, textvariable=self.adv_rule2_from).grid(row=3, column=1, sticky=tk.EW, pady=3)
        ttk.Label(rule_box, text="规则2·发信 收件").grid(row=4, column=0, sticky=tk.W, pady=3)
        ttk.Entry(rule_box, textvariable=self.adv_rule2_to).grid(row=4, column=1, sticky=tk.EW, pady=3)
        ttk.Label(rule_box, text="规则2·主题正则").grid(row=5, column=0, sticky=tk.W, pady=3)
        ttk.Entry(rule_box, textvariable=self.adv_rule2_subject).grid(row=5, column=1, sticky=tk.EW, pady=3)

        help_box = ttk.LabelFrame(inner, text="智能补全时，发生了什么", padding=(8, 6))
        help_box.pack(fill=tk.X, pady=6)
        help_text = (
            "收信时，缺少足够分清「这封信究竟寄给谁」的规则，因此做了如下猜测。\n"
            "只处理「收」且收件人栏为空的行；已有收件人的不会改。\n"
            "· 只登记了一个角色：必然发给此角色，空缺直接填成该角色。\n"
            "· 多个角色时：\n"
            "  - A / B 信：填所有「注册日期 ≤ 发信日期」的角色（A 信面向所有注册者）。\n"
            "  - C / A+C 信：在正文 content.txt 里检索角色主名及其别名；"
            "命中别名会归到主名，能唯一确定才填主名。\n"
            "· 发信的收件人多半在入库时已从 C 信类型等规则写出；这里主要补收信。\n"
            "因此「新下载了几封」和「补全了几封」不是一回事：收件人已经写好时，"
            "摘要会写「没有需要补全的空缺」。"
        )
        ttk.Label(help_box, text=help_text, justify=tk.LEFT, wraplength=500).pack(anchor=tk.W)

        foot = ttk.Frame(self.advanced_frame)
        foot.pack(fill=tk.X, pady=(8, 0))
        ttk.Button(foot, text="保存高级设置", command=self.save_advanced_settings).pack(side=tk.RIGHT)

    def _browse_root_dir(self):
        from tkinter import filedialog
        directory = filedialog.askdirectory(
            initialdir=self.root_dir_var.get(),
            title="选择根目录（邮件将存在 根目录\\email\\ 下）",
        )
        if directory:
            self.root_dir_var.set(directory)

    def save_advanced_settings(self):
        try:
            self.downloader.timeout = int(self.timeout_var.get())
            self.downloader.batch_size = int(self.batch_var.get())
            self.downloader.imap_port = int(self.imap_port_var.get())
            self.downloader.imap_server = self.server_var.get().strip()
            root_path = Path(self.root_dir_var.get())
            root_path.mkdir(parents=True, exist_ok=True)
            self.downloader.base_dir = root_path / "email"
            self.downloader.base_dir.mkdir(parents=True, exist_ok=True)

            if "DefaultRules" not in self.downloader.config:
                self.downloader.config["DefaultRules"] = {}
            rules = self.downloader.config["DefaultRules"]
            rules["rule1_from"] = self.adv_rule1_from.get().strip()
            rules["rule1_to"] = self.adv_rule1_to.get().strip() or "{self_email}"
            rules["rule1_subject_pattern"] = self.adv_rule1_subject.get().strip()
            rules["rule2_from"] = self.adv_rule2_from.get().strip() or "{self_email}"
            rules["rule2_to"] = self.adv_rule2_to.get().strip()
            rules["rule2_subject_pattern"] = self.adv_rule2_subject.get().strip()

            self.downloader.save_config()
            self.load_oc_from_tsv()
            self.refresh_inventory()
            messagebox.showinfo("已保存", "高级设置已保存。")
        except Exception as e:
            messagebox.showerror("错误", f"保存失败: {e}")

    # ==================== 库存 ====================
    def _build_inventory_tab(self):
        top = ttk.Frame(self.inventory_frame)
        top.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(top, text="已入库信件。在此改动信息。").pack(side=tk.LEFT)
        ttk.Button(top, text="刷新", command=self.refresh_inventory).pack(side=tk.RIGHT)

        cols = ("收/发", "日期", "发件人", "收件人", "信件类型", "路径")
        wrap = ttk.Frame(self.inventory_frame)
        wrap.pack(fill=tk.BOTH, expand=True)
        self.inv_tree = ttk.Treeview(wrap, columns=cols, show="headings", selectmode="browse")
        vsb = ttk.Scrollbar(wrap, orient=tk.VERTICAL, command=self.inv_tree.yview)
        self.inv_tree.configure(yscroll=vsb.set)
        self.inv_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        for c, w in zip(cols, (50, 100, 120, 120, 100, 220)):
            self.inv_tree.heading(c, text=c)
            self.inv_tree.column(c, width=w)
        self.inv_tree.bind("<<TreeviewSelect>>", self._on_inv_select)

        edit = ttk.LabelFrame(self.inventory_frame, text="编辑所选", padding=(8, 6))
        edit.pack(fill=tk.X, pady=6)
        self.inv_sender_var = tk.StringVar()
        self.inv_receiver_var = tk.StringVar()
        self.inv_type_var = tk.StringVar()
        self.inv_meta_var = tk.StringVar(value="")
        ttk.Label(edit, textvariable=self.inv_meta_var).grid(row=0, column=0, columnspan=4, sticky=tk.W, pady=2)
        ttk.Label(edit, text="发件人").grid(row=1, column=0, sticky=tk.W, pady=3)
        ttk.Entry(edit, textvariable=self.inv_sender_var, width=18).grid(row=1, column=1, padx=4)
        ttk.Label(edit, text="收件人").grid(row=1, column=2, sticky=tk.W, padx=(8, 0))
        ttk.Entry(edit, textvariable=self.inv_receiver_var, width=18).grid(row=1, column=3, padx=4)
        ttk.Label(edit, text="信件类型").grid(row=2, column=0, sticky=tk.W, pady=3)
        ttk.Entry(edit, textvariable=self.inv_type_var, width=18).grid(row=2, column=1, padx=4, pady=3)
        ttk.Button(edit, text="保存修改", command=self.save_inventory_row).grid(row=2, column=3, sticky=tk.E, padx=4)

        foot = ttk.Frame(self.inventory_frame)
        foot.pack(fill=tk.X, pady=(4, 0))
        ttk.Button(foot, text="打开该信文件夹", command=self.open_selected_mail_folder).pack(side=tk.LEFT)
        ttk.Button(foot, text="打开邮件总文件夹", command=self.open_storage_folder).pack(side=tk.LEFT, padx=6)
        ttk.Button(foot, text="打开[地缝可视化]", command=self.open_visualization_html).pack(side=tk.RIGHT)

        self._inv_rows = []  # parallel to tree children order

    def refresh_inventory(self):
        path = self.downloader.base_dir / "global_result.tsv"
        for item in self.inv_tree.get_children():
            self.inv_tree.delete(item)
        self._inv_rows = []
        self._inv_header_lines = []
        if not path.exists():
            self.status_var.set("还没有库存文件")
            return
        lines = path.read_text(encoding="utf-8").splitlines()
        data_start = -1
        header = []
        for i, line in enumerate(lines):
            if line.startswith("收/发类型") or line.startswith("收/发"):
                data_start = i
                header = lines[: i + 1]
                break
            header.append(line)
        self._inv_header_lines = header if data_start >= 0 else lines[:2]
        if data_start < 0:
            return
        for line in lines[data_start + 1:]:
            cols = line.split("\t")
            if len(cols) < 6:
                continue
            row = {
                "direction": cols[0],
                "date": cols[1],
                "sender": cols[2],
                "receiver": cols[3],
                "letter_type": cols[4],
                "path": cols[5],
            }
            self._inv_rows.append(row)
            self.inv_tree.insert(
                "",
                tk.END,
                values=(row["direction"], row["date"], row["sender"], row["receiver"], row["letter_type"], row["path"]),
            )
        self.status_var.set(f"库存 {len(self._inv_rows)} 封")

    def _on_inv_select(self, _event=None):
        sel = self.inv_tree.selection()
        if not sel:
            return
        idx = self.inv_tree.index(sel[0])
        if idx >= len(self._inv_rows):
            return
        row = self._inv_rows[idx]
        self.inv_meta_var.set(f"{row['direction']}  {row['date']} ")
        self.inv_sender_var.set(row["sender"])
        self.inv_receiver_var.set(row["receiver"])
        self.inv_type_var.set(row["letter_type"])

    def save_inventory_row(self):
        sel = self.inv_tree.selection()
        if not sel:
            messagebox.showinfo("提示", "请先选一封信。")
            return
        idx = self.inv_tree.index(sel[0])
        if idx >= len(self._inv_rows):
            return
        row = self._inv_rows[idx]
        row["sender"] = self.inv_sender_var.get().strip()
        row["receiver"] = self.inv_receiver_var.get().strip()
        row["letter_type"] = self.inv_type_var.get().strip()
        self.inv_tree.item(
            sel[0],
            values=(row["direction"], row["date"], row["sender"], row["receiver"], row["letter_type"], row["path"]),
        )
        self._write_inventory_tsv()
        messagebox.showinfo("已保存", "修改已写入库存表。")

    def _write_inventory_tsv(self):
        path = self.downloader.base_dir / "global_result.tsv"
        # 保留最新 OC 头
        oc_line = "oc_registerdata\t" + ",".join(self.downloader.oc_registerdata)
        groups = []
        for main, subs in self.downloader.oc_aliases.items():
            if subs:
                groups.append(main + "=" + "=".join(subs))
        lines_out = [oc_line]
        if groups:
            lines_out.append("oc_aliases\t" + ",".join(groups))
        lines_out.append("收/发类型\t发信日期\t发件人\t收件人\t信件类型\t信件下载位置")
        for row in self._inv_rows:
            lines_out.append(
                f"{row['direction']}\t{row['date']}\t{row['sender']}\t{row['receiver']}\t{row['letter_type']}\t{row['path']}"
            )
        path.write_text("\n".join(lines_out) + "\n", encoding="utf-8")

    def open_selected_mail_folder(self):
        sel = self.inv_tree.selection()
        if not sel:
            messagebox.showinfo("提示", "请先选一封信。")
            return
        idx = self.inv_tree.index(sel[0])
        row = self._inv_rows[idx]
        rel = row["path"]
        if rel.startswith(".\\"):
            rel = rel[2:]
        folder = (script_dir / rel.replace("\\", os.sep)).resolve()
        if not folder.exists():
            messagebox.showerror("错误", f"找不到文件夹:\n{folder}")
            return
        try:
            if sys.platform.startswith("win"):
                os.startfile(str(folder))
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(folder)])
            else:
                subprocess.Popen(["xdg-open", str(folder)])
        except Exception as e:
            messagebox.showerror("错误", f"无法打开: {e}")

    def open_storage_folder(self):
        try:
            path = str(self.downloader.base_dir)
            if sys.platform.startswith("win"):
                os.startfile(path)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", path])
            else:
                subprocess.Popen(["xdg-open", path])
        except Exception as e:
            messagebox.showerror("错误", f"无法打开文件夹: {e}")


# 主程序入口
def main():
    """主程序入口"""
    root = tk.Tk()
    app = EmailDownloaderGUI(root)
    
    # 设置窗口关闭时的操作
    def on_closing():
        if messagebox.askokcancel("退出", "确定要退出吗？"):
            # 关闭连接
            if app.downloader.client:
                app.downloader.close_connection()
            root.destroy()
    
    root.protocol("WM_DELETE_WINDOW", on_closing)

    # 运行应用程序
    root.mainloop()

if __name__ == "__main__":
    main()