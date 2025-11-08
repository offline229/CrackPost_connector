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
from tkcalendar import DateEntry  # 需要安装: pip install tkcalendar
import imaplib, ssl, sys

# 待办：
# 1、下载中断处理
# 2、保证tsv正确，done
# 3、网页debug
# 4、2023-02-26 2022-02-26
# 5、附件汇总
# 	统计所有收信
# 	第一步，对所有有明确的发件人 收件人的收信，检索其是否在contenttxt中带有附件 的字样	
# 	若有，则找到
# 处理掉对#的错误支持 135 195
# 允许配置下载位置
# 增加错误处理
# 增加日志
# 不匹配结果显示且仅显示满足收发件邮箱地址条件的，也就是即使不匹配，也必须是收发地址正确但格式不正确的
# 增加未匹配手动加入功能
# 增加html界面的手动打开功能
# 下载时，发件可能是A B C某人 
# 收件可能是
# 在选项卡恢复下载地址，但默认下载地址为根目录下email文件夹
# 看来不能用跨域了，必须要在html里面加一个手动载入

# 打包指南

# python -m venv venv
# venv\Scripts\activate  # Windows
# pip install --upgrade pip
# pip install pyinstaller pillow tkcalendar babel imapclient

# pyinstaller -F -w -i "icon\CrackPost_v1.ico" `
#     --hidden-import babel.numbers `
#     --hidden-import imapclient `
#     --hidden-import PIL.Image `
#     --hidden-import PIL.ImageTk `
#     --hidden-import tkcalendar `
#     download_gui.py

# ==================== 日志配置（全局，仅一次）====================
script_dir = Path(os.path.dirname(os.path.abspath(__file__)))
log_dir = script_dir / "log"
log_dir.mkdir(exist_ok=True, parents=True)

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
        # 获取脚本所在目录，确保配置文件保存在正确位置
        script_dir = Path(os.path.dirname(os.path.abspath(__file__)))
        # 存储目录
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
        
        # 搜索规则 - 默认为 A数字 模式
        self.search_pattern = self.config.get('Filters', 'search_pattern', fallback=r'A\d+')
        
        # 创建基本目录 - 现在指向系统根目录下的email文件夹
        self.base_dir = script_dir / "email"
        self.base_dir.mkdir(exist_ok=True, parents=True)
        
        # 邮箱连接和其他设置
        self.mail = None
        self.mailbox_mapping = {}
        self.client = None
        self.search_results = []
        self.oc_registerdata = []
        
        # 添加超时设置
        self.timeout = int(self.config.get('Connection', 'timeout', fallback='30'))
        self.batch_size = int(self.config.get('Connection', 'batch_size', fallback='100'))

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
        """判断邮件是否已完整下载：
        1. 对于当前要下载的email，解析info，并从 TSV 中查找匹配的已下载记录
        2. 验证该路径下 content.txt 存在且完整
        3. 验证附件（如果有）存在
        4. 决定是否下载
        """
        global_result_path = self.base_dir / "global_result.tsv"
        if not global_result_path.exists():
            return False, "", "TSV 不存在"

        # 1. 读取 TSV，构建已下载邮件的索引
        existing_mails = {}  # key: (sender, date, letter_type), value: directory_path
        try:
            with open(global_result_path, "r", encoding="utf-8") as f:
                lines = f.readlines()
            for line in lines[2:]:  # 跳过前两行（OC 和表头）
                cols = line.strip().split('\t')
                if len(cols) < 6:
                    continue
                sender = cols[2].strip()
                date = cols[1].strip()
                letter_type = cols[4].strip()
                directory = cols[5].strip()
                
                # 构建索引键
                key = (sender, date, letter_type)
                existing_mails[key] = directory
        except Exception as e:
            logger.debug(f"读取 TSV 失败: {e}")
            return False, "", f"读取 TSV 出错: {str(e)}"

        # 2. 从 email_info 中提取关键字段
        current_email = self.email_addr.lower()
        sender_email = (email_info.get('sender') or '').strip()
        
        # 解析发件人名称（从 content.txt 或主题中提取）
        sender_name = ""
        subject = email_info.get('subject', '').replace(' ', '')
        
        # 尝试从主题中提取信件类型
        letter_type = self.extract_letter_type(subject)
        
        # 解析日期
        date_str = email_info.get('date', '')
        try:
            date_obj = email.utils.parsedate_to_datetime(date_str)
            date_short = date_obj.strftime("%Y-%m-%d")
        except:
            date_short = date_str[:10] if len(date_str) >= 10 else ""

        # 3. 如果已有目录，优先从 content.txt 中读取发件人
        email_dir = self._expected_email_dir(email_info)
        content_file = email_dir / "content.txt"
        if content_file.exists():
            try:
                with open(content_file, "r", encoding="utf-8") as cf:
                    content = cf.read()
                    m = re.search(r"发件[人x]=【(.+?)】", content)
                    if m:
                        sender_name = m.group(1).strip()
            except Exception:
                pass
        
        # 如果没有从 content.txt 中提取到，尝试从主题中提取
        if not sender_name:
            m = re.search(r"来自(.+?)的信", subject)
            if m:
                sender_name = m.group(1).strip()
        
        # 兜底：用邮箱地址
        if not sender_name:
            sender_name = sender_email.split('@')[0] if sender_email else "未知"

        # 4. 在 TSV 索引中查找
        key = (sender_name, date_short, letter_type)
        if key not in existing_mails:
            logger.debug(f"TSV 中未找到: {key}")
            return False, str(email_dir), "TSV 中无匹配记录"

        # 5. 找到匹配，验证文件完整性
        tsv_directory = existing_mails[key]
        
        # 转换相对路径为绝对路径
        project_root = Path(os.path.dirname(os.path.abspath(__file__))).resolve()
        if tsv_directory.startswith('.\\'):
            tsv_directory = tsv_directory[2:]
        tsv_dir_path = project_root / tsv_directory.replace('\\', os.sep)
        
        if not tsv_dir_path.exists() or not tsv_dir_path.is_dir():
            logger.debug(f"TSV 记录的目录不存在: {tsv_dir_path}")
            return False, str(tsv_dir_path), "TSV 记录的目录不存在"
        
        content_file = tsv_dir_path / "content.txt"
        if not content_file.exists() or not content_file.is_file():
            logger.debug(f"content.txt 不存在: {tsv_dir_path}")
            return False, str(tsv_dir_path), "content.txt 缺失"
        
        try:
            file_size = content_file.stat().st_size
            if file_size <= 100:
                return False, str(tsv_dir_path), "content.txt 文件过小"
            
            with open(content_file, 'r', encoding='utf-8') as f:
                content = f.read()
                # 检查必要字段
                required_fields = ["主题:", "发件邮箱:", "收件邮箱:", "日期:"]
                missing_fields = [field for field in required_fields if field not in content]
                if missing_fields:
                    return False, str(tsv_dir_path), f"content.txt 缺少字段: {', '.join(missing_fields)}"
            
            # 检查附件（如果 email_info 标记有附件）
            has_attachments = email_info.get('has_attachments', False)
            if has_attachments:
                files_in_dir = [f for f in tsv_dir_path.iterdir() if f.is_file() and f.name != 'content.txt']
                if not files_in_dir:
                    return False, str(tsv_dir_path), "附件缺失"
            
            logger.info(f"邮件已完整下载（TSV验证）: {tsv_dir_path}")
            return True, str(tsv_dir_path), "已完整（TSV验证）"
        except Exception as e:
            logger.debug(f"检查文件时出错: {e}")
            return False, str(tsv_dir_path), f"检查出错: {str(e)}"  

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
            
            config['Filters'] = {
                'search_pattern': r'A\d+',
            }
            
            config['Storage'] = {
                'base_directory': os.path.join(os.path.expanduser('~'), 'email')
            }
            
            config['Connection'] = {
                'timeout': '30',
                'batch_size': '100'
            }
            
            # 默认规则与搜索规则保持一致，可见的只有搜索规则。不可见不可改的是下载和归一化规则
            config['DefaultRules'] = {
                'rule1_subject_pattern': r'^Fw[:：]?(A\d+|A\*\d+|B\d+|C\d+|A\+B\d+|A\+C\d+|A\+B\*\d+|A\+C\*\d+)',
                'rule1_from': 'crackpost2@126.com,crackpost@126.com',
                'rule1_to': '{self_email}',
                
                'rule2_subject_pattern': r'^(A|B|C\+.*|A\+B.*|A\+C.*)$',
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
        
        self.config['Filters'] = {
            'search_pattern': self.search_pattern,
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
        if not self.client:
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
                return []
            
            logger.info(f"找到 {len(msg_ids)} 封邮件，开始处理...")
            
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
                        subject = subject.replace(' ', '')  # 入口处去空格
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

            # debug输出：显示前20个未匹配邮件的主题
            unmatched = []
            if subject_pattern:
                for msg_id in msg_ids:
                    # 检查是否已在matched_emails
                    if not any(e['id'] == msg_id for e in matched_emails):
                        # 获取原始邮件主题
                        try:
                            response = self.client.fetch([msg_id], ['ENVELOPE'])
                            envelope = response[msg_id][b'ENVELOPE']
                            subject = self.decode_mime_header(envelope.subject.decode()) if envelope.subject else ""
                            subject = subject.replace(' ', '')  # 入口处去空格
                            unmatched.append(subject)
                        except Exception as e:
                            unmatched.append(f"[无法获取主题] id={msg_id}")

                logger.info(f"总邮件数: {len(msg_ids)}")
                logger.info(f"匹配当前规则邮件数: {len(matched_emails)}")
                logger.info(f"未匹配当前规则邮件数: {len(unmatched)}")
                if unmatched:
                    logger.info(f"未匹配邮件数量: {len(unmatched)}，前{min(300, len(unmatched))}个主题如下：")
                    for i, subj in enumerate(unmatched[:300]):
                        print(f"[未匹配{i+1}] {subj}")
            # 保存未匹配列表以便 GUI 弹窗展示
            self.last_unmatched = unmatched
            return matched_emails
            
        except Exception as e:
            logger.error(f"搜索邮件时出错: {str(e)}")
            return []

    def download_email(self, email_id, folder):
        """下载单个邮件，增强验证和错误处理"""
        if not self.client:
            if not self.connect_to_email():
                return None, "连接失败"

        try:
            # 选择文件夹
            try:
                self.client.select_folder(folder, readonly=True)
            except Exception as e:
                logger.warning(f"EXAMINE 失败，回退为 SELECT: {e}")
                try:
                    self.client.select_folder(folder, readonly=False)
                except Exception as e2:
                    logger.error(f"选择邮箱文件夹失败: {e2}")
                    return None, f"选择文件夹失败: {str(e2)}"

            # 获取邮件
            fetch_data = self.client.fetch([email_id], ['RFC822'])
            if not fetch_data or email_id not in fetch_data:
                logger.error(f"获取邮件 {email_id} 失败")
                return None, "获取邮件失败"

            raw_email = fetch_data[email_id][b'RFC822']
            msg = email.message_from_bytes(raw_email)

            # 解析基本信息
            subject = self.decode_mime_header(msg['subject'] or '')
            subject = subject.replace(' ', '')
            sender = email.utils.parseaddr(msg['from'])[1]
            date_str = msg['date'] or ''
            try:
                date_obj = email.utils.parsedate_to_datetime(date_str)
                date_folder = date_obj.strftime("%Y%m%d")
            except:
                date_folder = "unknown_date"

            # 判断收/发类型
            current_email = self.email_addr.lower()
            send_type = "发" if sender.lower() == current_email else "收"
            
            def safe_filename(name):
                return re.sub(r'[\\/:*?"<>|]', '_', name)

            email_dir = self.base_dir / send_type / f"{date_folder}_{safe_filename(subject[:30])}"
            email_dir.mkdir(parents=True, exist_ok=True)
            
            # 解析收件人
            to_addrs = []
            try:
                addrs = email.utils.getaddresses(msg.get_all('To', []) + msg.get_all('Cc', []))
                for name, addr in addrs:
                    if name and name.strip():
                        decoded_name = self.decode_mime_header(name.strip())
                        to_addrs.append(decoded_name)
                    elif addr:
                        local = addr.split('@')[0]
                        to_addrs.append(local)
            except Exception:
                to_addrs = []

# 保持原有的上下文

            # 解析正文部分
            body_text = ""
            has_attachments = False
            attachment_paths = []
            attachment_errors = []
            found_body = False

            for part in msg.walk():
                content_type = part.get_content_type()
                content_disp = str(part.get("Content-Disposition"))

                # 优先保存纯文本正文
                if content_type == "text/plain" and "attachment" not in content_disp:
                    charset = part.get_content_charset()
                    try:
                        if charset:
                            body_text = part.get_payload(decode=True).decode(charset, errors='replace')
                        else:
                            body_text = part.get_payload(decode=True).decode(errors='replace')
                        print("[DEBUG] 原始纯文本正文：", repr(body_text))  # 添加此行
                    except:
                        body_text = "无法解码邮件内容"
                    found_body = True
                    break
                
                # 备选 text/html
                elif content_type == "text/html" and "attachment" not in content_disp and not found_body:
                    charset = part.get_content_charset()
                    try:
                        if charset:
                            html_text = part.get_payload(decode=True).decode(charset, errors='replace')
                        else:
                            html_text = part.get_payload(decode=True).decode(errors='replace')

                        print("[DEBUG] 原始HTML正文：", repr(html_text))  # 添加此行
                        
                        # 日志：输出原始HTML内容
                        logger.debug(f"原始HTML内容: {html_text}")

                        import re as regex
                        # 替换 <br> <br/> <p> </p> 为换行
                        html_text = regex.sub(r'(<br\s*/?>|<p>|</p>|<div[^>]*>|</div>)', '\n', html_text, flags=regex.IGNORECASE)
                        
                        # 日志：输出处理后的HTML内容
                        logger.debug(f"处理后的HTML内容: {html_text}")
                        
                        # 替换&nbsp;为普通空格
                        html_text = html_text.replace('&nbsp;', ' ')
                        # 去除HTML标签
                        body_text = regex.sub('<[^<]+?>', '', html_text)
                        
                        # 额外处理：保留换行和段落格式
                        body_text = regex.sub(r'\n+', '\n', body_text)  # 保证不会多余的换行
                        
                        # 日志：输出最终的body_text
                        logger.debug(f"最终处理后的正文内容: {body_text}")

                        found_body = True
                    except:
                        body_text = "无法解码 HTML 邮件内容"
                        found_body = True

                # 下载附件
                elif "attachment" in content_disp or part.get_filename():
                    has_attachments = True
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

            # 写入 content.txt
            content_file = email_dir / "content.txt"
            try:
                with open(content_file, "w", encoding="utf-8") as f:
                    f.write(f"主题: {subject}\n")
                    f.write(f"发件邮箱: {sender}\n")
                    f.write(f"收件邮箱: {', '.join(to_addrs)}\n")
                    f.write(f"日期: {date_str}\n")
                    f.write("-" * 50 + "\n\n")
                    f.write(body_text if body_text.strip() else "[邮件内容为空或无法解析]")
                logger.info(f"已保存 content.txt: {content_file}")
            except Exception as e:
                error_msg = f"保存 content.txt 失败: {str(e)}"
                logger.error(error_msg)
                return None, error_msg

            # 验证下载完整性
            if not content_file.exists() or content_file.stat().st_size <= 100:
                return None, "content.txt 创建失败或文件过小"
            
            if attachment_errors:
                return None, "; ".join(attachment_errors)

            result = {
                'subject': subject,
                'sender': sender,
                'date': date_str,
                'has_attachments': has_attachments,
                'directory': str(email_dir),
                'attachment_count': len(attachment_paths),
                'attachments': attachment_paths,
                'to_addrs': to_addrs
            }
            logger.info(f"处理完成: {subject}")
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

            # 判断是否已完整下载
            already, email_dir, reason = self._email_already_downloaded(email_info)
            if already:
                logger.info(f"✓ 跳过已下载邮件 id={email_id} 主题={email_info.get('subject')}")
                results.append({
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
                })
                skipped_ids.add(email_id)  # ✅ 修改：单独记录跳过的
            else:
                # 需要下载
                logger.info(f"→ 开始下载 id={email_id} 主题={email_info.get('subject')} (原因: {reason})")
                result, error = self.download_email(email_id, folder)
                if result:
                    result['status'] = 'success'
                    result['reason'] = '下载成功'
                    results.append(result)
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
                if already:
                    logger.info(f"[重试{attempt}] ✓ 已有有效内容 id={email_id}")
                    results.append({
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
                    })
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
                    results.append(res)
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
    
    def generate_download_report(self, results, email_list, failed_details=None):
        """生成报告，TSV 去重，记录失败详情"""
        report_path = self.base_dir / "download_report.txt"
        global_result_path = self.base_dir / "global_result.tsv"
        current_email = self.email_addr.lower()

        # 1. 读取已有 TSV
        existing_oc = set()
        existing_mails = set()
        existing_rows = []
        if global_result_path.exists():
            with open(global_result_path, "r", encoding="utf-8") as f:
                lines = f.readlines()
                if lines:
                    # 解析 OC
                    if lines[0].startswith("oc_registerdata"):
                        oc_line = lines[0].strip().split('\t')
                        if len(oc_line) > 1 and oc_line[1]:
                            existing_oc = set([x.strip() for x in oc_line[1].split(',') if x.strip()])
                    # 解析邮件（关键：用更多字段构建唯一键）
                    for line in lines[2:]:
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
                    m = re.search(r"发件[人x]=【(.+?)】", content)
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
            subject = subject.replace(' ', '')
            letter_type = self.extract_letter_type(subject)
            if not letter_type:
                print(f"[DEBUG] letter_type解析失败，输入subject: '{subject}' 路径: {result.get('directory','')}")

            # 发件时的收件人逻辑
            if send_type == "发":
                if letter_type.startswith("C"):
                    receiver_name = letter_type[1:]
                elif letter_type.startswith("A+C"):
                    receiver_name = letter_type[3:]
                else:
                    receiver_name = ""
            elif send_type == "收" and not receiver_name and only_oc_name:
                receiver_name = only_oc_name

            save_path = result.get('directory', '')
            # 转换为相对路径
            save_path_formatted = save_path
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

            # 去重检查
            key = (sender_name, date_short, letter_type, save_path_formatted)
            if key not in existing_mails:
                row = f"{send_type}\t{date_short}\t{sender_name}\t{receiver_name}\t{letter_type}\t{save_path_formatted}"
                new_rows.append(row)
                existing_mails.add(key)
            else:
                duplicate_count += 1
                logger.info(f"跳过重复项: {sender_name} {date_short} {letter_type}")

        # 4. 写入 TSV
        with open(global_result_path, "w", encoding="utf-8") as f:
            f.write(f"oc_registerdata\t{','.join(all_oc)}\n")
            f.write("收/发类型\t发信日期\t发件人\t收件人\t信件类型\t信件下载位置\n")
            for row in existing_rows:
                f.write(row + "\n")
            for row in new_rows:
                f.write(row + "\n")

        # 5. 写入下载报告
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

    def __init__(self, root):
        self.root = root
        self.root.title("邮件下载工具")
        self.root.geometry("800x800")
        self.root.minsize(800, 800)
        
        # 创建下载器实例
        self.downloader = EmailDownloader()
        
        # 设置图标
        try:
            project_root = Path(os.path.dirname(os.path.abspath(__file__))).resolve()
            icon_dir = project_root / "icon"
            ico_path = icon_dir / "CrackPost.ico"
            png_path = icon_dir / "CrackPost_v1.png"

            # Windows 下优先使用 .ico 作为窗口图标（任务栏/标题栏）
            if ico_path.exists() and sys.platform.startswith("win"):
                try:
                    self.root.iconbitmap(str(ico_path))
                except Exception:
                    pass

            # 若存在 PNG，则用 PIL 生成小图用于 root.iconphoto（跨平台）和大图用于界面内显示
            if png_path.exists():
                pil_img = Image.open(str(png_path)).convert("RGBA")
                # 小图（用于 root.iconphoto，32x32）
                try:
                    small = pil_img.resize((32, 32), Image.LANCZOS)
                    self._app_icon_small = ImageTk.PhotoImage(small)
                    try:
                        self.root.iconphoto(True, self._app_icon_small)
                    except Exception:
                        pass
                except Exception:
                    pass
                # 大图（用于界面内显示，64x64）
                try:
                    large = pil_img.resize((64, 64), Image.LANCZOS)
                    self._app_icon_large = ImageTk.PhotoImage(large)
                except Exception:
                    self._app_icon_large = None
        except Exception:
            pass
            
        # 创建主界面
        self.create_widgets()
        
        # 检查配置并初始化
        self.initialize_app()
    
# --- 注册OC函数修复 ---
    def register_oc(self):
        val = self.oc_input_var.get().strip()
        if val:
            # 只注册不重复的OC
            if val not in self.downloader.oc_registerdata:
                self.downloader.oc_registerdata.append(val)
            # 立即写入 global_result.tsv 的 oc_registerdata 字段
            global_result_path = self.downloader.base_dir / "global_result.tsv"
            if os.path.exists(global_result_path):
                with open(global_result_path, "r", encoding="utf-8") as f:
                    lines = f.readlines()
                if lines and lines[0].startswith("oc_registerdata"):
                    # 更新第一行
                    oc_line = lines[0].strip().split('\t')
                    all_oc = set([x.strip() for x in oc_line[1].split(',') if x.strip()] if len(oc_line) > 1 else [])
                    all_oc.update(self.downloader.oc_registerdata)
                    new_oc_line = f"oc_registerdata\t{','.join(sorted(all_oc))}\n"
                    lines[0] = new_oc_line
                    with open(global_result_path, "w", encoding="utf-8") as f:
                        f.writelines(lines)
            messagebox.showinfo("注册成功", f"已注册: {val}")
            self.oc_input_var.set("")

# --- 智能猜测收件人特殊处理A/B/C信 ---
    def smart_guess_receivers(self):
        """智能猜测收件人并补全global_result.tsv（带详细日志和弹窗提示）"""
        import re

        global_result_path = self.downloader.base_dir / "global_result.tsv"
        if not global_result_path.exists():
            messagebox.showerror("错误", "未找到 global_result.tsv")
            return

        # 读取OC名单和注册日期
        with open(global_result_path, "r", encoding="utf-8") as f:
            lines = f.readlines()
        if not lines or not lines[0].startswith("oc_registerdata"):
            messagebox.showerror("错误", "global_result.tsv 格式错误")
            return

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

        # 判断是否只有一个 OC
        only_one_oc = len(oc_names) == 1
        only_oc_name = oc_names[0] if only_one_oc else None

        changed = 0
        context_logs = []
        new_lines = lines[:2]
        oc_guess_count = {oc: 0 for oc in oc_names}
        unguessed_count = 0

        for idx, row in enumerate(lines[2:]):
            cols = row.strip().split('\t')
            if len(cols) < 6:
                new_lines.append(row)
                continue
            
            send_type = cols[0]
            subject = cols[4]
            send_date = cols[1]
            receiver = cols[3]
            directory = cols[5]
            
            # 只处理收信 且 收件人为空
            if send_type == "收" and not receiver:
                # 1. 只有一个 OC：直接填
                if only_one_oc:
                    cols[3] = only_oc_name
                    changed += 1
                    oc_guess_count[only_oc_name] += 1
                    context_logs.append(
                        f"[单OC自动补全] 行号{idx+3} 原:{row.strip()}\n→ 新:{'\t'.join(cols)}"
                    )
                
                # 2. 多个 OC
                else:
                    # A 信
                    if subject.startswith("A"):
                        oc_to_fill = []
                        for oc in oc_names:
                            reg_date = oc_date_map.get(oc, "")
                            try:
                                if reg_date and send_date and reg_date <= send_date:
                                    oc_to_fill.append(oc)
                            except:
                                pass
                        if oc_to_fill:
                            cols[3] = ",".join(oc_to_fill)
                            changed += 1
                            for oc in oc_to_fill:
                                oc_guess_count[oc] += 1
                            context_logs.append(
                                f"[A信自动补全] 行号{idx+3} 原:{row.strip()}\n→ 新:{'\t'.join(cols)}"
                            )
                        else:
                            unguessed_count += 1
                            context_logs.append(
                                f"[A信无可补全] 行号{idx+3} 原:{row.strip()} 注册日期无符合条件的OC"
                            )
                    
                    # B 信：保持为空
                    elif subject.startswith("B"):
                        unguessed_count += 1
                        context_logs.append(
                            f"[B信无法猜测] 行号{idx+3} 原:{row.strip()}"
                        )
                    
                    # C 信或 A+C 信：从 content.txt 提取
                    elif subject.startswith("C") or subject.startswith("A+C"):
                        # 构建 content.txt 路径
                        project_root = Path(os.path.dirname(os.path.abspath(__file__))).resolve()
                        # directory 格式如 .\email\收\20220910_FW_C_2341_来自尹浦的信
                        if directory.startswith('.\\'):
                            rel_path = directory[2:]
                        else:
                            rel_path = directory
                        content_path = project_root / rel_path.replace('\\', os.sep) / "content.txt"
                        
                        if content_path.exists():
                            try:
                                with open(content_path, "r", encoding="utf-8") as cf:
                                    content_text = cf.read()
                                
                                # 统计每个 OC 名称出现次数
                                oc_count = {oc: content_text.count(oc) for oc in oc_names}
                                
                                # 找出有且仅有 1 次出现的 OC
                                unique_ocs = [oc for oc, count in oc_count.items() if count == 1]
                                
                                if len(unique_ocs) == 1:
                                    cols[3] = unique_ocs[0]
                                    changed += 1
                                    oc_guess_count[unique_ocs[0]] += 1
                                    context_logs.append(
                                        f"[C信自动补全] 行号{idx+3} 原:{row.strip()}\n→ 新:{'\t'.join(cols)} (从content.txt提取: {unique_ocs[0]})"
                                    )
                                else:
                                    unguessed_count += 1
                                    detail = f"各OC出现次数: {oc_count}"
                                    context_logs.append(
                                        f"[C信无法猜测] 行号{idx+3} 原:{row.strip()} content.txt中未找到唯一OC ({detail})"
                                    )
                            except Exception as e:
                                unguessed_count += 1
                                context_logs.append(
                                    f"[C信读取失败] 行号{idx+3} 原:{row.strip()} 错误:{e}"
                                )
                        else:
                            unguessed_count += 1
                            context_logs.append(
                                f"[C信无content.txt] 行号{idx+3} 原:{row.strip()} 路径:{content_path}"
                            )
                    
                    # 其他类型
                    else:
                        unguessed_count += 1
                        context_logs.append(
                            f"[未处理类型] 行号{idx+3} 原:{row.strip()} subject:{subject}"
                        )
            
            new_lines.append('\t'.join(cols) + '\n')

        # 写入文件
        with open(global_result_path, "w", encoding="utf-8") as f:
            for line in new_lines:
                f.write(line if line.endswith('\n') else line + '\n')

        log_text = f"已补全 {changed} 封邮件的收件人\n\n" + "\n".join(context_logs)
        self.add_download_log(log_text)

        # 前台弹窗提示
        detail = "\n".join([f"{oc}: {oc_guess_count[oc]} 封" for oc in oc_names])
        messagebox.showinfo(
            "智能猜测完成",
            f"自动补全结果：\n{detail}\n还有 {unguessed_count} 封无法被猜测。\n\n详细见下载日志区"
        )
    def toggle_search_mode(self):
        """切换检索模式（手动/配置化）"""
        mode = self.search_mode_var.get()
        if mode == "config":
            try:
                rule_num = int(self.rule_choice_var.get())
                rule = self.downloader.get_default_rule(rule_num)
                self.subject_var.set(rule.get('subject_pattern', ''))
                self.sender_var.set(rule.get('from', ''))
                self.recipient_var.set(rule.get('to', ''))
                # 只禁用文本输入框，不禁用DateEntry
                for widget in self.search_frame.winfo_children():
                    if isinstance(widget, ttk.Entry):
                        widget.configure(state="readonly")
            except Exception as e:
                messagebox.showerror("错误", f"加载配置规则失败: {str(e)}")
                self.search_mode_var.set("manual")
        else:
            for widget in self.search_frame.winfo_children():
                if isinstance(widget, ttk.Entry):
                    widget.configure(state="normal")

    def create_widgets(self):
        """创建GUI组件"""
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 使用Notebook创建选项卡
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True)
        # 保存为实例属性，便于其他方法切换选项卡
        self.notebook = notebook

        # 创建三个选项卡
        self.login_frame = ttk.Frame(notebook, padding="10")
        self.search_frame = ttk.Frame(notebook, padding="10")
        self.download_frame = ttk.Frame(notebook, padding="10")

        notebook.add(self.login_frame, text="登录设置")
        notebook.add(self.search_frame, text="搜索邮件")
        notebook.add(self.download_frame, text="下载结果")

        # 登录设置选项卡
        self.create_login_tab()
        # 搜索邮件选项卡
        self.create_search_tab()
        # 下载结果选项卡
        self.create_download_tab()

        # 底部状态栏
        status_frame = ttk.Frame(self.root, relief=tk.SUNKEN, padding=(2, 2))
        status_frame.pack(fill=tk.X, side=tk.BOTTOM)

        self.status_var = tk.StringVar()
        self.status_var.set("就绪")
        status_label = ttk.Label(status_frame, textvariable=self.status_var, anchor=tk.W)
        status_label.pack(side=tk.LEFT, fill=tk.X)

        # 版本信息
        version_label = ttk.Label(status_frame, text="v1.0", anchor=tk.E)
        version_label.pack(side=tk.RIGHT)
        
    def open_visualization_html(self):
        """后台启动本地服务器并用浏览器打开可视化页面"""
        import time

        # 1. 启动 http.server（如果已启动则忽略报错）
        try:
            # Windows下隐藏cmd窗口
            creationflags = 0
            startupinfo = None
            if sys.platform.startswith('win'):
                creationflags = subprocess.CREATE_NEW_CONSOLE
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW

            # 检查端口是否已被占用（简单方式）
            import socket
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            try:
                sock.connect(('localhost', 8000))
                sock.close()
                server_running = True
            except Exception:
                server_running = False

            if not server_running:
                subprocess.Popen(
                    [sys.executable, "-m", "http.server", "8000"],
                    cwd=os.path.dirname(os.path.abspath(__file__)),
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    creationflags=creationflags,
                    startupinfo=startupinfo
                )
                time.sleep(1)  # 等待服务器启动

        except Exception as e:
            messagebox.showerror("启动服务器失败", f"无法启动本地服务器: {e}")
            return

        # 2. 打开浏览器
        url = "http://localhost:8000/visualization_private.html"
        webbrowser.open_new_tab(url)

    def create_login_tab(self):
        """创建登录设置选项卡"""
        # 邮箱设置框架
        settings_frame = ttk.LabelFrame(self.login_frame, text="邮箱设置", padding=(10, 5))
        settings_frame.pack(fill=tk.X, pady=10)
        
        # 邮箱地址
        ttk.Label(settings_frame, text="邮箱地址:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.email_var = tk.StringVar(value=self.downloader.email_addr)
        email_entry = ttk.Entry(settings_frame, textvariable=self.email_var, width=40)
        email_entry.grid(row=0, column=1, sticky=tk.W, pady=5)
        
        # 密码/授权码
        ttk.Label(settings_frame, text="授权码:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.password_var = tk.StringVar(value=self.downloader.password)
        password_entry = ttk.Entry(settings_frame, textvariable=self.password_var, show="*", width=40)
        password_entry.grid(row=1, column=1, sticky=tk.W, pady=5)
        
        # IMAP服务器
        ttk.Label(settings_frame, text="IMAP服务器:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.server_var = tk.StringVar(value=self.downloader.imap_server)
        server_entry = ttk.Entry(settings_frame, textvariable=self.server_var, width=40)
        server_entry.grid(row=2, column=1, sticky=tk.W, pady=5)
        
        # 常用服务器下拉菜单
        ttk.Label(settings_frame, text="常用服务器:").grid(row=3, column=0, sticky=tk.W, pady=5)
        servers = {
            "QQ邮箱": "imap.qq.com",
            "Gmail": "imap.gmail.com",
            "Outlook": "outlook.office365.com",
            "163邮箱": "imap.163.com",
            "126邮箱": "imap.126.com"
        }
        server_names = list(servers.keys())
        self.server_combo = ttk.Combobox(settings_frame, values=server_names, width=38)
        self.server_combo.grid(row=3, column=1, sticky=tk.W, pady=5)
        self.server_combo.bind("<<ComboboxSelected>>", lambda e: self.server_var.set(servers[self.server_combo.get()]))
        
        # 高级设置框架
        advanced_frame = ttk.LabelFrame(self.login_frame, text="高级设置", padding=(10, 5))
        advanced_frame.pack(fill=tk.X, pady=10)
        
        # 超时设置
        ttk.Label(advanced_frame, text="连接超时(秒):").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.timeout_var = tk.StringVar(value=str(self.downloader.timeout))
        timeout_entry = ttk.Entry(advanced_frame, textvariable=self.timeout_var, width=10)
        timeout_entry.grid(row=0, column=1, sticky=tk.W, pady=5)
        
        # 批量大小
        ttk.Label(advanced_frame, text="批处理大小:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.batch_var = tk.StringVar(value=str(self.downloader.batch_size))
        batch_entry = ttk.Entry(advanced_frame, textvariable=self.batch_var, width=10)
        batch_entry.grid(row=1, column=1, sticky=tk.W, pady=5)
        
        # 按钮框架
        button_frame = ttk.Frame(self.login_frame)
        button_frame.pack(fill=tk.X, pady=20)
        
        # 测试连接按钮
        test_conn_btn = ttk.Button(button_frame, text="测试连接并登录", command=self.test_connection)
        test_conn_btn.pack(side=tk.LEFT, padx=10)
        
        # 保存设置按钮
        save_settings_btn = ttk.Button(button_frame, text="保存设置", command=self.save_settings)
        save_settings_btn.pack(side=tk.LEFT, padx=10)
        
        # 帮助信息
        help_text = """
        步骤说明：
        1. 获取授权码：在邮箱设置里获取 IMAP/授权码（例如 QQ 邮箱需开启 IMAP 并生成授权码），如有日期限制，请选择全部。
        2. 登录：在“登录设置”中填写邮箱地址、授权码与 IMAP 服务器，点击“测试连接并登录”确认可用。
        3. 选择检索规则：在“搜索邮件”页选择或填写主题/发件人/日期等规则，留空表示全选，然后点击“搜索邮件”。
        4. 注册 OC：在“下载结果”页注册你的 OC（格式：名称_YYYY-MM-DD），用于后续自动匹配与标注。
        5. 智能补齐：若同一 OC 有多个马甲，下载完成后可点击“智能猜测收件人”自动补齐 global_result.tsv 中的收件人字段。
        6. 手工修正：在生成的 global_result.tsv 中对不满意的条目（收件人、路径等）进行手工修正。
        7. 启动可视化：在“下载结果”页点击“打开邮件网络可视化”查看关系图。

        ！注意，当信件格式出现纰漏时，可能导致部分信息无法正确解析。当出现这种情况时：
        1. 在email文件夹，仿照其他其他已经被正确填装的信件，手工修正content.txt中的发件人等信息。
        2. 然后重新运行“智能猜测收件人”功能，补齐global_result.tsv中的收件人字段。
        3. 如果出现收发信人被识别为自己的邮箱而非oc名时，请手动修正为oc名。

        玩的开心！
        """
        
        help_frame = ttk.LabelFrame(self.login_frame, text="帮助信息")
        help_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        help_text_widget = scrolledtext.ScrolledText(help_frame, wrap=tk.WORD, height=8)
        help_text_widget.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        help_text_widget.insert(tk.END, help_text)
        help_text_widget.configure(state="disabled")
    
    def create_search_tab(self):
        """创建搜索选项卡"""
        # 文件夹选择框架
        folder_frame = ttk.LabelFrame(self.search_frame, text="选择邮箱文件夹", padding=(10, 5))
        folder_frame.pack(fill=tk.X, pady=5)
        
        # 文件夹列表
        ttk.Label(folder_frame, text="文件夹:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.folder_var = tk.StringVar()
        self.folder_combo = ttk.Combobox(folder_frame, textvariable=self.folder_var, width=40)
        self.folder_combo.grid(row=0, column=1, sticky=tk.W, pady=5)
        
        # 刷新按钮
        refresh_btn = ttk.Button(folder_frame, text="刷新列表", command=self.refresh_folders)
        refresh_btn.grid(row=0, column=2, padx=5, pady=5)
        
        # 检索方式选择
        mode_frame = ttk.LabelFrame(self.search_frame, text="检索方式", padding=(10, 5))
        mode_frame.pack(fill=tk.X, pady=5)
        # ...在mode_frame之后添加...
        self.rule_choice_var = tk.StringVar(value="1")
        rule_frame = ttk.Frame(self.search_frame)
        rule_frame.pack(fill=tk.X, pady=2)
        ttk.Label(rule_frame, text="自动化配置规则[仅在选择配置化检索时可用,1下载收件,2下载发件]:").pack(side=tk.LEFT)
        self.rule_combo = ttk.Combobox(rule_frame, textvariable=self.rule_choice_var, width=10, state="readonly")
        self.rule_combo['values'] = ["1", "2"]
        self.rule_combo.pack(side=tk.LEFT, padx=5)
        self.rule_combo.bind("<<ComboboxSelected>>", lambda e: self.toggle_search_mode())

        self.search_mode_var = tk.StringVar(value="manual")
        ttk.Radiobutton(mode_frame, text="手动检索[自行输入检索规则]", variable=self.search_mode_var, 
                    value="manual", command=self.toggle_search_mode).pack(side=tk.LEFT, padx=10)
        ttk.Radiobutton(mode_frame, text="配置化检索[采用已配置好的检索规则]", variable=self.search_mode_var, 
                    value="config", command=self.toggle_search_mode).pack(side=tk.LEFT, padx=10)
        
        # 搜索条件框架
        search_frame = ttk.LabelFrame(self.search_frame, text="搜索条件[留空则为全选]", padding=(10, 5))
        search_frame.pack(fill=tk.X, pady=10)
        
        # 日期范围
        ttk.Label(search_frame, text="起始日期:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.start_date_var = tk.StringVar()
        self.start_date_picker = DateEntry(search_frame, width=12, textvariable=self.start_date_var, 
                                          date_pattern='yyyy-mm-dd')
        self.start_date_picker.grid(row=0, column=1, sticky=tk.W, pady=5)
        self.start_date_picker.delete(0, tk.END)  # 清空默认值
        
        ttk.Label(search_frame, text="结束日期:").grid(row=0, column=2, sticky=tk.W, pady=5)
        self.end_date_var = tk.StringVar()
        self.end_date_picker = DateEntry(search_frame, width=12, textvariable=self.end_date_var, 
                                        date_pattern='yyyy-mm-dd')
        self.end_date_picker.grid(row=0, column=3, sticky=tk.W, pady=5)
        
        # 发件人
        ttk.Label(search_frame, text="发件邮箱:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.sender_var = tk.StringVar()
        sender_entry = ttk.Entry(search_frame, textvariable=self.sender_var, width=40)
        sender_entry.grid(row=1, column=1, columnspan=3, sticky=tk.W, pady=5)
        
        # 收件人
        ttk.Label(search_frame, text="收件邮箱:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.recipient_var = tk.StringVar()
        recipient_entry = ttk.Entry(search_frame, textvariable=self.recipient_var, width=40)
        recipient_entry.grid(row=2, column=1, columnspan=3, sticky=tk.W, pady=5)
        
        # 主题
        ttk.Label(search_frame, text="主题包含:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.subject_var = tk.StringVar()
        subject_entry = ttk.Entry(search_frame, textvariable=self.subject_var, width=40)
        subject_entry.grid(row=3, column=1, columnspan=3, sticky=tk.W, pady=5)
        
        # 最大结果数
        ttk.Label(search_frame, text="最多搜索:").grid(row=4, column=0, sticky=tk.W, pady=5)
        self.max_results_var = tk.StringVar(value="1500")
        max_results_entry = ttk.Entry(search_frame, textvariable=self.max_results_var, width=10)
        max_results_entry.grid(row=4, column=1, sticky=tk.W, pady=5)
        ttk.Label(search_frame, text="封邮件").grid(row=4, column=2, sticky=tk.W, pady=5)
        
        # 搜索按钮
        button_frame = ttk.Frame(self.search_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        search_btn = ttk.Button(button_frame, text="搜索邮件", command=self.search_emails)
        search_btn.pack(side=tk.LEFT, padx=10)
        
        clear_btn = ttk.Button(button_frame, text="清除条件", command=self.clear_search)
        clear_btn.pack(side=tk.LEFT, padx=10)
        
        # 搜索结果框架（垂直高度缩小：父框不再扩展占满纵向空间，Treeview 限制可见行数）
        results_frame = ttk.LabelFrame(self.search_frame, text="搜索结果")
        # 通过 fill=tk.X + expand=False 限制父框纵向占用，调整 pady 让上下间距更紧凑
        results_frame.pack(fill=tk.X, expand=False, pady=6)
        
        # 创建表格（height 控制可见行数，越小越短）
        columns = ("序号", "日期", "发件人", "主题")
        self.results_tree = ttk.Treeview(results_frame, columns=columns, show="headings", selectmode="extended", height=8)
        self.results_tree.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # 添加滚动条
        scrollbar = ttk.Scrollbar(results_frame, orient=tk.VERTICAL, command=self.results_tree.yview)
        self.results_tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        
        # 设置列宽和标题
        self.results_tree.heading("序号", text="序号")
        self.results_tree.heading("日期", text="日期")
        self.results_tree.heading("发件人", text="发件人")
        self.results_tree.heading("主题", text="主题")
        
        self.results_tree.column("序号", width=50, anchor="center")
        self.results_tree.column("日期", width=150)
        self.results_tree.column("发件人", width=200)
        self.results_tree.column("主题", width=400)
        
        
        # 状态和进度条框架
        status_frame = ttk.Frame(self.search_frame)
        status_frame.pack(fill=tk.X, pady=5)
        
        self.search_status_var = tk.StringVar()
        search_status = ttk.Label(status_frame, textvariable=self.search_status_var)
        search_status.pack(side=tk.LEFT, padx=5)
        
        self.search_progress = ttk.Progressbar(status_frame, orient=tk.HORIZONTAL, length=200, mode='determinate')
        self.search_progress.pack(side=tk.RIGHT, padx=5)
        
        # 下载按钮框架
        download_frame = ttk.Frame(self.search_frame)
        download_frame.pack(fill=tk.X, pady=10)
        
        self.select_all_var = tk.BooleanVar()
        select_all_check = ttk.Checkbutton(download_frame, text="全选", 
                                         variable=self.select_all_var, 
                                         command=self.toggle_select_all)
        select_all_check.pack(side=tk.LEFT, padx=10)
        
        download_selected_btn = ttk.Button(download_frame, text="下载选中邮件", 
                                        command=self.download_selected)
        download_selected_btn.pack(side=tk.RIGHT, padx=10)

    def register_oc(self):
        val = self.oc_input_var.get().strip()
        if val:
            # 只注册不重复的OC
            if val not in self.downloader.oc_registerdata:
                self.downloader.oc_registerdata.append(val)
            # 立即写入 global_result.tsv 的 oc_registerdata 字段
            global_result_path = self.downloader.base_dir / "global_result.tsv"
            if os.path.exists(global_result_path):
                with open(global_result_path, "r", encoding="utf-8") as f:
                    lines = f.readlines()
                if lines and lines[0].startswith("oc_registerdata"):
                    # 更新第一行
                    oc_line = lines[0].strip().split('\t')
                    all_oc = set([x.strip() for x in oc_line[1].split(',') if x.strip()] if len(oc_line) > 1 else [])
                    all_oc.update(self.downloader.oc_registerdata)
                    new_oc_line = f"oc_registerdata\t{','.join(sorted(all_oc))}\n"
                    lines[0] = new_oc_line
                    with open(global_result_path, "w", encoding="utf-8") as f:
                        f.writelines(lines)
            messagebox.showinfo("注册成功", f"已注册: {val}")
            self.oc_input_var.set("")

    def create_download_tab(self):
        """创建下载选项卡"""
        reg_frame = ttk.LabelFrame(self.download_frame, text="注册OC[eg:CrackPostUser_2022-07-01]", padding=(10, 5))
        reg_frame.pack(fill=tk.X, pady=5)
        self.oc_input_var = tk.StringVar()
        ttk.Entry(reg_frame, textvariable=self.oc_input_var, width=30).pack(side=tk.LEFT, padx=5)
        ttk.Button(reg_frame, text="注册", command=self.register_oc).pack(side=tk.LEFT, padx=5)

        # 下载进度框架
        progress_frame = ttk.LabelFrame(self.download_frame, text="下载进度", padding=(10, 5))
        progress_frame.pack(fill=tk.X, pady=5)
        
        self.download_status_var = tk.StringVar(value="等待下载...")
        download_status = ttk.Label(progress_frame, textvariable=self.download_status_var)
        download_status.pack(fill=tk.X, pady=5)
        
        self.download_progress = ttk.Progressbar(progress_frame, orient=tk.HORIZONTAL, length=100, 
                                            mode='determinate')
        self.download_progress.pack(fill=tk.X, pady=5)
        
        # 下载结果框架（垂直高度缩小）
        results_frame = ttk.LabelFrame(self.download_frame, text="下载结果")
        results_frame.pack(fill=tk.X, expand=False, pady=6)
        
       # 创建文本区域，height 指定可视行数，值越小区域越短
        self.download_log = scrolledtext.ScrolledText(results_frame, wrap=tk.WORD, height=10)
        self.download_log.pack(fill=tk.X, expand=True, padx=5, pady=5)

        # 提示块
        tip_frame = ttk.LabelFrame(self.download_frame, text="可视化提示", padding=(10, 5))
        tip_frame.pack(fill=tk.X, pady=10)

        # 智能猜测按钮
        smart_guess_btn = ttk.Button(
            tip_frame,
            text="智能猜测收件人",
            command=self.smart_guess_receivers
        )
        smart_guess_btn.pack(anchor=tk.W, padx=5, pady=5)

        tip_label = ttk.Label(
            tip_frame,
            text="点击打开可视化人际关系网络",
            # foreground="#0099cc"
        )
        tip_label.pack(anchor=tk.W, padx=5, pady=5)

        # 启动可视化按钮
        open_vis_btn = ttk.Button(
            tip_frame,
            text="打开邮件网络可视化",
            command=self.open_visualization_html
        )
        open_vis_btn.pack(anchor=tk.W, padx=5, pady=5)


        
        
        # 按钮框架
        button_frame = ttk.Frame(self.download_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        open_folder_btn = ttk.Button(button_frame, text="打开存储文件夹", 
                                   command=self.open_storage_folder)
        open_folder_btn.pack(side=tk.LEFT, padx=10)
        
        view_report_btn = ttk.Button(button_frame, text="查看下载报告", 
                                   command=self.view_download_report)
        view_report_btn.pack(side=tk.LEFT, padx=10)
        
        clear_log_btn = ttk.Button(button_frame, text="清除日志", 
                                 command=self.clear_download_log)
        clear_log_btn.pack(side=tk.RIGHT, padx=10)
    
    def initialize_app(self):
        """初始化应用程序"""
        # 检查配置是否有效
        if not self.downloader.email_addr or not self.downloader.password or not self.downloader.imap_server:
            messagebox.showinfo("首次使用", "请在'登录设置'选项卡中配置您的邮箱信息")
    
    def browse_directory(self):
        """浏览并选择存储目录"""
        from tkinter import filedialog
        directory = filedialog.askdirectory(initialdir=self.storage_var.get())
        if directory:
            self.storage_var.set(directory)
    
    def save_settings(self):
        """保存设置"""
        try:
            self.downloader.email_addr = self.email_var.get()
            self.downloader.password = self.password_var.get()
            self.downloader.imap_server = self.server_var.get()
            self.downloader.timeout = int(self.timeout_var.get())
            self.downloader.batch_size = int(self.batch_var.get())
            
            self.downloader.save_config()
            messagebox.showinfo("成功", "设置已保存")
        except Exception as e:
            messagebox.showerror("错误", f"保存设置失败: {str(e)}")
    
    def test_connection(self):
        """测试邮箱连接"""
        # 更新配置
        self.downloader.email_addr = self.email_var.get()
        self.downloader.password = self.password_var.get()
        self.downloader.imap_server = self.server_var.get()
        
        # 显示连接中...
        self.status_var.set("连接中...")
        self.root.update()
        
        # 尝试连接
        success = self.downloader.connect_to_email()
        
        if success:
            messagebox.showinfo("成功", "连接成功，邮箱登录验证通过")
            self.status_var.set("就绪")
            # 自动切换到“搜索邮件”选项卡，便于立即进行检索
            try:
                if hasattr(self, 'notebook'):
                    # 以页面对象切换
                    self.notebook.select(self.search_frame)
            except Exception:
                pass
    
            # 自动获取文件夹列表
            self.refresh_folders()
        else:
            messagebox.showerror("错误", "连接失败，请检查邮箱设置")
            self.status_var.set("连接失败")
    
    def refresh_folders(self):
        """刷新邮箱文件夹列表"""
        if not self.downloader.client:
            if not self.downloader.connect_to_email():
                messagebox.showerror("错误", "请先登录邮箱")
                return
        
        # 显示正在获取...
        self.status_var.set("正在获取文件夹...")
        self.root.update()
        
        # 获取文件夹列表
        folders = self.downloader.get_all_mailboxes()
        
        if folders:
            self.folder_combo['values'] = folders
            # 优先选 INBOX（126/163 常用），否则选第一个
            idx = folders.index('INBOX') if 'INBOX' in folders else 0
            self.folder_combo.current(idx)
            self.status_var.set(f"已获取 {len(folders)} 个文件夹")
        else:
            self.status_var.set("获取文件夹失败")
    
    def clear_search(self):
        """清除搜索条件"""
        self.start_date_picker.delete(0, tk.END)
        self.end_date_picker.delete(0, tk.END)
        self.sender_var.set("")
        self.recipient_var.set("")
        self.subject_var.set("")
        self.max_results_var.set("1500")
    
    def search_emails(self):
        """搜索邮件"""
        if not self.downloader.client:
            if not self.downloader.connect_to_email():
                messagebox.showerror("错误", "请先登录邮箱")
                return
        
        # 获取搜索条件
        folder = self.folder_var.get()
        if not folder:
            messagebox.showerror("错误", "请选择邮箱文件夹")
            return
        if self.search_mode_var.get() == "config":
            rule_num = int(self.rule_choice_var.get())
            rule = self.downloader.get_default_rule(rule_num)
            subject_pattern = rule.get('subject_pattern', '')
            sender = rule.get('from', '')
            recipient = rule.get('to', '')
        else:
            # 手动检索
            subject_pattern = self.subject_var.get()
            sender = self.sender_var.get()
            recipient = self.recipient_var.get()
        
        # 处理日期
        start_date = None
        end_date = None
        
        if self.start_date_var.get():
            try:
                start_date = datetime.strptime(self.start_date_var.get(), '%Y-%m-%d')
            except:
                messagebox.showerror("错误", "起始日期格式错误")
                return
        
        if self.end_date_var.get():
            try:
                end_date = datetime.strptime(self.end_date_var.get(), '%Y-%m-%d')
            except:
                messagebox.showerror("错误", "结束日期格式错误")
                return
        
        # 获取最大结果数
        try:
            max_emails = int(self.max_results_var.get())
            if max_emails <= 0:
                max_emails = 1500
        except:
            max_emails = 1500
        
        # 清空当前结果
        for item in self.results_tree.get_children():
            self.results_tree.delete(item)
        
        # 重置进度条
        self.search_progress['value'] = 0
        self.search_status_var.set("搜索中...")
        self.root.update()
        
        # 创建并启动搜索线程
        search_thread = threading.Thread(
            target=self.run_search,
            args=(folder, subject_pattern, sender, recipient, start_date, end_date, max_emails)
        )
        search_thread.daemon = True
        search_thread.start()
    
    def run_search(self, folder, subject_pattern, sender, recipient, start_date, end_date, max_emails):
        """在独立线程中运行搜索"""
        try:
            # 更新状态
            def update_progress(progress, found_count):
                self.search_progress['value'] = progress
                self.search_status_var.set(f"搜索中... {progress}% 已找到 {found_count} 封")
                self.root.update()
            
            # 执行搜索
            results = self.downloader.search_emails_advanced(
                folder=folder,
                subject_pattern=subject_pattern,
                from_address=sender,
                to_address=recipient,
                start_date=start_date,
                end_date=end_date,
                callback=update_progress,
                max_emails=max_emails
            )
            
            # 存储结果用于后续下载
            self.downloader.search_results = results

            # 若有未匹配项，主线程弹窗显示（可滚动）
            unmatched = getattr(self.downloader, 'last_unmatched', []) or []
            if unmatched:
                def _show_unmatched():
                    win = tk.Toplevel(self.root)
                    win.title("未匹配邮件列表")
                    win.geometry("700x420")
                    txt = scrolledtext.ScrolledText(win, wrap=tk.WORD)
                    txt.pack(fill=tk.BOTH, expand=True, padx=4, pady=3)
                    header = f"未匹配邮件数量: {len(unmatched)}\n显示前 {min(300, len(unmatched))} 条：\n\n"
                    txt.insert(tk.END, header + "\n".join(unmatched[:300]))
                    txt.configure(state='disabled')
                    # 模态
                    try:
                        win.transient(self.root)
                        win.grab_set()
                    except Exception:
                        pass
                self.root.after(0, _show_unmatched)

            # 在主线程中更新界面
            self.root.after(0, lambda: self.update_search_results(results))
            
        except Exception as e:
            # 在主线程中显示错误
            self.root.after(0, lambda: messagebox.showerror("错误", f"搜索时出错: {str(e)}"))
            logger.error(f"搜索时出错: {str(e)}")
            self.root.after(0, lambda: self.search_status_var.set("搜索失败"))
    
    def update_search_results(self, results):
        """更新搜索结果到表格"""
        # 清空当前结果
        for item in self.results_tree.get_children():
            self.results_tree.delete(item)
        
        # 添加结果到表格
        for i, email_info in enumerate(results):
            # 处理日期显示
            try:
                if email_info.get('date_obj'):
                    date_str = email_info['date_obj'].strftime("%Y-%m-%d %H:%M")
                else:
                    date_str = email_info['date'][:19] if email_info['date'] else "未知日期"
            except:
                date_str = str(email_info.get('date', "未知日期"))
            
            # 限制字段长度
            sender = email_info['sender']
            if len(sender) > 30:
                sender = sender[:27] + "..."
                
            subject = email_info['subject']
            if len(subject) > 50:
                subject = subject[:47] + "..."
            
            self.results_tree.insert("", tk.END, values=(
                i + 1,
                date_str,
                sender,
                subject
            ), tags=(str(i),))
        
        # 更新状态
        self.search_status_var.set(f"找到 {len(results)} 封符合条件的邮件")
        self.search_progress['value'] = 100
    
    def toggle_select_all(self):
        """切换全选/取消全选"""
        if self.select_all_var.get():
            # 选择所有项
            for item in self.results_tree.get_children():
                self.results_tree.selection_add(item)
        else:
            # 取消选择所有项
            for item in self.results_tree.get_children():
                self.results_tree.selection_remove(item)
    
    def download_selected(self):
        """下载选中的邮件"""
        # 获取选中的项
        selected_items = self.results_tree.selection()
        if not selected_items:
            messagebox.showinfo("提示", "请先选择需要下载的邮件")
            return
        
        # 获取选中的索引
        selected_indices = []
        for item in selected_items:
            item_values = self.results_tree.item(item, "values")
            if item_values:
                selected_indices.append(int(item_values[0]) - 1)
        
        # 检查是否有结果
        if not self.downloader.search_results:
            messagebox.showerror("错误", "没有可下载的邮件")
            return
        
        # 准备要下载的邮件
        emails_to_download = [self.downloader.search_results[i] for i in selected_indices 
                             if 0 <= i < len(self.downloader.search_results)]
        
        if not emails_to_download:
            messagebox.showerror("错误", "选择的邮件不可用")
            return
        
        # 切换到下载选项卡
        self.root.nametowidget(".!frame.!notebook").select(2)  # 索引从0开始，下载选项卡是第3个
        
        # 重置下载进度
        self.download_progress['value'] = 0
        self.download_status_var.set(f"准备下载 {len(emails_to_download)} 封邮件...")
        self.clear_download_log()
        self.root.update()
        
        # 创建并启动下载线程
        download_thread = threading.Thread(
            target=self.run_download,
            args=(emails_to_download,)
        )
        download_thread.daemon = True
        download_thread.start()

    def run_download(self, emails):
        """在独立线程中运行下载"""
        try:
            def update_progress(progress, current, total):
                self.download_progress['value'] = progress
                self.download_status_var.set(f"下载中... {progress}% ({current}/{total})")
                self.root.update()
                
                # ✅ 修改：从实际处理过的邮件中获取状态
                if current <= len(emails):
                    email_info = emails[current - 1]
                    status = email_info.get('status', '')
                    reason = email_info.get('reason', '')
                    
                    if status == 'skipped':
                        msg = f"[{current}/{total}] ⊙ 跳过: {email_info.get('subject', '')} ({reason})\n"
                    elif status == 'failed':
                        msg = f"[{current}/{total}] ✗ 失败: {email_info.get('subject', '')} ({reason})\n"
                    elif status == 'success':
                        msg = f"[{current}/{total}] ✓ 已下载: {email_info.get('subject', '')}\n"
                    else:
                        msg = f"[{current}/{total}] → 处理中: {email_info.get('subject', '')}\n"
                    
                    self.root.after(0, lambda m=msg: self.add_download_log(m))
            
            # 执行下载
            results = self.downloader.download_multiple_emails(emails, update_progress)
            
            self.root.after(0, lambda: self.download_completed(results, len(emails)))
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("错误", f"下载时出错: {str(e)}"))
            logger.error(f"下载时出错: {str(e)}")
            self.root.after(0, lambda: self.download_status_var.set("下载失败"))

    def download_completed(self, results, total):
        """下载完成后的处理"""
        self.download_progress['value'] = 100
        
        success_count = len([r for r in results if r.get('status') == 'success'])
        skipped_count = len([r for r in results if r.get('status') == 'skipped'])
        failed_count = len([r for r in results if r.get('status') == 'failed'])
        
        self.download_status_var.set(f"下载完成: 成功{success_count} 跳过{skipped_count} 失败{failed_count}")
        
        # 添加摘要到日志
        self.add_download_log("\n" + "="*50 + "\n")
        self.add_download_log(f"下载完成摘要:\n")
        self.add_download_log(f"总邮件数: {total}\n")
        self.add_download_log(f"✓ 成功下载: {success_count}\n")
        self.add_download_log(f"⊙ 跳过已有: {skipped_count}\n")
        self.add_download_log(f"✗ 下载失败: {failed_count}\n")
        self.add_download_log(f"存储目录: {self.downloader.base_dir}\n")
        
        # 失败详情
        if failed_count > 0:
            self.add_download_log("\n失败邮件详情:\n")
            for r in results:
                if r.get('status') == 'failed':
                    self.add_download_log(f"  - {r['subject']}: {r.get('reason', '未知原因')}\n")
        
        self.add_download_log("\n下载报告已保存到存储目录。\n")
        
        # 显示通知
        messagebox.showinfo("完成", f"下载完成\n成功: {success_count}\n跳过: {skipped_count}\n失败: {failed_count}")    
     
    def add_download_log(self, text):
        """添加文本到下载日志"""
        self.download_log.configure(state="normal")
        self.download_log.insert(tk.END, text)
        self.download_log.see(tk.END)  # 滚动到底部
        self.download_log.configure(state="disabled")
    
    def clear_download_log(self):
        """清空下载日志"""
        self.download_log.configure(state="normal")
        self.download_log.delete(1.0, tk.END)
        self.download_log.configure(state="disabled")
    
    def open_storage_folder(self):
        """打开存储文件夹"""
        try:
            import os
            import platform
            
            path = str(self.downloader.base_dir)
            
            if platform.system() == "Windows":
                os.startfile(path)
            elif platform.system() == "Darwin":  # macOS
                import subprocess
                subprocess.Popen(["open", path])
            else:  # Linux
                import subprocess
                subprocess.Popen(["xdg-open", path])
                
        except Exception as e:
            messagebox.showerror("错误", f"无法打开文件夹: {str(e)}")
    
    def view_download_report(self):
        """查看下载报告"""
        report_path = self.downloader.base_dir / "download_report.txt"
        
        if not os.path.exists(report_path):
            messagebox.showinfo("提示", "尚未生成下载报告")
            return
        
        try:
            # 显示报告内容
            with open(report_path, "r", encoding="utf-8") as f:
                report_text = f.read()
            
            # 创建对话框
            report_dialog = tk.Toplevel(self.root)
            report_dialog.title("下载报告")
            report_dialog.geometry("700x500")
            
            # 添加文本框
            report_text_widget = scrolledtext.ScrolledText(report_dialog, wrap=tk.WORD)
            report_text_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            report_text_widget.insert(tk.END, report_text)
            report_text_widget.configure(state="disabled")
            
            # 添加关闭按钮
            close_btn = ttk.Button(report_dialog, text="关闭", command=report_dialog.destroy)
            close_btn.pack(pady=10)
            
            # 设置模态
            report_dialog.transient(self.root)
            report_dialog.grab_set()
            self.root.wait_window(report_dialog)
        except Exception as e:
            messagebox.showerror("错误", f"无法打开下载报告: {str(e)}")
        

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
    
    # 应用程序图标
    try:
        project_root = Path(os.path.dirname(os.path.abspath(__file__))).resolve()
        icon_png = project_root / "icon" / "CrackPost_v1.png"
        if icon_png.exists():
            app_icon = tk.PhotoImage(file=str(icon_png))
            root.iconphoto(True, app_icon)
            # 保留引用
            root._app_icon = app_icon
    except Exception:
        pass
    
    # 运行应用程序
    root.mainloop()

if __name__ == "__main__":
    main()