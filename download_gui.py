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
import shutil
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
        
        # 添加超时设置（弱网/大附件默认 300s；过短会导致 fetch 失败或整批卡死）
        self.timeout = int(self.config.get('Connection', 'timeout', fallback='300'))
        self.batch_size = int(self.config.get('Connection', 'batch_size', fallback='100'))
        self.cancel_requested = False
        self.crackpost_addrs = {
            'crackpost2@126.com',
            'crackpost@126.com',
        }

    def request_cancel(self):
        """协作式取消当前大批量下载（以已合法信件为准落盘）。
        立刻丢弃 IMAP，避免卡在当前这封的 fetch/看门狗里「点了结束却不停」。
        """
        self.cancel_requested = True
        try:
            self._abandon_imap_client()
        except Exception:
            pass

    def reset_cancel(self):
        self.cancel_requested = False

    def effective_imap_timeout(self):
        return max(30, int(getattr(self, "timeout", 300) or 300))

    def effective_email_watchdog(self):
        """单封看门狗须 ≥ IMAP timeout，避免掐死本可成功的慢信。"""
        t = self.effective_imap_timeout()
        return max(t + 60, t * 2, 300)

    def effective_noop_timeout(self):
        t = self.effective_imap_timeout()
        return min(60, max(20, t // 10))

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
        从主题字符串中提取信件类型（如A123、B456、C789、A*123、A+B456、C风之帽等）。
        兼容下划线和空格分隔（如 FW:C_2485、FW:C 2485）；发信常见「C+称呼」无数字。
        """
        import re
        if not subject:
            return ""

        # 支持 Fw: / FW_ / FW： 后面的类型（允许下划线、空格、星号）
        m = re.search(
            r'(?i)Fw[:：_\s-]*\s*((?:A\s*\+\s*[BC]|B\s*\+\s*C|A|B|C)[\s_]*\*?\s*[\s_-]*\d+)',
            subject
        )
        if m:
            raw = m.group(1)
            # A_123 / A 123 → 视为 A*123
            if re.match(r'(?i)A[\s_]+\*?\s*\d+', raw) and '*' not in raw:
                raw = re.sub(r'(?i)A[\s_]+', 'A*', raw)
            return self.normalize_letter_type(re.sub(r'[\s_]+', '', raw))

        # 兜底：主题/路径中任意位置的数字类型（含「来自」前）
        m2 = re.search(
            r'(?i)(?<![A-Za-z+])((?:A\s*\+\s*[BC]|B\s*\+\s*C|A|B|C)[\s_]*\*?\s*[\s_-]*\d+)',
            subject
        )
        if m2:
            raw = m2.group(1)
            if re.match(r'(?i)A[\s_]+\*?\s*\d+', raw) and '*' not in raw:
                raw = re.sub(r'(?i)A[\s_]+', 'A*', raw)
            return self.normalize_letter_type(re.sub(r'[\s_]+', '', raw))

        # 发信常见：C风之帽 / CDD / C小菌+巧真（无数字；截到「来自」或路径边界）
        leaf = str(subject).replace("\\", "/").split("/")[-1]
        leaf = re.sub(r"^\d{8}_\d{6}_", "", leaf)  # 去时间戳前缀
        cleaned = re.sub(r"(?i)^(?:Fw|FW|转发)[:：_\s-]*", "", leaf).strip()
        cleaned = re.sub(r"来自.*$", "", cleaned).strip()
        cleaned = cleaned.strip("_- ")
        m3 = re.match(r"(?i)^((?:[AB]\s*\+\s*)?C)(.+)$", cleaned)
        if m3 and m3.group(2).strip():
            # 去掉文件夹里偶发的尾部噪声
            name = re.sub(r"[\s_]+", "", m3.group(2))
            name = re.sub(r"(的信|的)$", "", name)
            if name:
                return self.normalize_letter_type(m3.group(1) + name)

        # 广播主题：单独的 A / B / A+B（发 A 不入库，仅用于识别）
        m4 = re.match(r"(?i)^(A|B|A\s*\+\s*B)$", cleaned)
        if m4:
            return self.normalize_letter_type(m4.group(1))

        return ""

    def _safe_filename(self, name):
        """生成安全的文件名"""
        return re.sub(r'[\\/:*?"<>|]', '_', name)

    def _parse_to_datetime(self, date_val):
        """把 date_obj / 日期字符串解析为 datetime；失败返回 None。"""
        if date_val is None or date_val == "":
            return None
        try:
            if isinstance(date_val, datetime):
                return date_val
            if hasattr(date_val, 'strftime') and not isinstance(date_val, str):
                return date_val
            if isinstance(date_val, str):
                try:
                    return email.utils.parsedate_to_datetime(date_val)
                except Exception:
                    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%Y%m%d_%H%M%S", "%Y%m%d"):
                        try:
                            return datetime.strptime(date_val.strip(), fmt)
                        except Exception:
                            continue
        except Exception:
            return None
        return None

    def _date_folder_parts(self, date_val):
        """返回 (day_folder, fine_folder)。
        day: YYYYMMDD（旧格式）；fine: YYYYMMDD_HHMMSS（新格式，同日多封不撞车）。
        使用邮件头墙钟时间（不转 UTC），与旧目录日份保持一致。
        """
        dt = self._parse_to_datetime(date_val)
        if not dt:
            if isinstance(date_val, str) and date_val:
                m = re.search(r'(\d{4})[-/]?(\d{2})[-/]?(\d{2})', date_val)
                if m:
                    day = f"{m.group(1)}{m.group(2)}{m.group(3)}"
                    return day, day
            return "unknown_date", "unknown_date"
        try:
            day = dt.strftime("%Y%m%d")
            fine = dt.strftime("%Y%m%d_%H%M%S")
            return day, fine
        except Exception:
            return "unknown_date", "unknown_date"

    def _send_type_for_email(self, sender):
        current_email = self.email_addr.lower()
        sender = (sender or '').strip()
        return "发" if sender.lower() == current_email else "收"

    def _expected_email_dir(self, email_info):
        """新规范期望路径：发|收 / YYYYMMDD_HHMMSS_主题（同日同主题多封不冲突）。"""
        send_type = self._send_type_for_email(email_info.get('sender'))
        date_val = email_info.get('date_obj') or email_info.get('date')
        _, fine = self._date_folder_parts(date_val)
        subject = (email_info.get('subject') or '')[:30]
        safe = self._safe_filename(subject)
        return self.base_dir / send_type / f"{fine}_{safe}"

    def _content_date_matches(self, email_dir, email_info, tolerance_sec=2):
        """用 content.txt 的「日期:」与当前邮件时间比对，确认是否同一封（避免同日同主题误判）。
        允许极小误差（默认 2 秒），消化 IMAP envelope 与 Date 头的细微差异；
        同日多封通常间隔远大于此，不会误并。用 timestamp 比较，兼容带时区差异。
        """
        content_file = Path(email_dir) / "content.txt"
        if not content_file.exists():
            return False
        try:
            with open(content_file, 'r', encoding='utf-8') as f:
                for line in f:
                    if line.startswith("日期:"):
                        saved = line[len("日期:"):].strip()
                        break
                else:
                    return False
        except Exception:
            return False

        want_dt = self._parse_to_datetime(email_info.get('date_obj') or email_info.get('date'))
        got_dt = self._parse_to_datetime(saved)
        if not want_dt or not got_dt:
            return False
        # 1) 同一瞬间（处理好时区）
        try:
            if abs(want_dt.timestamp() - got_dt.timestamp()) <= tolerance_sec:
                return True
        except Exception:
            pass
        # 2) 墙钟一致（naive envelope 与带 +0800 的 Date 头常见组合）
        try:
            w = want_dt.replace(tzinfo=None)
            g = got_dt.replace(tzinfo=None)
            if abs((w - g).total_seconds()) <= tolerance_sec:
                return True
        except Exception:
            pass
        return False

    def _dir_download_complete(self, email_dir, email_info):
        """检查目录是否已有完整下载内容。返回 (ok, reason)。"""
        email_dir = Path(email_dir)
        if not email_dir.exists():
            return False, "目录不存在"

        content_file = email_dir / "content.txt"
        if not content_file.exists():
            return False, "content.txt 不存在"

        try:
            if content_file.stat().st_size <= 100:
                return False, "content.txt 文件过小"
        except Exception as e:
            return False, f"检查文件大小失败: {e}"

        try:
            with open(content_file, 'r', encoding='utf-8') as f:
                content = f.read()
            required_fields = ["主题:", "发件邮箱:", "收件邮箱:", "日期:"]
            missing = [field for field in required_fields if field not in content]
            if missing:
                return False, f"content.txt 缺少字段: {', '.join(missing)}"
        except Exception as e:
            return False, f"读取 content.txt 失败: {e}"

        if email_info.get('has_attachments', False):
            try:
                files_in_dir = [f for f in email_dir.iterdir() if f.is_file() and f.name != 'content.txt']
                if not files_in_dir:
                    return False, "附件缺失"
            except Exception as e:
                return False, f"检查附件失败: {e}"

        return True, "已完整（本地验证）"

    def _candidate_email_dirs(self, email_info):
        """新旧兼容候选目录：精确新路径、旧日级路径、以及同日同主题的其它细粒度目录。"""
        send_type = self._send_type_for_email(email_info.get('sender'))
        parent = self.base_dir / send_type
        date_val = email_info.get('date_obj') or email_info.get('date')
        day, fine = self._date_folder_parts(date_val)
        subject = (email_info.get('subject') or '')[:30]
        safe = self._safe_filename(subject)
        suffix = f"_{safe}"

        ordered = []
        seen = set()

        def _add(p):
            key = str(p)
            if key not in seen:
                seen.add(key)
                ordered.append(Path(p))

        _add(parent / f"{fine}_{safe}")
        _add(parent / f"{day}_{safe}")

        if parent.exists():
            try:
                for p in parent.iterdir():
                    if not p.is_dir():
                        continue
                    name = p.name
                    # 同日 + 同主题：旧 day_主题 / 新 day_HHMMSS_主题 / 撞车 day_HHMMSS_N_主题
                    if name.endswith(suffix) and name.startswith(day):
                        rest = name[: -len(suffix)] if suffix else name
                        if rest == day or rest == fine or re.fullmatch(
                            rf"{re.escape(day)}_\d{{6}}(?:_\d+)?", rest
                        ):
                            _add(p)
            except Exception:
                pass
        return ordered

    def _email_already_downloaded(self, email_info):
        """判断邮件是否已完整下载（仅检查本地文件）。
        新路径按秒区分同日多封；旧 YYYYMMDD_主题 目录仍可通过 content.txt 日期对上则视为已下载。
        """
        expected = self._expected_email_dir(email_info)
        candidates = self._candidate_email_dirs(email_info)
        last_reason = "目录不存在"
        day, fine = self._date_folder_parts(
            email_info.get('date_obj') or email_info.get('date')
        )
        subject = (email_info.get('subject') or '')[:30]
        safe = self._safe_filename(subject)
        legacy_name = f"{day}_{safe}"

        for email_dir in candidates:
            ok, reason = self._dir_download_complete(email_dir, email_info)
            if not ok:
                last_reason = reason
                continue
            # 无论新旧目录，都以 content.txt 日期确认是同一封（避免同日/同秒同主题误跳）
            if self._content_date_matches(email_dir, email_info):
                tag = "旧目录兼容" if email_dir.name == legacy_name else "本地验证"
                logger.info(f"邮件已完整下载（{tag}）: {email_dir}")
                return True, str(email_dir), f"已完整（{tag}）"
            last_reason = "同日同主题目录存在但日期不匹配（视为另一封）"

        return False, str(expected), last_reason

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
                'timeout': '300',
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
                timeout=self.effective_imap_timeout()
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
                            'message_id': self._bs_text(envelope.message_id).strip(),
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
                            'message_id': self._bs_text(envelope.message_id).strip(),
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

    def _noop_with_timeout(self, timeout_sec=None):
        """NOOP 检测；超时视为连接已死（弱网下 noop 也可能永远不返回）。"""
        if not self.client:
            return False
        if timeout_sec is None:
            timeout_sec = self.effective_noop_timeout()
        box = {"ok": False, "err": None, "done": False}

        def _run():
            try:
                self.client.noop()
                box["ok"] = True
            except Exception as e:
                box["err"] = e
            finally:
                box["done"] = True

        t = threading.Thread(target=_run, daemon=True)
        t.start()
        t.join(timeout_sec)
        if not box["done"]:
            logger.warning(f"连接检测超时（>{timeout_sec}s），丢弃连接")
            self._abandon_imap_client()
            return False
        if box["err"] is not None:
            raise box["err"]
        return box["ok"]

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
                        if self._noop_with_timeout():
                            return True
                        # noop 超时已丢弃 client
                    except Exception as e:
                        logger.warning(f"连接检测失败（尝试 {attempt}/{max_retries}）：{e}")
                        self._abandon_imap_client()
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
        
        # 日期（新存盘用到秒，兼容同日多封；旧日级目录仍可由已下载检查识别）
        date_str = msg['date'] or ''
        metadata['date_str'] = date_str
        
        try:
            date_obj = email.utils.parsedate_to_datetime(date_str)
            metadata['date_obj'] = date_obj
            day, fine = self._date_folder_parts(date_obj)
            metadata['date_folder'] = fine          # 新：YYYYMMDD_HHMMSS
            metadata['date_folder_day'] = day       # 旧：YYYYMMDD
        except Exception:
            metadata['date_obj'] = None
            metadata['date_folder'] = "unknown_date"
            metadata['date_folder_day'] = "unknown_date"
        
        return metadata
    
    # ========== 3. 确定存储目录 ==========
    def _determine_email_directory(self, metadata):
        """确定邮件存储目录（新：YYYYMMDD_HHMMSS_主题；若旧日级目录已是同一封则复用）。
        
        Returns:
            Path: 邮件存储目录路径
        """
        send_type = self._send_type_for_email(metadata.get('sender'))
        subject = (metadata.get('subject') or '')[:30]
        safe_subject = self._safe_filename(subject)
        fine = metadata.get('date_folder') or "unknown_date"
        day = metadata.get('date_folder_day') or fine.split('_')[0]

        email_info = {
            'sender': metadata.get('sender'),
            'subject': metadata.get('subject'),
            'date': metadata.get('date_str'),
            'date_obj': metadata.get('date_obj'),
        }

        # 若本地已有同一封（含旧 YYYYMMDD_主题），直接复用，避免重复建夹
        already, existing, _ = self._email_already_downloaded(email_info)
        if already and existing:
            return Path(existing)

        email_dir = self.base_dir / send_type / f"{fine}_{safe_subject}"
        # 极端：同一秒同主题撞车且 content 不是同一封 → 加序号
        if email_dir.exists() and not self._content_date_matches(email_dir, email_info):
            n = 2
            while True:
                alt = self.base_dir / send_type / f"{fine}_{n}_{safe_subject}"
                if not alt.exists() or self._content_date_matches(alt, email_info):
                    email_dir = alt
                    break
                n += 1
                if n > 99:
                    break

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

    @staticmethod
    def _bs_text(v):
        if isinstance(v, bytes):
            return v.decode("utf-8", errors="replace")
        return "" if v is None else str(v)

    def _walk_bodystructure(self, bs, prefix=""):
        """展开 IMAP BODYSTRUCTURE → [(段号, 'text/plain', {参数}, 编码, 是否附件, 文件名)]。"""
        parts = []
        if getattr(bs, "is_multipart", False):
            for i, sub in enumerate(bs[0], 1):
                parts.extend(self._walk_bodystructure(sub, f"{prefix}.{i}" if prefix else str(i)))
            return parts
        fields = list(bs)
        ctype = f"{self._bs_text(fields[0]).lower()}/{self._bs_text(fields[1]).lower()}"
        params = {}
        raw_params = fields[2] if len(fields) > 2 and isinstance(fields[2], (tuple, list)) else ()
        for k, v in zip(raw_params[::2], raw_params[1::2]):
            params[self._bs_text(k).lower()] = self._bs_text(v)
        encoding = self._bs_text(fields[5]).lower() if len(fields) > 5 else ""
        disposition, filename = "", params.get("name", "")
        for f in fields[7:]:
            if isinstance(f, (tuple, list)) and f and isinstance(f[0], bytes) \
                    and f[0].lower() in (b"attachment", b"inline"):
                disposition = self._bs_text(f[0]).lower()
                dparams = f[1] if len(f) > 1 and isinstance(f[1], (tuple, list)) else ()
                for k, v in zip(dparams[::2], dparams[1::2]):
                    if self._bs_text(k).lower().startswith("filename"):
                        filename = self._bs_text(v)
                break
        filename = self.decode_mime_header(filename) if filename else ""
        is_attachment = disposition == "attachment" or bool(filename)
        parts.append((prefix or "1", ctype, params, encoding, is_attachment, filename))
        return parts

    def fetch_email_preview(self, email_id, folder):
        """只取正文文本用于预览，不拉附件内容。返回 (正文, 附件名列表, 错误)。"""
        import base64
        import quopri
        try:
            if not self.ensure_connected():
                return "", [], "无法连接到邮箱"
            self.client.select_folder(folder, readonly=True)
            data = self.client.fetch([email_id], ["BODYSTRUCTURE"])
            bs = (data.get(email_id) or {}).get(b"BODYSTRUCTURE")
            if bs is None:
                return "", [], f"找不到邮件 {email_id}"
            parts = self._walk_bodystructure(bs)
            attachments = [p[5] or p[1] for p in parts if p[4]]
            body_parts = [p for p in parts if not p[4] and p[1] in ("text/plain", "text/html")]
            if not body_parts:
                return "", attachments, None
            part = next((p for p in body_parts if p[1] == "text/plain"), body_parts[0])
            part_no, ctype, params, encoding = part[:4]
            resp = self.client.fetch([email_id], [f"BODY.PEEK[{part_no}]"])
            item = resp.get(email_id) or {}
            raw = next((v for k, v in item.items() if k.startswith(b"BODY[")), b"") or b""
            if encoding == "base64":
                raw = base64.b64decode(raw)
            elif encoding == "quoted-printable":
                raw = quopri.decodestring(raw)
            charset = (params.get("charset") or "utf-8").lower()
            if charset in ("gb2312", "gbk"):
                charset = "gb18030"
            try:
                text = raw.decode(charset, errors="replace")
            except LookupError:
                text = raw.decode("utf-8", errors="replace")
            if ctype == "text/html":
                text = re.sub(r"(<br\s*/?>|<p>|</p>|<div[^>]*>|</div>)", "\n", text, flags=re.I)
                text = re.sub(r"<[^<]+?>", "", text.replace("&nbsp;", " "))
                text = re.sub(r"\n+", "\n", text)
            return text, attachments, None
        except Exception as e:
            return "", [], f"预览失败: {e}"
    
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
    def _abandon_imap_client(self):
        """丢弃可能已卡死的 IMAP 连接（弱网下 fetch/noop 可能永不返回）。"""
        client = self.client
        self.client = None
        if not client:
            return
        def _close():
            try:
                client.logout()
            except Exception:
                try:
                    sock = getattr(getattr(client, "_imap", None), "sock", None)
                    if sock:
                        sock.close()
                except Exception:
                    pass
        threading.Thread(target=_close, daemon=True).start()

    def _download_email_with_timeout(self, email_id, folder, timeout_sec=None):
        """带超时的单封下载。超时或用户取消则丢弃连接并返回失败，避免整批卡死。"""
        if timeout_sec is None:
            timeout_sec = self.effective_email_watchdog()
        box = {"result": None, "error": None, "done": False}

        def _run():
            try:
                res, err = self.download_email(email_id, folder)
                box["result"], box["error"] = res, err
            except Exception as e:
                box["result"], box["error"] = None, str(e)
            finally:
                box["done"] = True

        t = threading.Thread(target=_run, daemon=True)
        t.start()
        deadline = time.time() + max(1, float(timeout_sec))
        while not box["done"]:
            if self.cancel_requested:
                logger.info(f"用户取消：中断当前下载 id={email_id}")
                self._abandon_imap_client()
                t.join(2)
                return None, "用户取消"
            remaining = deadline - time.time()
            if remaining <= 0:
                break
            t.join(min(0.4, remaining))

        if not box["done"]:
            logger.error(
                f"下载超时（>{timeout_sec}s）id={email_id}，丢弃 IMAP 连接以免整批卡住"
            )
            self._abandon_imap_client()
            t.join(2)
            return None, f"下载超时（>{timeout_sec}s，弱网/大附件常见）"
        if self.cancel_requested and not box["result"]:
            return None, "用户取消"
        return box["result"], box["error"]

    def download_email(self, email_id, folder):
        """下载单个邮件(主流程)
        
        Returns:
            tuple: (result字典, 错误信息)
        """
        if not self.client:
            if not self.connect_to_email():
                return None, "连接失败"

        try:
            msg, error = self._fetch_raw_email(email_id, folder)
            if error:
                return None, error

            metadata = self._parse_email_metadata(msg)
            email_dir = self._determine_email_directory(metadata)
            to_addrs = self._parse_recipients(msg)
            body_text = self._extract_email_body(msg)
            attachment_paths, attachment_errors = self._download_attachments(msg, email_dir, email_id)

            success, error = self._save_content_file(email_dir, metadata, to_addrs, body_text)
            if not success:
                return None, error

            success, error = self._validate_download(email_dir, attachment_errors)
            if not success:
                return None, error

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

    def download_multiple_emails(
        self,
        email_list,
        progress_callback=None,
        soft_attempts=2,
        hard_attempts=10,
        mid_phase_callback=None,
        on_result_callback=None,
    ):
        """两阶段下载：主轮每封最多 soft 次后暂挂 → 中期落盘 → 后台补到 hard 次。

        progress_callback(progress, current, total, email_info)
        mid_phase_callback(results, deferred_list, failed_details, cancelled) — 主轮结束时
        on_result_callback(results, failed_details, stage) — 后台每成功一封可用来增量落盘
        """
        self.reset_cancel()
        results = []
        total = len(email_list)

        def _k(e):
            # IMAP UID 只在单个文件夹内唯一
            return (e.get('folder'), e['id'])

        index_by_id = {_k(e): i for i, e in enumerate(email_list)}
        succeeded_ids = set()
        skipped_ids = set()
        attempt_count = {_k(e): 0 for e in email_list}
        failed_details = []
        eff_timeout = self.effective_email_watchdog()
        soft_attempts = max(1, int(soft_attempts))
        hard_attempts = max(soft_attempts, int(hard_attempts))

        def _cancelled():
            return bool(self.cancel_requested)

        def _attach_user_overrides(email_info, res):
            for key in ('user_letter_type', 'user_sender_name', 'user_receiver_name', 'user_direction'):
                if email_info.get(key):
                    res[key] = email_info[key]
            return res

        def _notify(email_info, status, reason="", attempt=None):
            email_info['status'] = status
            email_info['reason'] = reason
            if attempt is not None:
                email_info['attempt'] = attempt
            if not progress_callback:
                return
            idx = index_by_id.get(_k(email_info), 0)
            done = len(succeeded_ids) + len(skipped_ids)
            current = min(total, max(done, idx + 1))
            progress = int(current / total * 100) if total else 100
            progress_callback(progress, current, total, email_info)

        def _mark_skip(email_info, email_dir, reason):
            eid = _k(email_info)
            if eid in succeeded_ids or eid in skipped_ids:
                return
            results.append(_attach_user_overrides(email_info, {
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
            skipped_ids.add(eid)

        def _mark_success(email_info, result, reason):
            eid = _k(email_info)
            if eid in succeeded_ids or eid in skipped_ids:
                return
            result['status'] = 'success'
            result['reason'] = reason
            results.append(_attach_user_overrides(email_info, result))
            succeeded_ids.add(eid)

        def _upsert_fail(email_info, error):
            key = _k(email_info)
            for fd in failed_details:
                if fd.get('key') == key:
                    fd['error'] = error
                    fd['attempts'] = attempt_count.get(key, 0)
                    return
            failed_details.append({
                'id': email_info['id'],
                'key': key,
                'subject': email_info.get('subject', ''),
                'error': error,
                'attempts': attempt_count.get(key, 0),
            })

        def _clear_fail(email_info):
            nonlocal failed_details
            key = _k(email_info)
            failed_details = [f for f in failed_details if f.get('key') != key]

        def _try_once(email_info, attempt_no, label_prefix):
            """尝试一次下载；返回 'success'|'skipped'|'fail'。"""
            email_id = email_info['id']
            folder = email_info['folder']
            already, email_dir, reason = self._email_already_downloaded(email_info)
            if already:
                _mark_skip(email_info, email_dir, reason)
                _notify(email_info, 'skipped', reason, attempt=attempt_no)
                return 'skipped'

            if attempt_no <= 1:
                _notify(email_info, 'processing', '处理中', attempt=attempt_no)
            else:
                _notify(email_info, 'retry', f'{label_prefix}{attempt_no}', attempt=attempt_no)

            res, error = self._download_email_with_timeout(
                email_id, folder, timeout_sec=eff_timeout
            )
            if res:
                _mark_success(email_info, res, f'{label_prefix}{attempt_no}成功' if attempt_no > 1 else '下载成功')
                _clear_fail(email_info)
                _notify(email_info, 'success', email_info.get('reason', '已下载'), attempt=attempt_no)
                return 'success'

            already2, email_dir2, reason2 = self._email_already_downloaded(email_info)
            if already2:
                _mark_skip(email_info, email_dir2, reason2)
                _clear_fail(email_info)
                _notify(email_info, 'skipped', reason2, attempt=attempt_no)
                return 'skipped'

            _upsert_fail(email_info, error or '失败')
            return 'fail'

        # ---------- 主轮：每封最多 soft_attempts 次，失败则暂挂 ----------
        deferred = []
        for i, email_info in enumerate(email_list):
            if _cancelled():
                logger.info("用户取消：主轮提前结束")
                break
            email_id = email_info['id']
            key = _k(email_info)
            already, email_dir, reason = self._email_already_downloaded(email_info)
            if already:
                logger.info(f"✓ 跳过已下载 id={email_id}")
                _mark_skip(email_info, email_dir, reason)
                _notify(email_info, 'skipped', reason, attempt=0)
                continue

            ok = False
            for attempt_no in range(1, soft_attempts + 1):
                if _cancelled():
                    break
                attempt_count[key] = attempt_no
                logger.info(
                    f"→ 主轮尝试{attempt_no}/{soft_attempts} id={email_id} "
                    f"主题={email_info.get('subject')}"
                )
                if attempt_no > 1:
                    self._abandon_imap_client()
                    if not self.ensure_connected():
                        _upsert_fail(email_info, '重连失败')
                        _notify(email_info, 'retry', f'失败重试{attempt_no}', attempt=attempt_no)
                        time.sleep(1.0)
                        continue
                outcome = _try_once(email_info, attempt_no, '失败重试')
                if _cancelled():
                    break
                if outcome in ('success', 'skipped'):
                    ok = True
                    break
                _notify(email_info, 'retry', f'失败重试{attempt_no}', attempt=attempt_no)
                # 可中断等待，避免点了「以当前进度结束」仍干等重试间隔
                wait_left = 0.8 if attempt_no < soft_attempts else 0.3
                while wait_left > 0 and not _cancelled():
                    step = min(0.2, wait_left)
                    time.sleep(step)
                    wait_left -= step
            if not ok and key not in succeeded_ids and key not in skipped_ids:
                _notify(email_info, 'deferred', f'暂挂（已试{soft_attempts}次）', attempt=soft_attempts)
                deferred.append(email_info)
            if i < total - 1 and not _cancelled():
                time.sleep(0.3)

        # 用户取消时：把未处理完且本地仍不完整的也算暂挂（不再主轮硬失败）
        if _cancelled():
            for email_info in email_list:
                eid = _k(email_info)
                if eid in succeeded_ids or eid in skipped_ids:
                    continue
                if email_info not in deferred:
                    deferred.append(email_info)
                    _notify(email_info, 'deferred', '任务已提前结束', attempt=attempt_count.get(eid, 0))

        cancelled = _cancelled()
        logger.info(
            f"主轮结束：成功{len(succeeded_ids)} 跳过{len(skipped_ids)} "
            f"暂挂{len(deferred)} 取消={cancelled}"
        )
        if mid_phase_callback:
            try:
                mid_phase_callback(list(results), list(deferred), list(failed_details), cancelled)
            except Exception as e:
                logger.error(f"中期回调失败: {e}")

        if cancelled:
            # 提前结束：不再后台补下
            for email_info in deferred:
                if _k(email_info) in succeeded_ids or _k(email_info) in skipped_ids:
                    continue
                email_info['status'] = 'failed'
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
                    'reason': '任务提前结束，未完成',
                })
                _notify(email_info, 'failed', '提前结束未完成')
            return {
                'results': results,
                'failed_details': failed_details,
                'deferred': [],
                'cancelled': True,
                'final_failed': deferred,
            }

        # ---------- 后台补下：累计到 hard_attempts ----------
        still = list(deferred)
        while still and not _cancelled():
            progressed = False
            self._abandon_imap_client()
            if not self.ensure_connected():
                logger.warning("后台补下：无法重连，稍后重试")
                wait_left = 3.0
                while wait_left > 0 and not _cancelled():
                    step = min(0.25, wait_left)
                    time.sleep(step)
                    wait_left -= step
                continue

            for email_info in list(still):
                if _cancelled():
                    break
                email_id = email_info['id']
                key = _k(email_info)
                if key in succeeded_ids or key in skipped_ids:
                    still.remove(email_info)
                    continue
                next_attempt = attempt_count.get(key, soft_attempts) + 1
                if next_attempt > hard_attempts:
                    continue
                attempt_count[key] = next_attempt
                logger.info(
                    f"→ 后台补下 尝试{next_attempt}/{hard_attempts} id={email_id} "
                    f"主题={email_info.get('subject')}"
                )
                outcome = _try_once(email_info, next_attempt, '失败重试')
                if _cancelled():
                    break
                if outcome in ('success', 'skipped'):
                    still.remove(email_info)
                    progressed = True
                    if on_result_callback:
                        try:
                            on_result_callback(list(results), list(failed_details), 'background')
                        except Exception as e:
                            logger.error(f"增量落盘回调失败: {e}")
                else:
                    _notify(email_info, 'retry', f'失败重试{next_attempt}', attempt=next_attempt)
                wait_left = 2.0
                while wait_left > 0 and not _cancelled():
                    step = min(0.25, wait_left)
                    time.sleep(step)
                    wait_left -= step

            # 去掉已达上限的
            remain = []
            for email_info in still:
                if attempt_count.get(_k(email_info), 0) >= hard_attempts:
                    email_info['status'] = 'failed'
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
                        'reason': f'多次尝试后失败（{hard_attempts}次）',
                    })
                    _notify(email_info, 'failed', f'失败(已达{hard_attempts}次)', attempt=hard_attempts)
                else:
                    remain.append(email_info)
            still = remain
            if not still:
                break
            if not progressed:
                time.sleep(2.5)

        final_failed = [
            e for e in email_list
            if _k(e) not in succeeded_ids and _k(e) not in skipped_ids
        ]
        for email_info in final_failed:
            if email_info.get('status') != 'failed':
                email_info['status'] = 'failed'
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
                    'reason': f'多次尝试后失败（{hard_attempts}次）',
                })
                _notify(email_info, 'failed', f'失败(已达{hard_attempts}次)')

        logger.info(
            f"下载结束：成功{len(succeeded_ids)} 跳过{len(skipped_ids)} "
            f"最终失败{len(final_failed)}"
        )
        return {
            'results': results,
            'failed_details': failed_details,
            'deferred': deferred,
            'cancelled': bool(self.cancel_requested),
            'final_failed': final_failed,
        }
    

    @staticmethod
    def extract_letter_sender_name(content_text="", directory="", mail_sender="", subject=""):
        """从正文/路径/主题解析「角色发件人」名，不要用中转站邮箱顶替。

        兼容：
        - 发件人=【甲】 / 发件人＝【甲】 / 发件人 =【甲】
        - 发件人【甲】（无等号，如 T.S.E）
        - 发件人=【【甲】】（嵌套方括号，如 4o4）
        - 路径 …来自甲的信 / …来自…的 / …自甲的信
        """
        def _looks_like_email(s):
            s = (s or "").strip()
            return bool(re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", s))

        def _finalize(name, content_text="", directory="", subject=""):
            name = (name or "").strip()
            if not name:
                return ""
            # 剥掉多余嵌套 【【x】】 → 【x】；普通 【x】 默认去括号
            while name.startswith("【") and name.endswith("】") and name.count("【") >= 2:
                name = name[1:-1].strip()
            keep_brackets = bool(
                re.search(r"【", str(directory or ""))
                or re.search(r"【", str(subject or ""))
                or re.search(r"请称呼我为\s*【", content_text or "")
                or re.search(r"带方括号", content_text or "")
            )
            if name.startswith("【") and name.endswith("】") and name.count("【") == 1:
                if keep_brackets:
                    pass
                else:
                    name = name[1:-1].strip()
            name = re.sub(r"\s+", " ", name).strip()
            return name

        sender_name = ""
        if content_text:
            # 1) 发件人[=＝:]?【…】 或 发件人【…】；【+…】+ 吃掉嵌套
            m = re.search(r"发件人\s*[=＝:]?\s*(【+.+?】+)", content_text)
            if not m:
                m = re.search(r"发件[^\n【]{0,8}[=＝:]\s*(【+.+?】+)", content_text)
            if m:
                sender_name = _finalize(m.group(1), content_text, directory, subject)
            if not sender_name:
                # 2) 无【】：发件人=甲 / 发件人：甲（一行内）
                m = re.search(r"发件人\s*[=＝:：]\s*([^\n\r【]+)", content_text)
                if m:
                    cand = m.group(1).strip()
                    if cand and not cand.startswith("邮箱") and not _looks_like_email(cand):
                        sender_name = _finalize(cand, content_text, directory, subject)

        if not sender_name:
            blob = " ".join([str(directory or ""), str(subject or "")])
            m = re.search(r"来自(.+?)的(?:信)?", blob)
            if not m:
                m = re.search(r"(?<![来])自(.+?)的(?:信)?", blob)
            if m:
                raw = m.group(1).replace("_", "/")
                # 路径里自带的 【name】 原样交给 finalize
                if raw.startswith("【"):
                    sender_name = _finalize(raw if raw.endswith("】") else raw + "】", content_text, directory, subject)
                else:
                    sender_name = _finalize(raw, content_text, directory, subject)

        if not sender_name and mail_sender and not _looks_like_email(mail_sender):
            sender_name = _finalize(mail_sender, content_text, directory, subject)

        if not sender_name:
            sender_name = "未知"
        return sender_name

    @staticmethod
    def extract_letter_receiver_name(content_text="", letter_type="", send_type="", only_oc_name=""):
        """解析收件人（角色名）。

        优先：
          1) 发信且类型为 C… → 取 C 后称呼
          2) 正文「致【甲】」「致甲：」「致 甲：」
          3) 收信且全库仅一个 OC → 该 OC
        """
        if send_type == "发":
            p = EmailDownloader.parse_letter_type(letter_type)
            if p and "C" in p["kinds"] and p["target"]:
                return re.sub(r"[\s=\[\]]+", "", p["target"])

        text = content_text or ""
        if text:
            m = re.search(r"致\s*(【+.+?】+)", text)
            if m:
                name = m.group(1).strip()
                while name.startswith("【") and name.endswith("】") and name.count("【") >= 1:
                    # 单层去括号；多层剥到内层再去
                    inner = name[1:-1].strip()
                    if inner.startswith("【"):
                        name = inner
                        continue
                    name = inner
                    break
                name = re.sub(r"\s+", " ", name).strip()
                if name:
                    return name
            m = re.search(r"致\s*([^\n\r【：:]{1,40})\s*[:：]", text)
            if m:
                name = re.sub(r"\s+", " ", m.group(1)).strip()
                # 排除「致此」「致敬」等
                if name and name not in ("此", "敬", "辞") and not name.startswith("信"):
                    return name

        if send_type == "收" and only_oc_name:
            return only_oc_name
        return ""

    @staticmethod
    def parse_letter_type(letter_type):
        """拆解信件类型 → {"kinds": "A+C", "star": "", "number": "12", "target": ""}；无法识别返回 None。

        组合类型 A+B / A+C / B+C 的顺序、大小写、全角＋均可，统一排成 A→B→C。
        后缀为纯数字（可带 *、_、-）视为编号，否则视为发 C 的对象称呼。
        """
        raw = (letter_type or "").strip()
        s = raw.replace("＋", "+").replace("＊", "*")
        m = re.match(r"^([ABC](?:\s*\+\s*[ABC])*)(.*)$", s, re.I | re.S)
        if not m:
            return None
        kind_set = {k.upper() for k in re.findall(r"[ABC]", m.group(1), re.I)}
        kinds = "+".join(k for k in "ABC" if k in kind_set)
        rest = m.group(2).strip()
        mnum = re.fullmatch(r"(\*?)[\s_\-·]*(\d+)", rest)
        if mnum:
            star = mnum.group(1)
            # 历史约定：主题/路径里 A_123 常表示 A*123
            if not star and re.search(r"A[\s_]", raw, re.I):
                star = "*"
            return {"kinds": kinds, "star": star, "number": mnum.group(2), "target": ""}
        target = re.sub(r"^[\s_=:：]+", "", rest).strip()
        return {"kinds": kinds, "star": "", "number": "", "target": target}

    @staticmethod
    def normalize_letter_type(letter_type):
        """统一信件类型：c3594→C3594，A_164→A*164，b+c 12→B+C12，c+a小菌→A+C小菌。"""
        s = (letter_type or "").strip()
        if not s:
            return s
        p = EmailDownloader.parse_letter_type(s)
        if not p:
            return re.sub(r"\s+", "", s)
        if p["number"]:
            return f"{p['kinds']}{p['star']}{p['number']}"
        return p["kinds"] + p["target"]

    @staticmethod
    def check_letter_type(letter_type, send_type=""):
        """按收/发方向校验并规范类型，返回 (规范后类型, 错误提示, 由类型推出的收件人)。

        收：一律带编号（A123 / B45 / C678 / A+B12 / A+C3 / B+C7）。
        发：单 A 不入库；B、A+B 不带后缀；含 C 的必须写对象（C某某 / A+C某某 / B+C某某）。
        """
        p = EmailDownloader.parse_letter_type(letter_type)
        if not p:
            return "", "无法识别信件类型。示例：收 A123 / B+C12，发 B / C某某 / A+C某某。", ""
        kinds = p["kinds"]
        if send_type == "收":
            if not p["number"]:
                return "", "收信类型必须带编号，如 A123、A*12、B45、C678、A+B12、A+C3、B+C7。", ""
            return f"{kinds}{p['star']}{p['number']}", "", ""
        if send_type == "发":
            if kinds == "A":
                return "", "发 A（广播）不入库：等广播回投后按「收 A编号」记录。", ""
            if "C" in kinds:
                target = p["target"] or p["number"]
                if not target:
                    return "", "发 C 必须写对象，如 C某某、A+C某某、B+C某某。", ""
                return kinds + target, "", target
            return kinds, "", ""
        if not EmailDownloader.is_valid_letter_type(letter_type):
            return "", "信件类型不合法。示例：A123、A*123、B45、C678、A+B1、B+C2、C某某。", ""
        return EmailDownloader.normalize_letter_type(letter_type), "", ""

    def repair_tsv_sender_names(self, tsv_path=None):
        """纠正 global_result 发件人（邮箱/未知/破损括号）并规范化信件类型。"""
        path = Path(tsv_path) if tsv_path else (self.base_dir / "global_result.tsv")
        if not path.exists():
            return 0
        try:
            from uitl_test import _resolve_email_dir
        except Exception:
            _resolve_email_dir = None

        lines = path.read_text(encoding="utf-8").splitlines()
        data_start = -1
        for i, line in enumerate(lines):
            if line.startswith("收/发类型") or line.startswith("收/发\t"):
                data_start = i
                break
        if data_start < 0:
            return 0

        def _needs_sender_fix(sender):
            s = (sender or "").strip()
            if not s or s == "未知":
                return True
            if re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", s):
                return True
            # 括号不配对 / 残留嵌套半截
            if s.count("【") != s.count("】"):
                return True
            if s.startswith("【") and not s.endswith("】"):
                return True
            return False

        changed = 0
        out = lines[: data_start + 1]
        for line in lines[data_start + 1:]:
            cols = line.split("\t")
            if len(cols) < 6:
                out.append(line)
                continue
            sender = (cols[2] or "").strip()
            letter_type = (cols[4] or "").strip()
            directory = cols[5]
            row_changed = False

            content_text = ""
            email_dir = None
            if _resolve_email_dir:
                email_dir = _resolve_email_dir(directory, path)
            if email_dir is None:
                rel = directory[2:] if str(directory).startswith(".\\") else directory
                cand = script_dir / str(rel).replace("\\", os.sep)
                if cand.is_dir():
                    email_dir = cand
            subject_hint = ""
            if email_dir:
                cf = Path(email_dir) / "content.txt"
                if cf.exists():
                    try:
                        content_text = cf.read_text(encoding="utf-8")
                        first = content_text.splitlines()[:1]
                        if first:
                            subject_hint = re.sub(r"^主题\s*[:：]\s*", "", first[0]).strip()
                    except Exception:
                        content_text = ""

            # 补信件类型：数字型 A/B/C + 发信 C称呼
            inferred_type = self.normalize_letter_type(self.extract_letter_type(letter_type))
            if not inferred_type:
                inferred_type = self.normalize_letter_type(self.extract_letter_type(directory))
            if not inferred_type and subject_hint:
                inferred_type = self.normalize_letter_type(self.extract_letter_type(subject_hint))
            if inferred_type and inferred_type != letter_type:
                cols[4] = inferred_type
                letter_type = inferred_type
                row_changed = True
            else:
                norm_type = self.normalize_letter_type(letter_type)
                if norm_type and norm_type != letter_type:
                    cols[4] = norm_type
                    letter_type = norm_type
                    row_changed = True

            parsed = self.extract_letter_sender_name(
                content_text=content_text,
                directory=directory,
                mail_sender=sender,
                subject=subject_hint,
            )
            if _needs_sender_fix(sender) or (parsed and parsed != "未知" and parsed != sender):
                # 仅当旧值明显坏掉，或正文/路径能给出更可信名字时覆盖
                if _needs_sender_fix(sender) or (
                    parsed
                    and parsed != "未知"
                    and (
                        "发件人" in (content_text or "")
                        or "发件者" in (content_text or "")
                        or "来自" in str(directory)
                        or re.search(r"(?<![来])自.+的", str(directory) or "")
                    )
                ):
                    if parsed and parsed != sender:
                        logger.info(f"纠正发件人: {sender!r} → {parsed!r} ({directory})")
                        cols[2] = parsed
                        row_changed = True
                    elif sender != (cols[2] or "").strip():
                        cols[2] = sender
                        row_changed = True
            elif cols[2] != sender:
                cols[2] = sender
                row_changed = True

            # 补收件人：发信 C… / 收信「致…」
            direction = (cols[0] or "").strip()
            receiver = (cols[3] or "").strip()
            if not receiver:
                inferred_recv = self.extract_letter_receiver_name(
                    content_text=content_text,
                    letter_type=letter_type,
                    send_type=direction,
                    only_oc_name="",
                )
                if inferred_recv:
                    cols[3] = inferred_recv
                    row_changed = True

            if row_changed:
                changed += 1
            out.append("\t".join(cols))

        if changed:
            path.write_text("\n".join(out) + "\n", encoding="utf-8")
            logger.info(f"已纠正 {changed} 行写入 {path}")
        return changed

    # 所有涉及收发件人推断的操作在此步进行
    def generate_download_report(self, results, email_list, failed_details=None, stage="最终汇总"):
        """生成报告，TSV 去重，记录失败详情。仅登记合法结果（跳过 status=failed）。"""
        report_path = self.base_dir / "download_report.txt"
        global_result_path = self.base_dir / "global_result.tsv"
        current_email = self.email_addr.lower()
        # 过滤：失败项不写 TSV；目录必须像本地路径
        results = [
            r for r in (results or [])
            if r.get('status') != 'failed' and r.get('directory') not in (None, '', 'FAILED')
        ]

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
                            # 【新增】：提取并保护旧别名数据，同时并入内存
                            existing_aliases_line = line.strip()
                            cols = line.strip().split("\t")
                            if len(cols) > 1 and cols[1]:
                                for group in cols[1].split(","):
                                    parts = [p.strip() for p in group.split("=") if p.strip()]
                                    if len(parts) > 1:
                                        main, rest = parts[0], parts[1:]
                                        merged = list(self.oc_aliases.get(main, []) or [])
                                        for a in rest:
                                            if a not in merged and a != main:
                                                merged.append(a)
                                        self.oc_aliases[main] = merged
                        elif line.startswith("收/发类型") or line.startswith("收/发"):
                            data_start_idx = i
                            break
                            
                    # 解析邮件（关键：用更多字段构建唯一键）；丢弃历史误入库的「发 A」广播
                    for line in lines[data_start_idx+1:]:
                        cols = line.strip().split('\t')
                        if len(cols) >= 6:
                            lt = (cols[4] or "").strip()
                            if not lt:
                                lt = self.extract_letter_type(cols[5]) or self.extract_letter_type(cols[5].replace("\\", "/").split("/")[-1])
                            if (cols[0] or "").strip() == "发" and self.is_bare_broadcast_a(lt):
                                continue
                            # 唯一标识：发件人 + 日期 + 主题/类型 + 路径
                            key = (cols[2], cols[1], cols[4], cols[5])
                            existing_mails.add(key)
                            existing_rows.append(line.strip())

        # 2. 合并 OC（循环中可能继续自动登记发信 OC，故保持 set，写盘前再排序）
        all_oc = set(existing_oc)
        all_oc.update(self.oc_registerdata)
        only_oc_name = None
        if len(all_oc) == 1:
            only_oc_name = next(iter(all_oc)).split('_')[0]

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

            # 补全收件人发件人（角色名，不要用 crackpost@… 中转邮箱）
            content_path = Path(result['directory']) / "content.txt"
            content_text = ""
            if content_path.exists():
                try:
                    content_text = content_path.read_text(encoding="utf-8")
                except Exception:
                    content_text = ""
            sender_name = self.extract_letter_sender_name(
                content_text=content_text,
                directory=str(result.get("directory", "")),
                mail_sender=mail_sender,
                subject=str(result.get("subject", "")),
            )

            subject = result.get('subject', '')
            subject = re.sub(r'\s+', '_', subject)
            letter_type = self.normalize_letter_type(self.extract_letter_type(subject))
            if not letter_type:
                # 主题失败时再从目录名兜底
                letter_type = self.normalize_letter_type(
                    self.extract_letter_type(str(result.get("directory", "")))
                )
            if not letter_type:
                print(f"[DEBUG] letter_type解析失败，输入subject: '{subject}' 路径: {result.get('directory','')}")

            # 用户指定覆盖须尽早应用（以便「发 A」跳过判断用到最终类型）
            if result.get('user_letter_type'):
                letter_type = result['user_letter_type']
            if result.get('user_sender_name'):
                sender_name = result['user_sender_name']
            if result.get('user_direction') in ("收", "发"):
                send_type = result['user_direction']

            # 发广播 A：只记录 CrackPost 回投的收 A，不入库用户发出的裸 A
            if send_type == "发" and self.is_bare_broadcast_a(letter_type):
                logger.info(f"跳过发广播 A（不入库）: {result.get('directory')}")
                continue

            receiver_name = self.extract_letter_receiver_name(
                content_text=content_text,
                letter_type=letter_type,
                send_type=send_type,
                only_oc_name=only_oc_name if send_type == "收" else "",
            )
            if result.get('user_receiver_name'):
                receiver_name = result['user_receiver_name']

            # 发信：正文「发件人=【角色】」自动登记为本邮箱 OC 候选（?/符号 ID 也合法）
            if send_type == "发":
                oc_from_send = sender_name
                if self.is_plausible_role_name(oc_from_send):
                    stamp = date_short if re.match(r"\d{4}-\d{2}-\d{2}$", date_short or "") else ""
                    name_to_entry = {}
                    for x in list(self.oc_registerdata) + list(all_oc):
                        xm = re.match(r"^(.*)_(\d{4}-\d{2}-\d{2})$", str(x))
                        if xm:
                            n, d = xm.group(1), xm.group(2)
                        else:
                            n, d = str(x), ""
                        if not n:
                            continue
                        prev = name_to_entry.get(n)
                        if not prev or (d and (not prev[1] or d < prev[1])):
                            name_to_entry[n] = (x, d)
                    added_alias = self.ensure_dot_aliases(oc_from_send)
                    if added_alias:
                        logger.info(f"自动登记 OC 分段别名: {oc_from_send} → {added_alias}")
                    if oc_from_send not in name_to_entry:
                        entry = f"{oc_from_send}_{stamp}" if stamp else oc_from_send
                        self.oc_registerdata.append(entry)
                        all_oc.add(entry)
                        logger.info(f"自动登记 OC（发信解析）: {entry}")
                    else:
                        old_entry, old_d = name_to_entry[oc_from_send]
                        if stamp and (not old_d or stamp < old_d):
                            new_entry = f"{oc_from_send}_{stamp}"
                            all_oc.discard(old_entry)
                            all_oc.add(new_entry)
                            def _oc_name(y):
                                ym = re.match(r"^(.*)_(\d{4}-\d{2}-\d{2})$", str(y))
                                return ym.group(1) if ym else str(y)

                            self.oc_registerdata = [
                                new_entry if _oc_name(y) == oc_from_send else y
                                for y in self.oc_registerdata
                            ]
                            if new_entry not in self.oc_registerdata:
                                self.oc_registerdata.append(new_entry)
                            logger.info(f"OC 注册日提前: {old_entry} → {new_entry}")
                        else:
                            all_oc.update(self.oc_registerdata)

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

        # 5. 写入 TSV（别名以内存为准，含下载中新登记的 · 分段别名）
        groups = []
        for main, subs in self.oc_aliases.items():
            if subs:
                groups.append(main + "=" + "=".join(subs))
        aliases_out = ("oc_aliases\t" + ",".join(groups)) if groups else existing_aliases_line

        with open(global_result_path, "w", encoding="utf-8") as f:
            f.write(f"oc_registerdata\t{','.join(sorted(all_oc))}\n")
            self.oc_registerdata = sorted(all_oc)
            if aliases_out:
                f.write(aliases_out.rstrip("\n") + "\n")
            f.write("收/发类型\t发信日期\t发件人\t收件人\t信件类型\t信件下载位置\n")
            for row in final_rows:
                f.write(row + "\n")

        # 6. 写入下载报告
        with open(report_path, "w", encoding="utf-8") as f:
            f.write(f"邮件下载报告（{stage}）\n")
            f.write("=" * 50 + "\n\n")
            f.write(f"下载时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"阶段: {stage}\n")
            f.write(f"总邮件数: {len(email_list or [])}\n")
            f.write(f"成功下载: {len([r for r in results if r.get('status') == 'success'])}\n")
            f.write(f"跳过已有: {len([r for r in results if r.get('status') == 'skipped'])}\n")
            f.write(f"仍在失败/暂挂: {len(failed_details or [])}\n")
            f.write(f"跳过重复: {duplicate_count}\n")
            f.write(f"存储目录: {self.base_dir}\n\n")

            f.write("成功下载的邮件:\n")
            f.write("-" * 50 + "\n")
            for i, result in enumerate([r for r in results if r.get('status') == 'success']):
                f.write(f"{i+1}. {result['subject']}\n")
                f.write(f"   发件邮箱: {result['sender']}\n")
                f.write(f"   日期: {result['date']}\n")
                f.write(f"   保存位置: {result['directory']}\n")
                if result.get('has_attachments'):
                    f.write(f"   附件数量: {result['attachment_count']}\n")
                f.write(f"   原因: {result.get('reason', '')}\n\n")

            f.write("\n跳过的邮件:\n")
            f.write("-" * 50 + "\n")
            for i, result in enumerate([r for r in results if r.get('status') == 'skipped']):
                f.write(f"{i+1}. {result['subject']}\n")
                f.write(f"   原因: {result.get('reason', '')}\n\n")

            if failed_details:
                f.write("\n失败/暂挂的邮件（详细）:\n")
                f.write("-" * 50 + "\n")
                for i, fd in enumerate(failed_details):
                    f.write(f"{i+1}. [{fd['id']}] {fd['subject']}\n")
                    f.write(f"   错误原因: {fd['error']}\n")
                    if fd.get('attempts'):
                        f.write(f"   已尝试: {fd['attempts']} 次\n")
                    f.write("\n")

        logger.info(f"全局结果已生成: {global_result_path}")
        logger.info(f"下载报告已生成: {report_path}（{stage}）")
        # 只对本次成功/跳过的信件做单封附件注册，避免全量重建覆盖人工改过的附件行
        try:
            dirs = []
            for r in results:
                if r.get("status") in ("success", "skipped") and r.get("directory"):
                    dirs.append(r["directory"])
            if dirs:
                self.register_letters_attachments(dirs)
        except Exception as e:
            logger.error(f"入库后单封附件注册失败: {e}")
        return report_path

    def register_letter_attachments(self, directory_or_meta, meta=None):
        """只注册/更新一封信的附件到 attachment_report.tsv（不触碰其他信件）。"""
        from uitl_test import (
            _resolve_email_dir,
            build_attachment_item_from_meta,
            upsert_attachment_report,
            normalize_letter_path_key,
        )
        global_result_path = self.base_dir / "global_result.tsv"
        out_path = self.base_dir / "attachment_report.tsv"
        project_root = script_dir

        if isinstance(directory_or_meta, dict):
            meta = dict(directory_or_meta)
            directory = meta.get("directory") or meta.get("path") or ""
        else:
            directory = str(directory_or_meta or "")
            meta = dict(meta or {})
            meta.setdefault("directory", directory)
            meta.setdefault("path", directory)

        # 若未给全元数据，从 global_result 按路径补齐
        if global_result_path.exists() and (
            not meta.get("sender") or not meta.get("letter_type") or not meta.get("date")
        ):
            key = normalize_letter_path_key(directory)
            for line in global_result_path.read_text(encoding="utf-8").splitlines():
                if line.startswith("oc_") or line.startswith("收/发"):
                    continue
                cols = line.split("\t")
                if len(cols) < 6:
                    continue
                if normalize_letter_path_key(cols[5]) == key:
                    meta.setdefault("send_type", cols[0])
                    meta.setdefault("direction", cols[0])
                    meta.setdefault("date", cols[1])
                    meta.setdefault("sender", cols[2])
                    meta.setdefault("receiver", cols[3])
                    meta.setdefault("letter_type", cols[4])
                    meta["directory"] = cols[5]
                    meta["path"] = cols[5]
                    directory = cols[5]
                    break

        email_dir = _resolve_email_dir(directory, global_result_path if global_result_path.exists() else out_path)
        if email_dir is None:
            # 回退：相对 script_dir
            rel = directory[2:] if str(directory).startswith(".\\") else directory
            cand = script_dir / str(rel).replace("\\", os.sep)
            if cand.is_dir():
                email_dir = cand
        if email_dir is None or not Path(email_dir).is_dir():
            logger.warning(f"单封附件注册失败：找不到信夹 {directory}")
            return None

        # 报告里统一写成相对路径，与 global_result / 库存一致，避免绝对路径键对不上
        try:
            rel_dir = Path(email_dir).resolve().relative_to(script_dir.resolve())
            meta["directory"] = ".\\" + str(rel_dir).replace("/", "\\")
            meta["path"] = meta["directory"]
        except Exception:
            pass

        item = build_attachment_item_from_meta(meta, email_dir, project_root=project_root)
        upsert_attachment_report(out_path, item, remove_empty=True)
        n_t = len(item.get("text_attachments") or [])
        n_f = len(item.get("file_attachments") or [])
        logger.info(f"单封附件已注册: {item.get('directory')}（文字 {n_t} / 文件 {n_f}）")
        return item

    def register_letters_attachments(self, directories):
        """批量单封注册（逐封 upsert，不整表重写无关行以外的语义）。"""
        items = []
        for d in directories or []:
            item = self.register_letter_attachments(d)
            if item is not None:
                items.append(item)
        return items

    def regenerate_attachment_report(self):
        """【仅高级设置修复用】全量重扫并覆盖 attachment_report.tsv。"""
        global_result_path = self.base_dir / "global_result.tsv"
        out_path = self.base_dir / "attachment_report.tsv"
        if not global_result_path.exists():
            logger.warning("无 global_result.tsv，跳过附件报告生成")
            return None
        try:
            # 先纠正误把中转邮箱当成发件人的行，附件柜归属才正确
            self.repair_tsv_sender_names(global_result_path)
            from uitl_test import extract_attachments_from_tsv, save_attachment_report
            results = extract_attachments_from_tsv(global_result_path)
            save_attachment_report(results, out_path)
            n_files = sum(len(r.get('file_attachments') or []) for r in results)
            n_text = sum(len(r.get('text_attachments') or []) for r in results)
            logger.info(
                f"附件报告已全量重建: {out_path}（{len(results)} 封含附件，"
                f"文件 {n_files} / 文字行 {n_text}）"
            )
            return out_path
        except Exception as e:
            logger.error(f"生成附件报告失败: {e}")
            return None

    @staticmethod
    def is_bare_broadcast_a(letter_type):
        """是否为无编号的广播 A（用户主动发出；只记收 A，不记发 A）。"""
        return bool(re.fullmatch(r"A", (letter_type or "").strip(), re.I))

    @staticmethod
    def is_valid_letter_type(letter_type):
        """合法类型：带编号的 A/B/C 及组合、发信含 C 的「C称呼」、以及广播用不带后缀的 A/B/A+B。"""
        p = EmailDownloader.parse_letter_type(letter_type)
        if not p:
            return False
        return bool(p["number"] or p["target"] or "C" not in p["kinds"])

    @staticmethod
    def is_plausible_role_name(name):
        """角色名是否可用：允许 ?/？/数字/符号 ID；排除空、未知、邮箱。"""
        s = (name or "").strip()
        if not s or s == "未知":
            return False
        if "@" in s:
            return False
        return True

    @staticmethod
    def aliases_from_dot_name(name):
        """名字含间隔号 ·/・/• 时，各分段都作为别名（不含全名本身）。

        例：纳西尔·法罗莫尔·盖乌斯·克尼格
        → 纳西尔、法罗莫尔、盖乌斯、克尼格
        """
        s = (name or "").strip()
        if not s:
            return []
        parts = [p.strip() for p in re.split(r"[·・•‧∙⋅]", s) if p.strip()]
        if len(parts) < 2:
            return []
        out = []
        for p in parts:
            if p != s and p not in out:
                out.append(p)
        return out

    def ensure_dot_aliases(self, name):
        """为带 · 的主名补登记分段别名；返回本次新加的别名列表。"""
        name = (name or "").strip()
        if not name:
            return []
        added = []
        cur = list(self.oc_aliases.get(name, []) or [])
        for a in self.aliases_from_dot_name(name):
            if a == name or a in cur:
                continue
            cur.append(a)
            added.append(a)
        if added:
            self.oc_aliases[name] = cur
        return added


class EmailDownloaderGUI:
    MAIL_PRESETS = {
        "QQ邮箱": "imap.qq.com",
        "163邮箱": "imap.163.com",
        "126邮箱": "imap.126.com",
        "Gmail": "imap.gmail.com",
        "Outlook": "outlook.office365.com",
        "其他": "",
    }
    # 下拉首项：遍历账号下全部 IMAP 文件夹（UID 按文件夹区分）
    ALL_FOLDERS_LABEL = "（全部文件夹）"

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
                raw = x.strip()
                if not raw:
                    continue
                # 名_YYYY-MM-DD；名字可含符号/?，日期取末段
                m = re.match(r"^(.*)_(\d{4}-\d{2}-\d{2})$", raw)
                if m:
                    name, d = m.group(1), m.group(2)
                else:
                    name, d = raw, ""
                if name:
                    oc_names.append(name)
                    oc_date_map[name] = d

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

        def _body_for_oc_match(content_text):
            """去掉 content.txt 元数据头，只留正文。

            收件邮箱行会罗列几乎全部 OC 名，若拿全文匹配，C 信永远无法「唯一命中」。
            """
            s = content_text or ""
            m = re.search(r"(?m)^-{5,}\s*$", s)
            if m:
                return s[m.end():].lstrip("\r\n")
            kept = []
            for line in s.splitlines():
                if re.match(r"^(主题|发件邮箱|收件邮箱|日期|发件人|收件人)\s*[:：]", line):
                    continue
                kept.append(line)
            return "\n".join(kept)

        def _is_weak_match_token(token):
            """极弱匹配词：空串或纯符号（如 ? / ？ / …）。单字汉字/字母仍可命中。"""
            t = (token or "").strip()
            if not t:
                return True
            return bool(re.fullmatch(r"[\W_]+", t, flags=re.UNICODE))

        def _count_oc_in_text(content_text):
            """主名与别名都计入该角色；较长的名字优先匹配，减少短名误伤。"""
            oc_count = {oc: 0 for oc in oc_names}
            remaining = _body_for_oc_match(content_text)
            # 按名字长度降序替换计数，避免「甲」吃掉「甲二」之类
            tokens = []
            for oc in oc_names:
                tokens.append((oc, oc))
                for alias in aliases_by_main.get(oc, []):
                    tokens.append((oc, alias))
            tokens.sort(key=lambda t: len(t[1]), reverse=True)
            for main, token in tokens:
                if not token or _is_weak_match_token(token):
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

            # 清理历史误入库的发广播 A（类型栏为 A，或路径叶名为裸 A）
            if send_type == "发":
                leaf = re.sub(
                    r"^\d{8}_\d{6}_",
                    "",
                    str(directory).replace("\\", "/").split("/")[-1],
                )
                if EmailDownloader.is_bare_broadcast_a(subject) or re.fullmatch(r"A", leaf or "", re.I):
                    continue

            if send_type == "收":
                if receiver:
                    already_had += 1
                else:
                    empty_inbound += 1
                    subj = (subject or "").strip()
                    if only_one_oc:
                        cols[3] = only_oc_name
                        changed += 1
                        oc_guess_count[only_oc_name] += 1
                        context_logs.append(f"[单OC] 行{data_start + idx + 2} → {only_oc_name}")
                    elif re.match(r"(?i)^(A\+C|C)", subj):
                        # C / A+C：正文唯一命中主名或别名
                        content_path = None
                        try:
                            from uitl_test import _resolve_email_dir
                            ed = _resolve_email_dir(directory, global_result_path)
                            if ed:
                                content_path = Path(ed) / "content.txt"
                        except Exception:
                            content_path = None
                        if content_path is None:
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
                                        f"[C] 行{data_start + idx + 2} → {chosen}（正文唯一命中）"
                                    )
                                else:
                                    unguessed_count += 1
                                    hit = ", ".join(
                                        f"{oc}×{oc_count[oc]}" for oc in nonzero
                                    ) or "无"
                                    context_logs.append(
                                        f"[C] 行{data_start + idx + 2} 未补全（正文命中：{hit}）"
                                    )
                            except Exception:
                                unguessed_count += 1
                        else:
                            unguessed_count += 1
                    elif re.match(r"(?i)^(A\+B|B\+C|A|B)", subj):
                        # A / B / A+B / B+C（收到的 B+C 按 B 广播收）：注册日 ≤ 发信日的全部 OC
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
        self.folder_var = tk.StringVar(value=self.ALL_FOLDERS_LABEL)
        self.start_date_var = tk.StringVar()
        self.end_date_var = tk.StringVar()
        self.work_status_var = tk.StringVar(
            value="可在下方查看下载相关信息。"
        )
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
            if cls in ("Text", "Listbox", "Treeview", "Combobox", "TCombobox"):
                return
            # Combobox 下拉列表在独立 Toplevel 上；bind_all 仍会收到事件，勿带动整页
            try:
                if w.winfo_toplevel() is not holder.winfo_toplevel():
                    return
            except Exception:
                return
            try:
                cur = w
                while cur is not None:
                    try:
                        c = cur.winfo_class()
                    except Exception:
                        break
                    if c in ("Text", "Listbox", "Treeview", "Combobox", "TCombobox"):
                        return
                    cur = getattr(cur, "master", None)
            except Exception:
                pass
            if canvas.winfo_exists():
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
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
        self.login_btn = ttk.Button(footer, text="登录", command=self.do_login)
        self.login_btn.pack(side=tk.RIGHT)

    def _on_mailbox_type(self, _event=None):
        name = self.mailbox_type_var.get()
        host = self.MAIL_PRESETS.get(name, "")
        if host:
            self.server_var.set(host)
            self.server_row.grid_remove()
        else:
            self.server_row.grid(row=3, column=0, columnspan=2, sticky=tk.EW, pady=6)

    def do_login(self):
        if getattr(self, "busy", False):
            return
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

        # IMAP 登录放后台，避免 UI 线程卡死（弱网可卡到超时，默认可达数分钟）
        self.busy = True
        try:
            self.login_btn.configure(state=tk.DISABLED)
        except Exception:
            pass
        self.status_var.set("正在登录…（界面可拖动；完成后会自动跳转）")

        def work():
            ok = False
            folders = []
            err = ""
            try:
                ok = bool(self.downloader.connect_to_email())
                if ok:
                    folders = self.downloader.get_all_mailboxes() or []
            except Exception as e:
                err = str(e)
                ok = False

            def done():
                self.busy = False
                try:
                    self.login_btn.configure(state=tk.NORMAL)
                except Exception:
                    pass
                if not ok:
                    messagebox.showerror(
                        "错误",
                        "登录失败，请检查邮箱、授权码和邮箱类型。"
                        + (f"\n{err}" if err else ""),
                    )
                    self.status_var.set("登录失败")
                    return
                if folders:
                    self._set_folder_choices(folders)
                else:
                    messagebox.showerror("错误", "登录成功，但没有找到收件箱。")
                    self.status_var.set("登录成功但无文件夹")
                    return
                self.status_var.set(f"已登录，默认文件夹 {self.selected_folder}")
                self.notebook.select(self.import_frame)

            self.root.after(0, done)

        threading.Thread(target=work, daemon=True).start()

    def _set_folder_choices(self, folders, prefer=None):
        """文件夹下拉：首项为「全部文件夹」（默认），其余为真实邮箱夹。"""
        real = [f for f in (folders or []) if f and f != self.ALL_FOLDERS_LABEL]
        values = [self.ALL_FOLDERS_LABEL] + real
        try:
            self.folder_combo["values"] = values
        except Exception:
            pass
        cur = (self.folder_var.get() or "").strip()
        if cur == self.ALL_FOLDERS_LABEL or cur in real:
            chosen = cur
        elif prefer and prefer in real:
            chosen = prefer
        else:
            chosen = self.ALL_FOLDERS_LABEL
        self.folder_var.set(chosen)
        self.selected_folder = chosen

    # ==================== 检索入库（子页：OC 管理 / 信件下载）====================
    def _build_import_tab(self):
        self.import_notebook = ttk.Notebook(self.import_frame)
        self.import_notebook.pack(fill=tk.BOTH, expand=True)

        self.oc_page = ttk.Frame(self.import_notebook, padding="4")
        self.download_page = ttk.Frame(self.import_notebook, padding="4")
        # 推荐：先下载 → 自动拆分补全 → 再到 OC 管理补别名/再补全
        self.import_notebook.add(self.download_page, text="1. 信件下载")
        self.import_notebook.add(self.oc_page, text="2. OC 管理")

        self._build_download_page()
        self._build_oc_page()
        self.import_notebook.bind("<<NotebookTabChanged>>", self._on_import_tab_changed)
        try:
            self.import_notebook.select(self.download_page)
        except Exception:
            pass

    def _on_import_tab_changed(self, _event=None):
        """切到 OC 管理时从 TSV 重载，避免下载新建的 OC 要退出重进才看得见。"""
        try:
            if self.import_notebook.index(self.import_notebook.select()) == self.import_notebook.index(self.oc_page):
                self.load_oc_from_tsv()
        except Exception:
            pass

    def _build_oc_page(self):
        inner = self._scrollable(self.oc_page)

        oc_box = ttk.LabelFrame(inner, text="角色（OC）", padding=(8, 6))
        oc_box.pack(fill=tk.BOTH, expand=True, pady=4)
        ttk.Label(
            oc_box,
            text="信件入库后发信人会自动登记；此处用来修改与核对 OC 信息。\n"
                 "名字含「·」时会自动把各分段记为别名。更新别名后请自行使用下方的「补全收件人」。\n"
                 "别名用于猜测收件人与可视化。例如：名为「小A」的 OC 常被称为 A、A君，"
                 "可把这些记为别名；正文唯一出现相关称呼时，会猜测此信寄给该 OC。",
            wraplength=520,
        ).pack(anchor=tk.W, pady=(0, 3))

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

        guess_row = ttk.Frame(oc_box)
        guess_row.pack(fill=tk.X, pady=(8, 2))
        ttk.Button(guess_row, text="补全收件人", command=self.run_smart_guess_receivers).pack(side=tk.LEFT)
        ttk.Label(
            guess_row,
            text="按已登记 OC / 别名，填充库存里仍空的收件人（不覆盖已填写的）",
            foreground="gray",
        ).pack(side=tk.LEFT, padx=8)

        foot = ttk.Frame(self.oc_page)
        foot.pack(fill=tk.X, pady=(8, 0))
        ttk.Button(foot, text="返回：信件下载", command=self._goto_download_page).pack(side=tk.LEFT)
        ttk.Button(foot, text="打开[地缝可视化]", command=self.open_visualization_html).pack(side=tk.RIGHT)

    def _build_download_page(self):
        inner = self._scrollable(self.download_page)

        pull = ttk.LabelFrame(inner, text="拉取来往信件", padding=(8, 6))
        pull.pack(fill=tk.X, pady=6)
        ttk.Label(
            pull,
            text="1. 可选「（全部文件夹）」扫遍所有文件夹并下载来往信件。\n"
                 "2. 程序会自行分辨来往信件，并从中拆分出各个发件、收件人。\n"
                 "3. 由于规则化识别无法保障每一种情况，请在完成信件下载后自行核对与补全收件人。\n",
            wraplength=520,
        ).pack(anchor=tk.W)
        fr = ttk.Frame(pull)
        fr.pack(fill=tk.X, pady=4)
        ttk.Label(fr, text="邮箱文件夹").pack(side=tk.LEFT)
        self.folder_combo = ttk.Combobox(fr, textvariable=self.folder_var, width=36, state="readonly")
        self.folder_combo.pack(side=tk.LEFT, padx=6)
        ttk.Label(fr, text="「全部文件夹」速度可能较慢", foreground="gray").pack(side=tk.LEFT, padx=4)
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
        self.match_tree.bind("<Double-1>", self._on_match_tree_double_click)
        ttk.Label(res_box, text="双击信件可预览正文", foreground="gray").pack(anchor=tk.W)

        act = ttk.Frame(res_box)
        act.pack(fill=tk.X, pady=6)
        ttk.Button(act, text="处理主题不合规的信…", command=self.open_unmatched_dialog).pack(side=tk.LEFT)
        ttk.Button(act, text="下载入库", command=self.start_download_pipeline).pack(side=tk.RIGHT)

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
        self.cancel_download_btn = ttk.Button(
            foot, text="以当前进度结束", command=self.cancel_download_pipeline, state=tk.DISABLED
        )
        self.cancel_download_btn.pack(side=tk.LEFT, padx=6)
        ttk.Button(foot, text="下一步：OC 管理", command=self._goto_oc_page).pack(side=tk.RIGHT, padx=(6, 0))
        ttk.Button(foot, text="打开[地缝可视化]", command=self.open_visualization_html).pack(side=tk.RIGHT)

    def _goto_download_page(self):
        self.import_notebook.select(self.download_page)

    def _goto_oc_page(self):
        self.load_oc_from_tsv()
        self.refresh_oc_tree()
        self.import_notebook.select(self.oc_page)

    # ----- 人名补全（OC + 库存已有收发件人）-----
    def _collect_known_person_names(self):
        names = set()
        for entry in self.downloader.oc_registerdata or []:
            n = (entry.split("_")[0] if "_" in entry else entry).strip()
            if n:
                names.add(n)
        for main, aliases in (self.downloader.oc_aliases or {}).items():
            if main:
                names.add(main)
            for a in aliases or []:
                if a:
                    names.add(a)
        tsv = self.downloader.base_dir / "global_result.tsv"
        if tsv.exists():
            try:
                for line in tsv.read_text(encoding="utf-8").splitlines():
                    if not (line.startswith("收\t") or line.startswith("发\t")):
                        continue
                    cols = line.split("\t")
                    if len(cols) < 4:
                        continue
                    for raw in (cols[2], cols[3]):
                        for part in str(raw or "").split(","):
                            p = part.strip()
                            if p and p not in ("未知", "[未知的收信人]") and "@" not in p:
                                names.add(p)
            except Exception:
                pass
        return sorted(names, key=lambda s: (s.casefold(), s))

    def _bind_name_autocomplete(self, combo):
        """可输入 + 下拉；Tab 补全当前前缀的第一条匹配。"""
        def on_tab(_event=None):
            cur = (combo.get() or "").strip()
            vals = list(combo.cget("values") or ())
            if not vals:
                return "break"
            if not cur:
                combo.set(vals[0])
                combo.icursor(tk.END)
                return "break"
            matches = [v for v in vals if v.startswith(cur)]
            if not matches:
                cf = cur.casefold()
                matches = [v for v in vals if v.casefold().startswith(cf)]
            if matches:
                combo.set(matches[0])
                combo.icursor(tk.END)
            return "break"

        def on_focus(_event=None):
            try:
                combo.configure(values=self._collect_known_person_names())
            except Exception:
                pass

        combo.bind("<Tab>", on_tab)
        combo.bind("<FocusIn>", on_focus)
        return combo

    LETTER_TYPE_HINTS = {
        "收": "收信须带编号：A123 / A*12 / B45 / C678 / A+B12 / A+C3 / B+C7",
        "发": "发信：B、A+B 不带后缀；C某某 / A+C某某 / B+C某某 须写对象；单独 A 不录入",
        "": "收：A123 / B45 / C678 / A+B12　发：B / A+B / C某某 / A+C某某 / B+C某某",
    }

    def _bind_type_feedback(self, label, type_var, get_direction, watch_vars=()):
        """类型输入框旁的实时提示：空时按收/发给示例，输入后显示规范结果或错误。"""
        def refresh(*_):
            d = (get_direction() or "").strip()
            raw = (type_var.get() or "").strip()
            if not raw:
                label.configure(text=self.LETTER_TYPE_HINTS.get(d, self.LETTER_TYPE_HINTS[""]), foreground="gray")
                return
            lt, err, implied = EmailDownloader.check_letter_type(raw, d)
            if err:
                label.configure(text="⚠ " + err, foreground="#c0392b")
                return
            note = f"✓ 将记为 {lt}"
            if implied:
                note += f"（收件人：{implied}）"
            elif d == "发":
                note += "（广播，可不填收件人）"
            label.configure(text=note, foreground="#1e8449")

        for v in (type_var, *watch_vars):
            v.trace_add("write", refresh)
        refresh()
        return refresh

    def _confirm_missing_receiver(self, parent=None):
        return messagebox.askyesno(
            "收件人未填",
            "收件人为空。\n入库后可由补全收件人推测（A/B 按当时已登记的 OC，C 按正文「致…」），"
            "推测不到的会显示为 [未知的收信人]。\n\n仍要继续吗？",
            parent=parent or self.root,
        )

    def _checked_letter_type(self, type_var, direction, parent=None):
        """校验并回填规范后的类型；不合法时弹窗并返回 None。返回 (类型, 推出的收件人)。"""
        lt, err, implied = EmailDownloader.check_letter_type(type_var.get(), direction)
        if err:
            messagebox.showwarning("信件类型", err, parent=parent or self.root)
            return None
        type_var.set(lt)
        return lt, implied

    def _make_name_combo(self, parent, textvariable, width=18):
        cb = ttk.Combobox(
            parent,
            textvariable=textvariable,
            width=width,
            values=self._collect_known_person_names(),
        )
        return self._bind_name_autocomplete(cb)

    def _refresh_name_combos(self, *combos):
        names = self._collect_known_person_names()
        for cb in combos:
            if cb is None:
                continue
            try:
                cb.configure(values=names)
            except Exception:
                pass

    def _count_empty_receivers(self):
        path = self.downloader.base_dir / "global_result.tsv"
        if not path.exists():
            return 0
        n = 0
        try:
            for line in path.read_text(encoding="utf-8").splitlines():
                if not (line.startswith("收\t") or line.startswith("发\t")):
                    continue
                cols = line.split("\t")
                if len(cols) >= 4 and not (cols[3] or "").strip():
                    n += 1
        except Exception:
            return 0
        return n

    def _after_ingest_open_or_prompt(self):
        """入库后：收件人已齐则打开可视化；否则引导去 OC 管理补别名并再补全。"""
        empty_n = self._count_empty_receivers()
        if empty_n <= 0:
            self.open_visualization_html()
            return
        go_oc = messagebox.askyesno(
            "需补充收件人",
            f"入库完成，仍有 {empty_n} 封收件人为空。\n\n"
            f"建议到「OC 管理」核对角色、补别名，后点击「补全收件人」。\n"
            f"仍空的可到「信件库存」手工改。\n\n"
            f"是否现在打开 OC 管理？",
        )
        if go_oc:
            try:
                self.notebook.select(self.import_frame)
                self._goto_oc_page()
            except Exception:
                pass

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
        # 已有带 · 的主名补上分段别名（仅内存；有新增时落盘）
        alias_added = 0
        for entry in self.downloader.oc_registerdata:
            name = entry.split("_")[0] if "_" in entry else entry
            alias_added += len(self.downloader.ensure_dot_aliases(name))
        if alias_added:
            self._persist_oc_header()
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
        main = self.oc_name_var.get().strip()
        if main and name == main:
            messagebox.showinfo("提示", "别名不能与主名相同。")
            return
        self.alias_listbox.insert(tk.END, name)
        self.alias_input_var.set("")
        # 若别名本身是已登记 OC：合并过来（取更早注册日），从主列表移除
        if main:
            self._absorb_oc_as_alias(main, name)

    def remove_alias_item(self):
        sel = list(self.alias_listbox.curselection())
        for i in reversed(sel):
            self.alias_listbox.delete(i)

    def _absorb_oc_as_alias(self, main_name, alias_name):
        """把已登记的 alias_name 并入 main_name：别名登记 + 主名日期取更早者。"""
        alias_date = ""
        kept = []
        for entry in self.downloader.oc_registerdata:
            n = entry.split("_")[0]
            d = entry.split("_", 1)[1] if "_" in entry else ""
            if n == alias_name:
                alias_date = d
                # 别名自己的别名也挂到主名下
                for sub in self.downloader.oc_aliases.pop(alias_name, []) or []:
                    if sub not in self._alias_list_get() and sub != main_name:
                        self.alias_listbox.insert(tk.END, sub)
                continue
            kept.append(entry)
        if alias_date:
            self.downloader.oc_registerdata = kept
            cur = self.oc_date_var.get().strip()
            if not cur or (alias_date and alias_date < cur):
                self.oc_date_var.set(alias_date)

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
        # 名字含 · 时自动把各分段加入别名列表
        for a in EmailDownloader.aliases_from_dot_name(name):
            if a not in self._alias_list_get() and a != name:
                self.alias_listbox.insert(tk.END, a)
                self._absorb_oc_as_alias(name, a)
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
        tip = f"已保存角色 {name}"
        if aliases:
            tip += f"（别名 {len(aliases)} 个）"
        self.status_var.set(tip)

    def run_smart_guess_receivers(self):
        """OC 管理页：手动发起收件人补全收件人。"""
        if self.busy:
            messagebox.showinfo("请稍候", "正在检索或下载，结束后再补全。")
            return
        # 表单里已编辑但未点「添加/更新」的别名也先并进内存再落盘，避免白改
        name = self.oc_name_var.get().strip()
        if name:
            pending = self.alias_input_var.get().strip()
            if pending and pending not in self._alias_list_get() and pending != name:
                self.alias_listbox.insert(tk.END, pending)
                self.alias_input_var.set("")
            for a in EmailDownloader.aliases_from_dot_name(name):
                if a not in self._alias_list_get() and a != name:
                    self.alias_listbox.insert(tk.END, a)
            aliases = self._alias_list_get()
            if aliases:
                self.downloader.oc_aliases[name] = aliases
            # 若该名尚未在登记表里，顺带登记，否则补全认不到
            if not any((e.split("_")[0] == name) for e in self.downloader.oc_registerdata):
                date_str = self.oc_date_var.get().strip() or datetime.now().strftime("%Y-%m-%d")
                self.downloader.oc_registerdata.append(f"{name}_{date_str}")
                self.refresh_oc_tree()
        self._persist_oc_header()
        msg = self.smart_guess_receivers(notify=True)
        self.refresh_inventory()
        self._refresh_name_combos(
            getattr(self, "inv_sender_combo", None),
            getattr(self, "inv_receiver_combo", None),
        )
        if msg:
            self.status_var.set(msg.split("\n")[0][:80])

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

    def _resolve_fetch_folders(self):
        """返回本次要扫的真实文件夹列表（不含「全部文件夹」哨兵）。"""
        sel = (self.folder_var.get() or self.selected_folder or "").strip()
        all_real = [
            f for f in (self.downloader.get_all_mailboxes() or [])
            if f and f != self.ALL_FOLDERS_LABEL
        ]
        self._set_folder_choices(all_real, prefer=sel if sel != self.ALL_FOLDERS_LABEL else None)
        sel = (self.folder_var.get() or "").strip()
        if sel == self.ALL_FOLDERS_LABEL:
            return list(all_real), True
        if sel and sel in all_real:
            return [sel], False
        if all_real:
            return [all_real[0]], False
        return [], False

    def start_fetch(self):
        if self.busy:
            return
        if not self.downloader.client and not self.downloader.connect_to_email():
            messagebox.showerror("错误", "请先在「登录」页登录。")
            self.notebook.select(self.login_frame)
            return
        folders, scan_all = self._resolve_fetch_folders()
        if not folders:
            messagebox.showerror("错误", "没有找到邮箱文件夹。")
            return
        self.selected_folder = self.folder_var.get() or self.ALL_FOLDERS_LABEL
        # 允许无 OC 直接拉取：发信下载后会自动登记，再去 OC 管理打磨
        self.busy = True
        self.work_progress["value"] = 0
        self.summary_var.set("")
        self.clear_work_log()
        for item in self.match_tree.get_children():
            self.match_tree.delete(item)
        tip = f"全部 {len(folders)} 个文件夹" if scan_all else folders[0]
        self.work_status_var.set(f"正在检索往来信件（{tip}）…")
        self.status_var.set("正在检索往来信件…")
        self.matched_results = []
        self.unmatched_results = []
        self.manual_queue = []
        self._fetch_folder_list = folders
        self._fetch_scan_all = scan_all
        threading.Thread(target=self._run_fetch, daemon=True).start()

    def _run_fetch(self):
        try:
            folders = list(getattr(self, "_fetch_folder_list", None) or [])
            scan_all = bool(getattr(self, "_fetch_scan_all", False))
            if not folders:
                folders = [self.folder_var.get() or self.selected_folder]
                scan_all = False
            start_date = self._parse_optional_date(self.start_date_var)
            end_date = self._parse_optional_date(self.end_date_var)
            merged = []
            # IMAP UID 仅在同一文件夹内唯一 → 用 (folder, id) 去重
            seen = set()
            unmatched_all = []
            unmatched_seen = set()
            n_folders = max(len(folders), 1)
            rules = ((1, "收到的信"), (2, "发出的信"))
            total_steps = n_folders * len(rules)
            step_i = 0

            for fi, folder in enumerate(folders):
                for index, (rule_num, label) in enumerate(rules):
                    rule = self.downloader.get_default_rule(rule_num)
                    step_i += 1
                    step_base = (step_i - 1) / total_steps * 100
                    step_span = 100 / total_steps
                    fol_tip = f"{fi + 1}/{n_folders} {folder}" if scan_all else folder

                    def cb(progress, found, _base=step_base, _span=step_span, _label=label, _fol=fol_tip):
                        def _ui(p=progress, f=found, b=_base, sp=_span, lab=_label, fol=_fol):
                            if f > 0:
                                self.work_status_var.set(
                                    f"正在检索（{fol} / {lab}）… 本夹已匹配约 {f} 封"
                                )
                            else:
                                self.work_status_var.set(f"正在检索（{fol} / {lab}）…")
                            self.status_var.set("正在检索往来信件…")
                            self.work_progress["value"] = b + (p / 100.0) * sp
                        self.root.after(0, _ui)

                    self.root.after(
                        0,
                        lambda fol=fol_tip, lab=label: self.work_status_var.set(
                            f"正在检索（{fol} / {lab}）…"
                        ),
                    )

                    self.downloader.last_unmatched = []
                    try:
                        results = self.downloader.search_emails_advanced(
                            folder=folder,
                            subject_pattern=rule.get("subject_pattern", ""),
                            from_address=rule.get("from", ""),
                            to_address=rule.get("to", ""),
                            start_date=start_date,
                            end_date=end_date,
                            callback=cb,
                            max_emails=1500,
                        ) or []
                    except Exception as e:
                        logger.error(f"检索文件夹失败 {folder}: {e}")
                        results = []

                    for email_info in results:
                        eid = email_info.get("id")
                        fol = email_info.get("folder") or folder
                        # 同一封信可能被复制到多个文件夹（163 还会沿用同一 UID）→ 优先按 Message-ID 去重
                        key = email_info.get("message_id") or (fol, eid)
                        if key not in seen:
                            seen.add(key)
                            if rule_num == 1:
                                direction = "收"
                            elif rule_num == 2:
                                direction = "发"
                            else:
                                direction = self.downloader.direction_from_addresses(
                                    email_info.get("sender", ""),
                                    email_info.get("recipients") or [],
                                )
                            email_info["direction"] = direction
                            email_info["folder"] = fol
                            merged.append(email_info)
                    for item in getattr(self.downloader, "last_unmatched", []) or []:
                        eid = item.get("id")
                        fol = item.get("folder") or folder
                        key = item.get("message_id") or (fol, eid)
                        if key not in unmatched_seen and key not in seen:
                            unmatched_seen.add(key)
                            unmatched_all.append(item)

                    snap = list(merged)
                    last = step_i >= total_steps

                    def _mid_refresh(s=snap, lab=label, n=len(snap), fol=fol_tip, is_last=last):
                        self.matched_results = s
                        self._fill_match_tree()
                        if is_last:
                            self.work_status_var.set("正在检索往来信件… 汇总中")
                        else:
                            self.work_status_var.set(
                                f"正在检索…（{fol}/{lab}，已纳入 {n} 封）"
                            )

                    self.root.after(0, _mid_refresh)

            self.matched_results = merged
            self.unmatched_results = unmatched_all
            self.downloader.search_results = list(merged)

            def done():
                self._fill_match_tree()
                self.work_progress["value"] = 100
                scope = f"扫了 {len(folders)} 个文件夹，" if scan_all else ""
                self.work_status_var.set(
                    f"{scope}找到 {len(merged)} 封合规来往信；另有 {len(unmatched_all)} 封主题不合规。"
                )
                self.status_var.set(
                    f"检索完成：{len(merged)} 封合规 / {len(unmatched_all)} 封待处理"
                )
                self.add_work_log(
                    f"拉取完成：{'全部文件夹×'+str(len(folders))+'；' if scan_all else ''}"
                    f"匹配 {len(merged)}，未匹配 {len(unmatched_all)}\n"
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

    def _on_match_tree_double_click(self, event):
        iid = self.match_tree.identify_row(event.y)
        if not iid:
            return
        items = self.matched_results + self.manual_queue
        idx = self.match_tree.index(iid)
        if idx < len(items):
            self._open_letter_preview(items[idx])

    def _open_letter_preview(self, item, parent=None):
        """弹窗预览正文（只取文本段，不下载附件）。"""
        if self.busy:
            messagebox.showinfo("请稍候", "正在检索、下载或拉取另一封预览，稍后再试。", parent=parent or self.root)
            return
        folder = item.get("folder") or self.selected_folder
        eid = item.get("id")
        if eid is None or not folder or folder == self.ALL_FOLDERS_LABEL:
            messagebox.showwarning("提示", "这封信缺少文件夹信息，无法预览。", parent=parent or self.root)
            return
        win = tk.Toplevel(parent or self.root)
        win.title("正文预览：" + (item.get("subject") or "")[:40])
        win.geometry("680x560")
        win.transient(parent or self.root)
        head = ttk.Frame(win, padding=(10, 8))
        head.pack(fill=tk.X)
        date_str = item["date_obj"].strftime("%Y-%m-%d %H:%M") if item.get("date_obj") else (item.get("date") or "")[:19]
        meta = (
            f"方向：{item.get('direction') or item.get('user_direction') or '?'}　日期：{date_str}　文件夹：{folder}\n"
            f"发件邮箱：{item.get('sender', '')}\n主题：{item.get('subject', '')}"
        )
        ttk.Label(head, text=meta, justify=tk.LEFT, wraplength=640).pack(anchor=tk.W)
        parsed_var = tk.StringVar(value="正在拉取正文…")
        ttk.Label(head, textvariable=parsed_var, justify=tk.LEFT, wraplength=640, foreground="#1f5f8b").pack(
            anchor=tk.W, pady=(4, 0)
        )
        box = scrolledtext.ScrolledText(win, wrap=tk.WORD)
        box.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        box.configure(state="disabled")

        cache = self.__dict__.setdefault("_preview_cache", {})
        key = (folder, eid)

        def show(result):
            if not win.winfo_exists():
                return
            text, attachments, err = result
            if err:
                parsed_var.set("⚠ " + err)
                return
            cache[key] = result
            d = self.downloader
            direction = item.get("direction") or item.get("user_direction") or ""
            letter_type = d.normalize_letter_type(d.extract_letter_type(re.sub(r"\s+", "_", item.get("subject", ""))))
            sender_name = d.extract_letter_sender_name(
                content_text=text, directory="", mail_sender=item.get("sender", ""), subject=item.get("subject", "")
            )
            receiver_name = d.extract_letter_receiver_name(text, letter_type, direction)
            att = "、".join(attachments) if attachments else "无"
            parsed_var.set(
                f"解析：类型 {letter_type or '?'}　发件人 {sender_name or '?'}　收件人 {receiver_name or '?'}\n"
                f"文件附件（未下载）：{att}"
            )
            box.configure(state="normal")
            box.delete("1.0", tk.END)
            box.insert(tk.END, text or "（没有正文文本）")
            box.configure(state="disabled")

        if key in cache:
            show(cache[key])
            return

        def work():
            result = self.downloader.fetch_email_preview(eid, folder)

            def finish():
                self.busy = False
                show(result)
            self.root.after(0, finish)

        # 与检索/下载共用一条 IMAP 连接，预览期间占用 busy
        self.busy = True
        threading.Thread(target=work, daemon=True).start()

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
        type_tip = ttk.Label(form, wraplength=420, justify=tk.LEFT)
        type_tip.grid(row=1, column=2, sticky=tk.W, padx=6)
        self._bind_type_feedback(type_tip, type_var, dir_var.get, watch_vars=(dir_var,))
        ttk.Label(form, text="发件人").grid(row=2, column=0, sticky=tk.W, pady=3)
        sender_cb = self._make_name_combo(form, sender_var, width=24)
        sender_cb.grid(row=2, column=1, sticky=tk.EW, pady=3)
        ttk.Label(form, text="收件人").grid(row=3, column=0, sticky=tk.W, pady=3)
        receiver_cb = self._make_name_combo(form, receiver_var, width=24)
        receiver_cb.grid(row=3, column=1, sticky=tk.EW, pady=3)
        ttk.Label(form, text="可输入 / 下拉 / Tab 补全", foreground="gray").grid(
            row=3, column=2, sticky=tk.W, padx=6
        )

        def on_sel(_e=None):
            sel = tree.selection()
            if not sel:
                return
            item = self.unmatched_results[id_map[sel[0]]]
            dir_var.set(item.get("direction", ""))
            self._refresh_name_combos(sender_cb, receiver_cb)
            # 收：收件人常为自己的 OC；发：发件人常为自己的 OC
            if item.get("direction") == "收" and self.downloader.oc_registerdata:
                receiver_var.set(self.downloader.oc_registerdata[0].split("_")[0])
            elif item.get("direction") == "发" and self.downloader.oc_registerdata:
                sender_var.set(self.downloader.oc_registerdata[0].split("_")[0])

        tree.bind("<<TreeviewSelect>>", on_sel)

        def on_double(event):
            iid = tree.identify_row(event.y)
            if iid and iid in id_map:
                self._open_letter_preview(self.unmatched_results[id_map[iid]], parent=win)

        tree.bind("<Double-1>", on_double)

        def add_selected():
            sel = tree.selection()
            if not sel:
                messagebox.showwarning("提示", "请先选一封信。", parent=win)
                return
            idx = id_map[sel[0]]
            direction = self.unmatched_results[idx].get("direction") or dir_var.get()
            checked = self._checked_letter_type(type_var, direction, parent=win)
            if not checked:
                return
            letter_type, implied_receiver = checked
            sname = sender_var.get().strip()
            rname = receiver_var.get().strip() or implied_receiver
            receiver_var.set(rname)
            if not sname:
                messagebox.showwarning("提示", "请填写发件人。", parent=win)
                return
            # 发 B / A+B 是广播，无固定收件人；其余空收件人只提醒
            if not rname and not (direction == "发" and not implied_receiver):
                if not self._confirm_missing_receiver(win):
                    return
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
        self._mid_flushed = False
        self.work_progress["value"] = 0
        self.work_status_var.set(f"准备下载 {len(queue)} 封…")
        self.clear_work_log()
        try:
            self.cancel_download_btn.configure(state=tk.NORMAL)
        except Exception:
            pass
        threading.Thread(target=self._run_download_pipeline, args=(queue,), daemon=True).start()

    def cancel_download_pipeline(self):
        if not self.busy:
            return
        self.downloader.request_cancel()
        self.work_status_var.set("正在以当前进度结束…（只保留已完整的信件）")
        self.add_work_log("\n—— 用户选择：以当前进度结束 ——\n")

    def _status_label(self, status, attempt=None):
        if status == "skipped":
            return "跳过"
        if status == "failed":
            return "失败"
        if status == "success":
            return "已下载"
        if status == "deferred":
            return "暂挂待后台"
        if status == "retry":
            n = attempt or ""
            return f"失败重试{n}" if n != "" else "失败重试"
        if status == "processing":
            return "处理中"
        return "处理中"

    def upsert_work_log_mail(self, mail_key, shown, total, status, subject, attempt=None):
        """按邮件 id 就地翻新详细记录行（修复重试成功后仍显示处理中）。"""
        label = self._status_label(status, attempt)
        msg = f"[{shown}/{total}] {label}: {subject}\n"
        tag = f"mail_{mail_key}"

        def _ui():
            self.work_log.configure(state="normal")
            ranges = self.work_log.tag_ranges(tag)
            if ranges:
                self.work_log.delete(ranges[0], ranges[1])
                self.work_log.insert(ranges[0], msg, tag)
            else:
                self.work_log.insert(tk.END, msg, tag)
            self.work_log.see(tk.END)
            self.work_log.configure(state="disabled")

        self.root.after(0, _ui)

    def _flush_download_artifacts(self, results, queue, failed_details, stage, open_viz=False):
        """中期/最终：写 TSV+报告、智能猜测、刷新库存。"""
        self._persist_oc_header()
        self.downloader.generate_download_report(
            results, queue, failed_details, stage=stage
        )
        guess = self.smart_guess_receivers(notify=False) or ""
        self.refresh_inventory()
        self.refresh_oc_tree()
        self._refresh_name_combos(
            getattr(self, "inv_sender_combo", None),
            getattr(self, "inv_receiver_combo", None),
        )
        if open_viz:
            self._after_ingest_open_or_prompt()
        return guess

    def _run_download_pipeline(self, queue):
        mid_box = {"guess": "", "deferred_n": 0, "cancelled": False}
        try:
            def progress(p, current, total, email_info=None):
                def _ui(pv=p, c=current, t=total):
                    # 主轮占到 70%，后台补下 70→95
                    self.work_progress["value"] = min(95, max(pv * 0.7, self.work_progress["value"]))
                    self.work_status_var.set(f"下载中… {c}/{t}")
                self.root.after(0, _ui)
                info = email_info
                if not info:
                    return
                try:
                    shown = queue.index(info) + 1
                except ValueError:
                    shown = current
                self.upsert_work_log_mail(
                    shown,
                    shown,
                    total,
                    info.get("status", ""),
                    info.get("subject", ""),
                    attempt=info.get("attempt"),
                )

            def mid_phase(results, deferred, failed_details, cancelled):
                mid_box["deferred_n"] = len(deferred)
                mid_box["cancelled"] = cancelled
                try:
                    # 文件 IO 可在工作线程；Tk/智能猜测放到 UI 线程
                    self.downloader.generate_download_report(
                        results, queue, failed_details, stage="中期汇总（主轮结束）"
                    )
                    self._mid_flushed = True
                except Exception as e:
                    logger.error(f"中期落盘失败: {e}")

                ok_n = len([r for r in results if r.get("status") in ("success", "skipped")])

                def _ui_mid(n_ok=ok_n, n_def=len(deferred), was_cancel=cancelled):
                    try:
                        self._persist_oc_header()
                        guess = self.smart_guess_receivers(notify=False) or ""
                        mid_box["guess"] = guess
                    except Exception as e:
                        logger.error(f"中期智能猜测失败: {e}")
                        guess = ""
                    self.work_progress["value"] = 75
                    self.refresh_inventory()
                    self.refresh_oc_tree()
                    if was_cancel:
                        self.summary_var.set(
                            f"已按当前进度结束。已登记合法信件 {n_ok}/{len(queue)} 封。\n\n{guess}"
                        )
                        self.work_status_var.set("已提前结束，合法信件已入库")
                        self.add_work_log("\n—— 中期落盘完成（提前结束）——\n")
                        self._after_ingest_open_or_prompt()
                    elif n_def:
                        self.summary_var.set(
                            f"大部分任务已成功（已登记 {n_ok}/{len(queue)}）。\n"
                            f"另有 {n_def} 封因网络延迟暂挂，正在后台继续补下…\n\n{guess}"
                        )
                        self.work_status_var.set(f"主轮完成，后台补下 {n_def} 封…")
                        self.add_work_log(
                            f"\n—— 中期落盘完成；后台补下 {n_def} 封 ——\n"
                        )
                    else:
                        self.summary_var.set(
                            f"全部主轮完成，已登记 {n_ok}/{len(queue)} 封。\n\n{guess}"
                        )
                        self.work_status_var.set("主轮完成，正在收尾…")
                        self.add_work_log("\n—— 中期落盘完成 ——\n")

                self.root.after(0, _ui_mid)

            def on_bg_success(results, failed_details, stage):
                try:
                    self.downloader.generate_download_report(
                        results, queue, failed_details, stage="后台补下增量"
                    )
                except Exception as e:
                    logger.error(f"后台增量落盘失败: {e}")
                self.root.after(0, self.refresh_inventory)

            outcome = self.downloader.download_multiple_emails(
                queue,
                progress_callback=progress,
                soft_attempts=2,
                hard_attempts=10,
                mid_phase_callback=mid_phase,
                on_result_callback=on_bg_success,
            )
            results = outcome.get("results") or []
            failed_details = outcome.get("failed_details") or []
            final_failed = outcome.get("final_failed") or []
            cancelled = outcome.get("cancelled", False)
            success = len([r for r in results if r.get("status") == "success"])
            skipped = len([r for r in results if r.get("status") == "skipped"])

            def finish():
                try:
                    if not self._mid_flushed:
                        self._flush_download_artifacts(
                            results, queue, failed_details,
                            stage="最终汇总", open_viz=False,
                        )
                    else:
                        self.downloader.generate_download_report(
                            results, queue, failed_details, stage="最终汇总"
                        )
                        guess2 = self.smart_guess_receivers(notify=False) or mid_box.get("guess", "")
                        self.refresh_inventory()
                        self.refresh_oc_tree()
                        self._refresh_name_combos(
                            getattr(self, "inv_sender_combo", None),
                            getattr(self, "inv_receiver_combo", None),
                        )
                        mid_box["guess"] = guess2

                    guess = mid_box.get("guess", "")
                    self.work_progress["value"] = 100
                    fail_subjects = [
                        (e.get("subject") or str(e.get("id"))) for e in final_failed
                    ]
                    empty_n = self._count_empty_receivers()
                    empty_hint = (
                        f"另有 {empty_n} 封收件人为空：请到「OC 管理」补别名并点「补全收件人」，"
                        f"仍空再到「信件库存」手工改。"
                        if empty_n else "收件人已齐，可打开可视化。"
                    )
                    if cancelled:
                        self.summary_var.set(
                            f"已提前结束。新下载 {success}，跳过 {skipped}。\n{empty_hint}\n\n{guess}"
                        )
                        self.work_status_var.set("已结束（当前进度）")
                    elif fail_subjects:
                        listed = "；".join(fail_subjects[:12])
                        more = f" 等共 {len(fail_subjects)} 封" if len(fail_subjects) > 12 else ""
                        msg = (
                            f"入库完成。新下载 {success}，跳过 {skipped}。\n"
                            f"{empty_hint}\n"
                            f"另有部分任务【{listed}{more}】因网络失败，"
                            f"可重试或通过「信件库存 → 手动录入」导入。\n\n{guess}"
                        )
                        self.summary_var.set(msg)
                        self.work_status_var.set("入库完成（部分失败）")
                        self.add_work_log("\n—— 最终汇总：部分信件失败 ——\n")
                        messagebox.showinfo("下载完成", msg)
                    else:
                        self.summary_var.set(
                            f"入库完成。新下载 {success}，跳过 {skipped}。\n{empty_hint}\n\n{guess}"
                        )
                        self.work_status_var.set("入库完成")
                        self.add_work_log("\n—— 最终汇总：全部成功 ——\n")
                    # 提前结束已在中期弹过；正常收尾在此决定开可视化或提示补收件人
                    if not cancelled:
                        self._after_ingest_open_or_prompt()
                finally:
                    self.busy = False
                    try:
                        self.cancel_download_btn.configure(state=tk.DISABLED)
                    except Exception:
                        pass

            self.root.after(0, finish)
        except Exception as e:
            logger.error(f"入库失败: {e}")
            self.root.after(0, lambda err=e: messagebox.showerror("错误", f"入库失败: {err}"))
            self.root.after(0, lambda: self.work_status_var.set("入库失败"))
            self.root.after(0, lambda: setattr(self, "busy", False))
            self.root.after(
                0,
                lambda: self.cancel_download_btn.configure(state=tk.DISABLED)
                if hasattr(self, "cancel_download_btn") else None,
            )

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

    @staticmethod
    def _is_unclear_letter_sender(name):
        """仅空或占位「未知」算不明；?/？/数字/符号 ID 都是合法发件人。"""
        s = (name or "").strip()
        return (not s) or s == "未知"

    def _count_unclear_senders(self):
        """统计 global_result 里发件人仍为空/未知的信件。"""
        path = self.downloader.base_dir / "global_result.tsv"
        if not path.exists():
            return 0, []
        unclear = []
        try:
            lines = path.read_text(encoding="utf-8").splitlines()
        except Exception:
            return 0, []
        for line in lines:
            if not (line.startswith("收\t") or line.startswith("发\t")):
                continue
            cols = line.split("\t")
            if len(cols) < 6:
                continue
            # 发广播 A 不应在库；若残留则不算「发件人不明」提醒对象
            if (cols[0] or "").strip() == "发" and EmailDownloader.is_bare_broadcast_a(cols[4]):
                continue
            sender = (cols[2] or "").strip()
            if self._is_unclear_letter_sender(sender):
                unclear.append(
                    f"{cols[0]} {cols[1]} {(cols[4] or '').strip() or '(无类型)'} → {cols[5]}"
                )
        return len(unclear), unclear

    def open_visualization_html(self):
        html_path = script_dir / "visualization_private.html"
        if not html_path.exists():
            messagebox.showerror("错误", f"未找到可视化文件:\n{html_path}")
            return False
        n, samples = self._count_unclear_senders()
        if n:
            preview = "\n".join(samples[:5])
            more = f"\n…另有 {n - 5} 封" if n > 5 else ""
            messagebox.showwarning(
                "发件人不明",
                f"有 {n} 封信件发件人为空或「未知」。\n"
                f"（?/符号/数字 ID 视为合法名字，不会因此提醒。）\n"
                f"可在「信件库存」里补全后再打开可视化。\n\n"
                f"{preview}{more}",
            )
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
        ttk.Label(
            path_box,
            text="弱网/大附件建议 ≥300；过短易导致下载卡住或反复失败。",
            foreground="gray",
        ).grid(row=4, column=2, sticky=tk.W, padx=6)
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

        maint_box = ttk.LabelFrame(inner, text="库存维护", padding=(8, 6))
        maint_box.pack(fill=tk.X, pady=6)
        ttk.Label(
            maint_box,
            text=(
                "日常入库/手动录入只会「单封」注册附件，不会覆盖你改过的其他信。\n"
                "此处「修复」会按全部信夹全量重扫并覆盖 attachment_report.tsv。"
            ),
            justify=tk.LEFT,
            wraplength=500,
            foreground="gray",
        ).pack(anchor=tk.W, pady=(0, 6))
        ttk.Button(
            maint_box,
            text="全量修复附件库存…",
            command=self.repair_attachment_inventory,
        ).pack(anchor=tk.W)

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

    def repair_attachment_inventory(self):
        """手动根据 global_result.tsv 重扫信夹，重建 attachment_report.tsv。"""
        tsv = self.downloader.base_dir / "global_result.tsv"
        if not tsv.exists():
            messagebox.showwarning("提示", "还没有 global_result.tsv，请先下载入库。")
            return
        try:
            from uitl_test import extract_attachments_from_tsv
            out = self.downloader.regenerate_attachment_report()
            if not out:
                messagebox.showerror("错误", "重建附件报告失败，请查看日志。")
                return
            results = extract_attachments_from_tsv(tsv)
            n_files = sum(len(r.get("file_attachments") or []) for r in results)
            n_text = sum(len(r.get("text_attachments") or []) for r in results)
            self.refresh_inventory()
            messagebox.showinfo(
                "已修复",
                f"已重建附件库存报告：\n"
                f"· 含附件信件：{len(results)} 封\n"
                f"· 文件附件：{n_files}\n"
                f"· 文字附件行：{n_text}\n\n"
                f"请刷新可视化页面查看。",
            )
        except Exception as e:
            logger.error(f"手动修复附件库存失败: {e}")
            messagebox.showerror("错误", f"修复失败: {e}")

    # ==================== 库存 ====================
    def _build_inventory_tab(self):
        top = ttk.Frame(self.inventory_frame)
        top.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(
            top,
            text="已入库信件。",
        ).pack(side=tk.LEFT)
        ttk.Button(top, text="手动录入…", command=self.open_manual_inventory_dialog).pack(side=tk.RIGHT, padx=4)
        ttk.Button(top, text="刷新", command=self.refresh_inventory).pack(side=tk.RIGHT)

        cols = ("收/发", "日期", "发件人", "收件人", "信件类型", "附件", "路径")
        wrap = ttk.Frame(self.inventory_frame)
        wrap.pack(fill=tk.BOTH, expand=True)
        self.inv_tree = ttk.Treeview(wrap, columns=cols, show="headings", selectmode="browse")
        vsb = ttk.Scrollbar(wrap, orient=tk.VERTICAL, command=self.inv_tree.yview)
        self.inv_tree.configure(yscroll=vsb.set)
        self.inv_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        for c, w in zip(cols, (50, 90, 110, 110, 80, 50, 200)):
            self.inv_tree.heading(c, text=c)
            self.inv_tree.column(c, width=w)
        self.inv_tree.bind("<<TreeviewSelect>>", self._on_inv_select)
        self.inv_tree.bind("<Double-1>", self._on_inv_tree_double_click)
        ttk.Label(
            self.inventory_frame,
            text="双击一封信可打开本地 content.txt 正文（不联网）",
            foreground="gray",
        ).pack(anchor=tk.W)

        edit = ttk.LabelFrame(self.inventory_frame, text="编辑所选", padding=(8, 6))
        edit.pack(fill=tk.X, pady=6)
        edit.columnconfigure(1, weight=1)
        edit.columnconfigure(3, weight=1)
        self.inv_sender_var = tk.StringVar()
        self.inv_receiver_var = tk.StringVar()
        self.inv_type_var = tk.StringVar()
        self.inv_meta_var = tk.StringVar(value="")
        self.inv_has_attach_var = tk.StringVar(value="无")
        ttk.Label(edit, textvariable=self.inv_meta_var).grid(row=0, column=0, columnspan=4, sticky=tk.W, pady=2)
        ttk.Label(edit, text="发件人").grid(row=1, column=0, sticky=tk.W, pady=3)
        self.inv_sender_combo = self._make_name_combo(edit, self.inv_sender_var, width=18)
        self.inv_sender_combo.grid(row=1, column=1, padx=4, sticky=tk.EW)
        ttk.Label(edit, text="收件人").grid(row=1, column=2, sticky=tk.W, padx=(8, 0))
        self.inv_receiver_combo = self._make_name_combo(edit, self.inv_receiver_var, width=18)
        self.inv_receiver_combo.grid(row=1, column=3, padx=4, sticky=tk.EW)
        ttk.Label(
            edit,
            text="可直接输入，或下拉/Tab 从已有 OC 与库存人名补全",
            foreground="gray",
        ).grid(row=4, column=0, columnspan=4, sticky=tk.W, pady=(4, 0))
        self._inv_cur_direction = ""
        inv_type_tip = ttk.Label(edit, wraplength=520, justify=tk.LEFT)
        inv_type_tip.grid(row=5, column=0, columnspan=4, sticky=tk.W)
        self._refresh_inv_type_tip = self._bind_type_feedback(
            inv_type_tip, self.inv_type_var, lambda: self._inv_cur_direction
        )
        ttk.Label(edit, text="信件类型").grid(row=2, column=0, sticky=tk.W, pady=3)
        ttk.Entry(edit, textvariable=self.inv_type_var, width=18).grid(row=2, column=1, padx=4, pady=3, sticky=tk.W)
        ttk.Label(edit, text="附件").grid(row=2, column=2, sticky=tk.W, padx=(8, 0))
        ttk.Label(edit, textvariable=self.inv_has_attach_var).grid(row=2, column=3, sticky=tk.W, padx=4)
        btn_row = ttk.Frame(edit)
        btn_row.grid(row=3, column=0, columnspan=4, sticky=tk.EW, pady=(6, 0))
        ttk.Button(btn_row, text="编辑附件…", command=self.open_edit_attachments_dialog).pack(side=tk.LEFT)
        ttk.Button(btn_row, text="保存修改", command=self.save_inventory_row).pack(side=tk.RIGHT)

        foot = ttk.Frame(self.inventory_frame)
        foot.pack(fill=tk.X, pady=(4, 0))
        ttk.Button(foot, text="打开该信文件夹", command=self.open_selected_mail_folder).pack(side=tk.LEFT)
        ttk.Button(foot, text="打开邮件总文件夹", command=self.open_storage_folder).pack(side=tk.LEFT, padx=6)
        ttk.Button(foot, text="打开[地缝可视化]", command=self.open_visualization_html).pack(side=tk.RIGHT)

        self._inv_rows = []  # parallel to tree children order
        self._attach_index = {}

    def _inventory_attach_flag(self, path_str):
        from uitl_test import normalize_letter_path_key
        key = normalize_letter_path_key(path_str)
        return "有" if self._attach_index.get(key) else "无"

    def _reload_attach_index(self):
        from uitl_test import attachment_paths_index
        report = self.downloader.base_dir / "attachment_report.tsv"
        try:
            self._attach_index = attachment_paths_index(report)
        except Exception:
            self._attach_index = {}

    def _resolve_inv_email_dir(self, row):
        from uitl_test import _resolve_email_dir
        tsv = self.downloader.base_dir / "global_result.tsv"
        ed = _resolve_email_dir(row.get("path"), tsv if tsv.exists() else self.downloader.base_dir)
        if ed:
            return Path(ed)
        rel = row.get("path") or ""
        if rel.startswith(".\\"):
            rel = rel[2:]
        cand = script_dir / str(rel).replace("\\", os.sep)
        return cand if cand.is_dir() else None

    def _path_fmt_for_dir(self, email_dir):
        email_dir = Path(email_dir)
        try:
            rel = email_dir.resolve().relative_to(script_dir.resolve())
            return ".\\" + str(rel).replace("/", "\\")
        except Exception:
            try:
                rel = email_dir.resolve().relative_to(self.downloader.base_dir.parent.resolve())
                return ".\\" + str(rel).replace("/", "\\")
            except Exception:
                return str(email_dir)

    def refresh_inventory(self):
        path = self.downloader.base_dir / "global_result.tsv"
        for item in self.inv_tree.get_children():
            self.inv_tree.delete(item)
        self._inv_rows = []
        self._inv_header_lines = []
        self._reload_attach_index()
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
            flag = self._inventory_attach_flag(row["path"])
            row["has_attach"] = flag
            self._inv_rows.append(row)
            recv_show = (row["receiver"] or "").strip() or "[未知的收信人]"
            self.inv_tree.insert(
                "",
                tk.END,
                values=(
                    row["direction"], row["date"], row["sender"], recv_show,
                    row["letter_type"], flag, row["path"],
                ),
            )
        self.status_var.set(f"库存 {len(self._inv_rows)} 封")
        self._refresh_name_combos(
            getattr(self, "inv_sender_combo", None),
            getattr(self, "inv_receiver_combo", None),
        )

    def _on_inv_select(self, _event=None):
        sel = self.inv_tree.selection()
        if not sel:
            return
        idx = self.inv_tree.index(sel[0])
        if idx >= len(self._inv_rows):
            return
        row = self._inv_rows[idx]
        flag = self._inventory_attach_flag(row["path"])
        row["has_attach"] = flag
        self.inv_meta_var.set(f"{row['direction']}  {row['date']}  附件:{flag}")
        self._inv_cur_direction = row["direction"]
        self.inv_sender_var.set(row["sender"])
        self.inv_receiver_var.set(row["receiver"])
        self.inv_type_var.set(row["letter_type"])
        self._refresh_inv_type_tip()
        self.inv_has_attach_var.set(flag)

    def _on_inv_tree_double_click(self, event):
        iid = self.inv_tree.identify_row(event.y)
        if not iid:
            return
        idx = self.inv_tree.index(iid)
        if idx < len(self._inv_rows):
            self._open_inventory_content_preview(self._inv_rows[idx])

    def _open_inventory_content_preview(self, row):
        """弹窗打开本地 content.txt 正文。"""
        email_dir = self._resolve_inv_email_dir(row)
        if not email_dir:
            messagebox.showwarning("提示", f"找不到信件目录：\n{row.get('path')}")
            return
        content_path = Path(email_dir) / "content.txt"
        if not content_path.is_file():
            messagebox.showwarning("提示", f"没有 content.txt：\n{content_path}")
            return
        try:
            text = content_path.read_text(encoding="utf-8")
        except Exception as e:
            messagebox.showerror("错误", f"读取失败：{e}")
            return

        win = tk.Toplevel(self.root)
        title = f"{row.get('direction')} {row.get('letter_type')} · {row.get('sender')} → {row.get('receiver') or '？'}"
        win.title("正文：" + title[:60])
        win.geometry("720x580")
        win.transient(self.root)
        head = ttk.Frame(win, padding=(10, 8))
        head.pack(fill=tk.X)
        ttk.Label(
            head,
            text=(
                f"方向：{row.get('direction')}　日期：{row.get('date')}　类型：{row.get('letter_type')}\n"
                f"发件人：{row.get('sender')}　收件人：{row.get('receiver') or '[未知的收信人]'}\n"
                f"{content_path}"
            ),
            justify=tk.LEFT,
            wraplength=680,
        ).pack(anchor=tk.W)
        box = scrolledtext.ScrolledText(win, wrap=tk.WORD)
        box.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        box.insert(tk.END, text)
        box.configure(state="disabled")
        bar = ttk.Frame(win, padding=(10, 0))
        bar.pack(fill=tk.X, pady=(0, 10))
        ttk.Button(bar, text="打开文件夹", command=lambda: self.open_selected_mail_folder()).pack(side=tk.LEFT)
        ttk.Button(bar, text="关闭", command=win.destroy).pack(side=tk.RIGHT)

    def save_inventory_row(self):
        sel = self.inv_tree.selection()
        if not sel:
            messagebox.showinfo("提示", "请先选一封信。")
            return
        idx = self.inv_tree.index(sel[0])
        if idx >= len(self._inv_rows):
            return
        row = self._inv_rows[idx]
        letter_type = self.inv_type_var.get().strip()
        implied_receiver = ""
        if letter_type:
            checked = self._checked_letter_type(self.inv_type_var, row.get("direction", ""))
            if not checked:
                return
            letter_type, implied_receiver = checked
        recv = self.inv_receiver_var.get().strip()
        if recv in ("[未知的收信人]", "未知的收信人"):
            recv = ""
        if not recv and implied_receiver:
            recv = implied_receiver
            self.inv_receiver_var.set(recv)
        is_broadcast_send = row.get("direction") == "发" and letter_type and not implied_receiver
        if not recv and not is_broadcast_send and not self._confirm_missing_receiver():
            return
        row["sender"] = self.inv_sender_var.get().strip()
        row["receiver"] = recv
        row["letter_type"] = letter_type
        flag = self._inventory_attach_flag(row["path"])
        row["has_attach"] = flag
        self.inv_tree.item(
            sel[0],
            values=(
                row["direction"], row["date"], row["sender"], row["receiver"],
                row["letter_type"], flag, row["path"],
            ),
        )
        self._write_inventory_tsv(register_paths=[row["path"]])
        # 元数据变更后只 upsert 本信附件行（发件人等字段同步进报告）
        try:
            self.downloader.register_letter_attachments(row)
        except Exception as e:
            logger.error(f"保存库存后单封附件注册失败: {e}")
        self._reload_attach_index()
        self._on_inv_select()
        messagebox.showinfo("已保存", "修改已写入库存；本信附件已单独注册。")

    def open_edit_attachments_dialog(self):
        """编辑所选信件的文字附件行与文件附件（只注册本信）。"""
        sel = self.inv_tree.selection()
        if not sel:
            messagebox.showinfo("提示", "请先选一封信。")
            return
        idx = self.inv_tree.index(sel[0])
        row = self._inv_rows[idx]
        email_dir = self._resolve_inv_email_dir(row)
        if not email_dir:
            messagebox.showerror("错误", f"找不到信件文件夹:\n{row.get('path')}")
            return

        from uitl_test import scan_letter_attachments, load_attachment_report, normalize_letter_path_key

        dlg = tk.Toplevel(self.root)
        dlg.title(f"编辑附件 — {row.get('letter_type', '')} {row.get('sender', '')}")
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.geometry("640x480")

        text_atts, file_atts = scan_letter_attachments(email_dir, project_root=script_dir)
        # 若报告里有人工维护过的文字行，优先用报告（磁盘扫描仍是文件真源）
        try:
            key = normalize_letter_path_key(row["path"])
            for r in load_attachment_report(self.downloader.base_dir / "attachment_report.tsv"):
                if normalize_letter_path_key(r.get("directory")) == key and r.get("text_attachments"):
                    text_atts = list(r["text_attachments"])
                    break
        except Exception:
            pass

        frm = ttk.Frame(dlg, padding=10)
        frm.pack(fill=tk.BOTH, expand=True)
        frm.rowconfigure(1, weight=1)
        frm.rowconfigure(3, weight=1)
        frm.columnconfigure(0, weight=1)

        ttk.Label(frm, text="文字附件（每行一条；正文含「附件」的行也会被登记）").grid(row=0, column=0, sticky=tk.W)
        text_box = scrolledtext.ScrolledText(frm, height=8, wrap=tk.WORD)
        text_box.grid(row=1, column=0, sticky=tk.NSEW, pady=4)
        text_box.insert("1.0", "\n".join(text_atts))

        ttk.Label(frm, text="文件附件（保存在该信文件夹内）").grid(row=2, column=0, sticky=tk.W, pady=(8, 0))
        file_frame = ttk.Frame(frm)
        file_frame.grid(row=3, column=0, sticky=tk.NSEW)
        file_frame.columnconfigure(0, weight=1)
        file_frame.rowconfigure(0, weight=1)
        file_list = tk.Listbox(file_frame, height=8)
        file_list.grid(row=0, column=0, sticky=tk.NSEW)
        fsb = ttk.Scrollbar(file_frame, orient=tk.VERTICAL, command=file_list.yview)
        fsb.grid(row=0, column=1, sticky=tk.NS)
        file_list.configure(yscrollcommand=fsb.set)

        # 显示文件名；内部保留路径
        pending_files = []  # list of Path currently in folder or to copy in

        def _refresh_files():
            file_list.delete(0, tk.END)
            pending_files.clear()
            for p in sorted(email_dir.iterdir()):
                if p.is_file() and p.name != "content.txt":
                    pending_files.append(p)
                    file_list.insert(tk.END, p.name)

        _refresh_files()

        def _add_files():
            from tkinter import filedialog
            paths = filedialog.askopenfilenames(parent=dlg, title="选择要加入的附件文件")
            for raw in paths:
                src = Path(raw)
                if not src.is_file():
                    continue
                dest = email_dir / src.name
                if dest.resolve() != src.resolve():
                    shutil.copy2(src, dest)
            _refresh_files()

        def _remove_file():
            sel_i = file_list.curselection()
            if not sel_i:
                return
            p = pending_files[sel_i[0]]
            if messagebox.askyesno("确认", f"从信夹删除文件？\n{p.name}", parent=dlg):
                try:
                    p.unlink()
                except Exception as e:
                    messagebox.showerror("错误", f"删除失败: {e}", parent=dlg)
                _refresh_files()

        def _add_text_line():
            top = tk.Toplevel(dlg)
            top.title("添加文字附件行")
            top.transient(dlg)
            top.grab_set()
            var = tk.StringVar(value="【附件：】")
            ttk.Entry(top, textvariable=var, width=50).pack(padx=10, pady=10)

            def ok():
                t = var.get().strip()
                if t:
                    cur = text_box.get("1.0", tk.END).rstrip("\n")
                    text_box.delete("1.0", tk.END)
                    text_box.insert("1.0", (cur + "\n" + t).strip() + "\n")
                top.destroy()

            ttk.Button(top, text="添加", command=ok).pack(pady=6)

        bf = ttk.Frame(frm)
        bf.grid(row=4, column=0, sticky=tk.EW, pady=8)
        ttk.Button(bf, text="添加文字行", command=_add_text_line).pack(side=tk.LEFT, padx=2)
        ttk.Button(bf, text="添加文件…", command=_add_files).pack(side=tk.LEFT, padx=2)
        ttk.Button(bf, text="删除所选文件", command=_remove_file).pack(side=tk.LEFT, padx=2)

        def _save():
            # 把文字附件行写回 content：保留原正文，替换/追加含「附件」的登记行策略：
            # 用编辑框内容作为「登记用文字附件」；同步写入 content 末尾独立区块，并清理旧「附件」登记行可选。
            new_text_lines = [ln.strip() for ln in text_box.get("1.0", tk.END).splitlines() if ln.strip()]
            content_path = email_dir / "content.txt"
            if content_path.exists():
                raw = content_path.read_text(encoding="utf-8")
            else:
                raw = (
                    f"主题: {row.get('letter_type', '')}\n发件邮箱: \n收件邮箱: \n"
                    f"日期: {row.get('date', '')}\n{'-'*50}\n\n发件人=【{row.get('sender', '')}】\n"
                )
            # 去掉旧的「附件」行，再追加当前编辑的文字附件行，避免重复堆积
            kept = [ln for ln in raw.splitlines() if "附件" not in ln]
            if new_text_lines:
                kept.append("")
                kept.extend(new_text_lines)
            content_path.write_text("\n".join(kept).rstrip() + "\n", encoding="utf-8")

            self.downloader.register_letter_attachments(row)
            self._write_inventory_tsv(register_paths=None)  # 仅写 tsv 元数据不变也没关系
            self._reload_attach_index()
            # 刷新树
            flag = self._inventory_attach_flag(row["path"])
            row["has_attach"] = flag
            self.inv_tree.item(
                sel[0],
                values=(
                    row["direction"], row["date"], row["sender"], row["receiver"],
                    row["letter_type"], flag, row["path"],
                ),
            )
            self.inv_has_attach_var.set(flag)
            self.inv_meta_var.set(f"{row['direction']}  {row['date']}  附件:{flag}")
            dlg.destroy()
            messagebox.showinfo("已保存", "本信附件已更新并单独注册到附件报告。")

        ttk.Button(bf, text="取消", command=dlg.destroy).pack(side=tk.RIGHT, padx=2)
        ttk.Button(bf, text="保存并注册本信", command=_save).pack(side=tk.RIGHT, padx=2)

    def open_manual_inventory_dialog(self):
        """手动录入：收发人、类型、内容、附件；自动建夹并单封注册附件。"""
        dlg = tk.Toplevel(self.root)
        dlg.title("手动录入信件")
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.geometry("620x560")

        direction = tk.StringVar(value="收")
        date_var = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        sender_var = tk.StringVar()
        receiver_var = tk.StringVar()
        type_var = tk.StringVar()
        upload_files = []  # Path list chosen by user

        frm = ttk.Frame(dlg, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)
        frm.columnconfigure(1, weight=1)
        frm.rowconfigure(6, weight=1)

        ttk.Label(frm, text="收/发").grid(row=0, column=0, sticky=tk.W, pady=3)
        ttk.Combobox(frm, textvariable=direction, values=("收", "发"), width=8, state="readonly").grid(
            row=0, column=1, sticky=tk.W, pady=3
        )
        ttk.Label(frm, text="发信日期").grid(row=1, column=0, sticky=tk.W, pady=3)
        self._make_date_row(frm, date_var, width=12).grid(row=1, column=1, sticky=tk.W, pady=3)
        ttk.Label(frm, text="发件人").grid(row=2, column=0, sticky=tk.W, pady=3)
        self._make_name_combo(frm, sender_var).grid(row=2, column=1, sticky=tk.EW, pady=3)
        ttk.Label(frm, text="收件人").grid(row=3, column=0, sticky=tk.W, pady=3)
        self._make_name_combo(frm, receiver_var).grid(row=3, column=1, sticky=tk.EW, pady=3)
        ttk.Label(frm, text="可输入，或下拉/Tab 从已有人名补全", foreground="gray").grid(
            row=3, column=2, sticky=tk.W, padx=6
        )
        ttk.Label(frm, text="信件类型").grid(row=4, column=0, sticky=tk.W, pady=3)
        ttk.Entry(frm, textvariable=type_var).grid(row=4, column=1, sticky=tk.EW, pady=3)
        type_tip = ttk.Label(frm, wraplength=460, justify=tk.LEFT)
        type_tip.grid(row=5, column=1, columnspan=2, sticky=tk.W)
        self._bind_type_feedback(type_tip, type_var, direction.get, watch_vars=(direction,))

        ttk.Label(frm, text="正文内容").grid(row=6, column=0, sticky=tk.NW, pady=3)
        content_box = scrolledtext.ScrolledText(frm, height=12, wrap=tk.WORD)
        content_box.grid(row=6, column=1, sticky=tk.NSEW, pady=3)

        file_row = ttk.Frame(frm)
        file_row.grid(row=7, column=1, sticky=tk.EW, pady=4)
        file_label = ttk.Label(file_row, text="未选择文件附件")
        file_label.pack(side=tk.LEFT)

        def _pick_files():
            from tkinter import filedialog
            paths = filedialog.askopenfilenames(parent=dlg, title="上传附件文件（可多选）")
            upload_files.clear()
            upload_files.extend(Path(p) for p in paths)
            if upload_files:
                file_label.configure(text=f"已选 {len(upload_files)} 个文件：" + ", ".join(p.name for p in upload_files[:3])
                                     + ("…" if len(upload_files) > 3 else ""))
            else:
                file_label.configure(text="未选择文件附件")

        ttk.Button(file_row, text="上传文件附件…", command=_pick_files).pack(side=tk.RIGHT)

        def _ok():
            d = direction.get().strip()
            date_s = date_var.get().strip()
            sender = sender_var.get().strip()
            receiver = receiver_var.get().strip()
            letter_type = type_var.get().strip()
            body = content_box.get("1.0", tk.END).rstrip()
            if d not in ("收", "发"):
                messagebox.showwarning("提示", "收/发只能是「收」或「发」。", parent=dlg)
                return
            try:
                datetime.strptime(date_s, "%Y-%m-%d")
            except Exception:
                messagebox.showwarning("提示", "日期格式应为 YYYY-MM-DD。", parent=dlg)
                return
            checked = self._checked_letter_type(type_var, d, parent=dlg)
            if not checked:
                return
            letter_type, implied_receiver = checked
            if not receiver and implied_receiver:
                receiver = implied_receiver
                receiver_var.set(receiver)
            if not sender:
                messagebox.showwarning("提示", "请填写发件人。", parent=dlg)
                return
            # 发 B / A+B 是广播，无固定收件人；其余空收件人只提醒
            if not receiver and not (d == "发" and not implied_receiver):
                if not self._confirm_missing_receiver(dlg):
                    return

            base = self.downloader.base_dir
            date_folder = date_s.replace("-", "")
            safe_sender = self.downloader._safe_filename(sender)[:30]
            safe_type = self.downloader._safe_filename(letter_type)[:40]
            folder_name = f"{date_folder}_{safe_type}来自{safe_sender}的信"
            email_dir = base / d / folder_name
            email_dir.mkdir(parents=True, exist_ok=True)

            # 正文：若用户未写发件人标记，自动补一行，便于解析
            body_out = body
            if "发件人" not in body_out:
                body_out = f"发件人=【{sender}】\n{body_out}"
            content_path = email_dir / "content.txt"
            content_path.write_text(
                f"主题: {letter_type}来自{sender}的信\n"
                f"发件邮箱: \n"
                f"收件邮箱: \n"
                f"日期: {date_s}\n"
                f"{'-'*50}\n\n"
                f"{body_out}\n",
                encoding="utf-8",
            )
            for src in upload_files:
                if src.is_file():
                    shutil.copy2(src, email_dir / src.name)

            path_fmt = self._path_fmt_for_dir(email_dir)
            row = {
                "direction": d,
                "date": date_s,
                "sender": sender,
                "receiver": receiver,
                "letter_type": letter_type,
                "path": path_fmt,
            }
            # 同路径覆盖，否则追加
            updated = False
            for old in self._inv_rows:
                if old.get("path") == path_fmt:
                    old.update(row)
                    updated = True
                    break
            if not updated:
                self._inv_rows.append(row)

            self._write_inventory_tsv(register_paths=None)
            # 只注册本信：扫描正文「附件」行 + 上传的文件
            self.downloader.register_letter_attachments(row)
            self.refresh_inventory()
            dlg.destroy()
            messagebox.showinfo("已录入", "已写入库存，并仅为本信注册附件。")

        btns = ttk.Frame(frm)
        btns.grid(row=8, column=0, columnspan=2, sticky=tk.E, pady=(10, 0))
        ttk.Button(btns, text="取消", command=dlg.destroy).pack(side=tk.RIGHT, padx=4)
        ttk.Button(btns, text="保存到库存", command=_ok).pack(side=tk.RIGHT)

    def _write_inventory_tsv(self, register_paths=None):
        """写入 global_result.tsv。默认不再全量重建附件报告。

        register_paths: 若传入路径列表，仅对这些信做单封附件注册（可选）。
        """
        path = self.downloader.base_dir / "global_result.tsv"
        path.parent.mkdir(parents=True, exist_ok=True)
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
        if register_paths:
            for p in register_paths:
                try:
                    self.downloader.register_letter_attachments(p)
                except Exception as e:
                    logger.error(f"单封附件注册失败 {p}: {e}")

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