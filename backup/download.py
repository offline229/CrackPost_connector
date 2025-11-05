import imaplib
import email
import os
import re
import logging
import time
import configparser
from datetime import datetime
from pathlib import Path
from getpass import getpass
import imaplib

# 在文件开头添加导入
import imapclient

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("email_downloader.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

class EmailDownloader:
    def __init__(self, config_file="email_config.ini"):
        """初始化邮件下载器"""
        self.config = self._load_config(config_file)
        self.email_addr = self.config.get('Credentials', 'email')
        self.password = self.config.get('Credentials', 'password')
        self.imap_server = self.config.get('Servers', 'imap_server')
        
        # 搜索规则 - 默认为 A数字 模式
        self.search_pattern = self.config.get('Filters', 'search_pattern', fallback=r'A\d+')
        
        # 创建基本目录
        self.base_dir = Path(self.config.get('Storage', 'base_directory', fallback='email_data'))
        self.base_dir.mkdir(exist_ok=True)
        
        # 邮箱连接
        self.mail = None
    # filepath: [download.py](http://_vscodecontentref_/1)
    # 在类的开头添加此函数
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
            return header  # 返回原始标题作为后备选项
    def _load_config(self, config_file):
        """加载或创建配置文件"""
        config = configparser.ConfigParser()
        
        if not os.path.exists(config_file):
            # 创建配置文件
            config['Credentials'] = {
                'email': input("请输入您的邮箱地址: "),
                'password': getpass("请输入您的邮箱授权码(而非普通登录密码): ")
            }
            
            config['Servers'] = {
                'imap_server': input("请输入IMAP服务器地址(例如 imap.gmail.com): "),
            }
            
            config['Filters'] = {
                'search_pattern': input("请输入搜索模式正则表达式(默认为 A\\d+ 即A后跟数字): ") or r'A\d+',
            }
            
            config['Storage'] = {
                'base_directory': input("请输入存储目录(默认为 email_data): ") or "email_data"
            }
            
            with open(config_file, 'w') as f:
                config.write(f)
            
            logger.info(f"配置文件已创建: {config_file}")
        else:
            config.read(config_file)
            
        return config
    
    def connect_to_email(self):
        """连接到邮箱"""
        try:
            logger.info(f"连接到IMAP服务器: {self.imap_server}")
            self.mail = imaplib.IMAP4_SSL(self.imap_server)
            self.mail.login(self.email_addr, self.password)
            self.mail.select('inbox')
            print(f"\n✅ 成功登录到 {self.email_addr} 的邮箱\n")
            logger.info(f"成功登录到 {self.email_addr} 的邮箱")
            return True
        except Exception as e:
            logger.error(f"连接邮箱失败: {str(e)}")
            print(f"\n❌ 连接失败: {str(e)}\n")
            return False
    
    def close_connection(self):
        """关闭邮箱连接"""
        if self.mail:
            try:
                self.mail.close()
                self.mail.logout()
                logger.info("已关闭邮箱连接")
            except Exception as e:
                logger.error(f"关闭连接时发生错误: {str(e)}")
    def list_mailboxes(self):
        """列出所有邮箱文件夹"""
        if not self.mail:
            self.connect_to_email()
        status, mailboxes = self.mail.list()
        print("所有邮箱文件夹：")
        for box in mailboxes:
            print(box.decode())
    # 替换get_all_mailboxes方法
    def get_all_mailboxes(self):
        """返回所有邮箱文件夹名称列表（使用imapclient自动解码）"""
        # 创建IMAPClient实例
        client = imapclient.IMAPClient(self.imap_server, ssl=True)
        client.login(self.email_addr, self.password)
        
        # 获取文件夹列表（自动解码）
        folders = client.list_folders()
        
        # 创建映射字典 - 保存原始名称用于select操作
        self.mailbox_mapping = {}
        mailbox_names = []
        
        for flags, delimiter, name in folders:
            # IMAPClient已自动解码，但我们需要记住原始名称
            raw_name = name  # 这是IMAPClient已经解码的名称
            display_name = name  # 显示用的名称
            
            # 保存映射
            self.mailbox_mapping[display_name] = raw_name
            mailbox_names.append(display_name)
        
        # 关闭连接
        client.logout()
        return mailbox_names
    def select_mailbox(self, mailbox="INBOX"):
        """选择指定邮箱文件夹"""
        self.mail.select(mailbox)
        logger.info(f"已选择文件夹: {mailbox}")
    def search_emails(self, current_folder=None):
        """搜索当前文件夹内所有符合规则的邮件"""
        if not self.mail:
            if not self.connect_to_email():
                return []
        try:
            status, data = self.mail.search(None, 'ALL')
            if status != 'OK':
                logger.warning(f"search ALL 失败: {status}, {data}")
                return []
            all_email_ids = data[0].split()
            pattern = re.compile(self.search_pattern, re.IGNORECASE)
            matched_emails = []
            for email_id in all_email_ids:
                if isinstance(email_id, bytes):
                    email_id = email_id.decode()
                try:
                    status, fetch_data = self.mail.fetch(email_id, '(BODY[HEADER.FIELDS (SUBJECT FROM DATE)])')
                    if status != 'OK':
                        continue
                    header_data = fetch_data[0][1]
                    msg = email.message_from_bytes(header_data)
                    subject_header = msg['subject']
                    if not subject_header:
                        continue
                    decoded_header = email.header.decode_header(subject_header)
                    subject = ""
                    for part, encoding in decoded_header:
                        if isinstance(part, bytes):
                            subject += part.decode(encoding or 'utf-8', errors='replace')
                        else:
                            subject += str(part)
                    if pattern.search(subject):
                        sender = email.utils.parseaddr(msg['from'])[1]
                        date_str = msg['date'] or ''
                        matched_emails.append({
                            'id': email_id,
                            'subject': subject,
                            'sender': sender,
                            'date': date_str,
                            'folder': current_folder
                        })
                except Exception as e:
                    logger.error(f"处理邮件ID {email_id} 时出错: {str(e)}")
                    continue
            logger.info(f"找到 {len(matched_emails)} 封匹配邮件")
            print(f"\n共找到 {len(matched_emails)} 封匹配 '{self.search_pattern}' 的邮件")
            return matched_emails
        except Exception as e:
            logger.error(f"搜索邮件时出错: {str(e)}")
            print(f"❌ 搜索邮件时出错: {str(e)}")
            return []
    def download_email(self, email_id, folder=None, client=None):
        """下载单个邮件的内容和附件"""
        if not self.mail:
            if not self.connect_to_email():
                return None
        
        try:
            # 如果没有传入client，则创建新的连接
            if not client:
                client = imapclient.IMAPClient(self.imap_server, ssl=True)
                client.login(self.email_addr, self.password)
                if folder:
                    client.select_folder(mailbox, readonly=True)
            
            # 获取完整邮件
            fetch_data = client.fetch([email_id], ['RFC822'])
            raw_email = fetch_data[email_id][b'RFC822']
            msg = email.message_from_bytes(raw_email)
            # 如果提供了文件夹，先选择到该文件夹
            if folder:
                self.mail.select(folder)
                logger.info(f"已切换到文件夹: {folder}")
            
            # 增加更严格的错误检查
            try:
                status, data = self.mail.fetch(email_id, '(RFC822)')
                if status != 'OK' or not data or len(data) == 0:
                    logger.warning(f"获取邮件 {email_id} 失败: {status}, {data}")
                    return None
                    
                # 检查数据格式
                if not isinstance(data[0], tuple) or len(data[0]) < 2:
                    logger.error(f"邮件数据格式异常: {data}")
                    return None
                    
                raw_email = data[0][1]
            except Exception as e:
                logger.error(f"获取邮件数据时出错: {str(e)}")
                return None
            
            raw_email = data[0][1]
            msg = email.message_from_bytes(raw_email)
            
            # 提取基本信息
            subject = email.header.decode_header(msg['subject'])[0][0]
            if isinstance(subject, bytes):
                subject = subject.decode(errors='replace')
            
            sender = email.utils.parseaddr(msg['from'])[1]
            date_str = msg['date']
            
            # 使用日期和主题创建唯一的目录名
            date_obj = email.utils.parsedate_to_datetime(date_str)
            date_folder = date_obj.strftime("%Y%m%d_%H%M%S")
            
            # 为此邮件创建目录
            safe_subject = re.sub(r'[^\w\s-]', '_', subject)
            if len(safe_subject) > 50:
                safe_subject = safe_subject[:50]
            
            email_dir = self.base_dir / f"{date_folder}_{safe_subject}"
            email_dir.mkdir(parents=True, exist_ok=True)
            
            # 保存邮件内容和附件
            body_text = ""
            body_html = ""
            has_attachments = False
            attachment_paths = []
            
            for part in msg.walk():
                content_type = part.get_content_type()
                content_disp = str(part.get("Content-Disposition"))
                
                # 提取纯文本正文
                if content_type == "text/plain" and "attachment" not in content_disp:
                    charset = part.get_content_charset()
                    try:
                        if charset:
                            body_text = part.get_payload(decode=True).decode(charset, errors='replace')
                        else:
                            body_text = part.get_payload(decode=True).decode(errors='replace')
                    except:
                        body_text = "无法解码邮件内容"
                    
                    # 保存正文到文本文件
                    with open(email_dir / "content.txt", "w", encoding="utf-8") as f:
                        f.write(f"主题: {subject}\n")
                        f.write(f"发件人: {sender}\n")
                        f.write(f"日期: {date_str}\n")
                        f.write("-" * 50 + "\n\n")
                        f.write(body_text)
                
                # 提取HTML正文
                elif content_type == "text/html" and "attachment" not in content_disp:
                    charset = part.get_content_charset()
                    try:
                        if charset:
                            body_html = part.get_payload(decode=True).decode(charset, errors='replace')
                        else:
                            body_html = part.get_payload(decode=True).decode(errors='replace')
                    except:
                        body_html = "<html><body>无法解码HTML内容</body></html>"
                    
                    # 保存HTML到文件
                    with open(email_dir / "content.html", "w", encoding="utf-8") as f:
                        f.write(body_html)
                
                # 下载附件
                elif "attachment" in content_disp or part.get_filename():
                    has_attachments = True
                    filename = part.get_filename()
                    if filename:
                        # 处理编码问题
                        if isinstance(filename, bytes):
                            filename = filename.decode(errors='replace')
                        
                        # 清理文件名
                        filename = re.sub(r'[^\w\s.-]', '_', filename)
                        file_path = email_dir / filename
                        
                        try:
                            # 保存附件
                            with open(file_path, 'wb') as f:
                                f.write(part.get_payload(decode=True))
                            
                            attachment_paths.append(str(file_path))
                            logger.info(f"保存附件: {file_path}")
                        except Exception as e:
                            logger.error(f"保存附件 {filename} 时出错: {str(e)}")
            
            result = {
                'subject': subject,
                'sender': sender,
                'date': date_str,
                'has_attachments': has_attachments,
                'directory': str(email_dir),
                'attachment_count': len(attachment_paths),
                'attachments': attachment_paths
            }
            
            logger.info(f"处理完成: {subject}")
            return result
            
        except Exception as e:
            logger.error(f"下载邮件 {email_id} 时出错: {str(e)}")
            return None
    
    def download_matched_emails(self, matched_emails):
        """下载所有匹配的邮件"""
        if not matched_emails:
            print("没有找到匹配的邮件")
            return []
        
        results = []
        total = len(matched_emails)
        
        print(f"\n开始下载 {total} 封邮件...")
        
        for i, email_info in enumerate(matched_emails):  # 不是matched_emails
            email_id = email_info['id']
            subject = email_info['subject']
            folder = email_info.get('folder')  # 获取邮件所属文件夹
            client = email_info.get('client')
            
            print(f"[{i+1}/{total}] 正在下载: {subject}")
            result = self.download_email(email_id, folder, client)
            
            if result:
                results.append(result)
                print(f"✅ 已下载到: {result['directory']}")
                if result['has_attachments']:
                    print(f"   包含 {result['attachment_count']} 个附件")
            else:
                print(f"❌ 下载失败: {subject}")
            
            # 避免过快请求导致服务器限制
            if i < total - 1:
                time.sleep(0.5)
        
        return results
    
    def run(self):
        """运行完整的下载流程（遍历所有文件夹）"""
        try:
            # 连接到邮箱
            if not self.connect_to_email():
                return

            # 获取所有文件夹
            mailboxes = self.get_all_mailboxes()
            print("\n将遍历以下文件夹：", mailboxes)

            all_matched_emails = []
            for mailbox in mailboxes:
                print(f"\n正在搜索文件夹: {mailbox}")
                try:
                    # 关键修改: 不使用原始编码的名称进行select
                    # 而是让IMAPClient处理文件夹选择和编码转换
                    client = imapclient.IMAPClient(self.imap_server, ssl=True)
                    client.login(self.email_addr, self.password)
                    
                    # 使用IMAPClient选择文件夹，它会自动处理编码问题
                    client.select_folder(mailbox, readonly=True)
                    
                    # 然后用它的搜索功能
                    matched_emails = []
                    messages = client.search(['ALL'])
                    
                    # 遍历找到的所有邮件
                    for msg_id in messages:
                        # 位于run方法中遍历邮件的部分
                        try:
                            # 获取邮件头部信息
                            header_data = client.fetch(msg_id, ['ENVELOPE', 'BODY[HEADER.FIELDS (SUBJECT FROM DATE)]'])
                            envelope = header_data[msg_id][b'ENVELOPE']
                            subject = self.decode_mime_header(envelope.subject.decode()) if envelope.subject else ""

                            
                            # 使用正则表达式匹配
                            if re.search(self.search_pattern, subject, re.IGNORECASE):
                                sender = envelope.from_[0].mailbox.decode() + '@' + envelope.from_[0].host.decode() if envelope.from_ else ""
                                
                                # 修复: 处理不同类型的日期对象
                                if envelope.date:
                                    if isinstance(envelope.date, datetime):
                                        date_str = envelope.date.strftime("%Y-%m-%d %H:%M:%S")
                                    elif isinstance(envelope.date, bytes):
                                        date_str = envelope.date.decode()
                                    else:
                                        date_str = str(envelope.date)
                                else:
                                    date_str = ""
                                
                                matched_emails.append({
                                    'id': msg_id,
                                    'subject': subject,
                                    'sender': sender,
                                    'date': date_str,
                                    'folder': mailbox,
                                    'client': client  # 保存client以便后续下载使用
                                })
                        except Exception as e:
                            logger.error(f"处理邮件ID {msg_id} 时出错: {str(e)}")
                            continue
                            
                    print("前五个匹配邮件标题：")
                    for i, mail in enumerate(matched_emails[:5]):
                        print(f"{i+1}: {mail['subject']}")
                    all_matched_emails.extend(matched_emails)
                    
                    # 用完后关闭连接
                    if not matched_emails:
                        client.logout()
                except Exception as e:
                    print(f"处理文件夹 {mailbox} 时出错: {str(e)}")
                    continue

            if not all_matched_emails:
                print("\n没有找到匹配的邮件")
                return

            # 显示匹配结果
            print(f"\n找到 {len(all_matched_emails)} 封匹配邮件:")
            print("-" * 80)
            print(f"{'序号':<6}{'日期':<20}{'发件人':<30}{'主题'}")
            print("-" * 80)
            
            # 修改run方法中的显示邮件列表部分
            for i, email_info in enumerate(all_matched_emails):
                # 截断过长的字段以保持格式整齐
                sender = email_info['sender']
                if len(sender) > 28:
                    sender = sender[:25] + "..."
                
                # 确保主题已经被正确解码
                subject = email_info['subject']
                if subject.startswith('=?'):
                    # 如果主题仍是MIME编码格式，尝试再次解码
                    try:
                        subject = self.decode_mime_header(subject)
                        header_parts = email.header.decode_header(subject)
                        subject = ""
                        for part, encoding in header_parts:
                            if isinstance(part, bytes):
                                subject += part.decode(encoding or 'utf-8', errors='replace')
                            else:
                                subject += str(part)
                    except Exception as e:
                        logger.error(f"解码主题失败: {subject}, 错误: {str(e)}")
                
                # 截断过长的主题
                if len(subject) > 40:
                    subject = subject[:37] + "..."
                
                # 尝试解析日期为更友好的格式
                try:
                    date_obj = email.utils.parsedate_to_datetime(email_info['date'])
                    date_str = date_obj.strftime("%Y-%m-%d %H:%M")
                except:
                    date_str = email_info['date'][:19]
                
                print(f"{i+1:<6}{date_str:<20}{sender:<30}{subject}")
            
            print("-" * 80)
            
            # 确认下载
            confirm = input("\n是否下载这些邮件? (y/n): ").strip().lower()
            if confirm != 'y':
                print("操作已取消")
                return
            
            # 下载邮件
            results = self.download_matched_emails(all_matched_emails)
            
            # 显示下载结果
            print("\n下载完成摘要:")
            print(f"总邮件数: {len(all_matched_emails)}")  # 修正：使用all_matched_emails
            print(f"成功下载: {len(results)}")
            print(f"下载失败: {len(all_matched_emails) - len(results)}")  # 修正：使用all_matched_emails
            print(f"存储目录: {self.base_dir}\n")
            
        except Exception as e:
            logger.error(f"运行时出错: {str(e)}")
            print(f"\n❌ 处理邮件时出错: {str(e)}")
        finally:
            # 关闭连接
            self.close_connection()

if __name__ == "__main__":
    downloader = EmailDownloader()
    downloader.run()