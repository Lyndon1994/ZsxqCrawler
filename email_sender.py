#!/usr/bin/env python3
"""
邮件发送模块
支持通过SMTP发送邮件
"""
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.utils import encode_rfc2231
from typing import List, Optional
import os
from urllib.parse import quote


class EmailSender:
    """邮件发送器"""
    
    def __init__(self, smtp_server: str, smtp_port: int, sender_email: str, 
                 sender_password: str, use_tls: bool = True):
        """初始化邮件发送器
        
        Args:
            smtp_server: SMTP服务器地址
            smtp_port: SMTP端口
            sender_email: 发件人邮箱
            sender_password: 发件人密码或应用专用密码
            use_tls: 是否使用TLS加密
        """
        self.smtp_server = smtp_server
        self.smtp_port = smtp_port
        self.sender_email = sender_email
        self.sender_password = sender_password
        self.use_tls = use_tls
    
    def send_email(self, receiver_emails: List[str], subject: str, 
                   body: str, attachments: Optional[List[str]] = None) -> bool:
        """发送邮件
        
        Args:
            receiver_emails: 收件人邮箱列表
            subject: 邮件主题
            body: 邮件正文（支持HTML）
            attachments: 附件文件路径列表
        
        Returns:
            bool: 发送成功返回True，失败返回False
        """
        try:
            # 创建邮件对象
            msg = MIMEMultipart()
            msg['From'] = self.sender_email
            msg['To'] = ', '.join(receiver_emails)
            msg['Subject'] = subject
            
            # 添加邮件正文
            msg.attach(MIMEText(body, 'html', 'utf-8'))
            
            # 添加附件
            if attachments:
                for file_path in attachments:
                    if os.path.exists(file_path):
                        filename = os.path.basename(file_path)
                        # 根据文件扩展名设置正确的MIME类型
                        if filename.lower().endswith('.pdf'):
                            maintype, subtype = 'application', 'pdf'
                        else:
                            maintype, subtype = 'application', 'octet-stream'
                        
                        with open(file_path, 'rb') as f:
                            part = MIMEBase(maintype, subtype)
                            part.set_payload(f.read())
                            encoders.encode_base64(part)
                            
                            # 使用RFC 2231编码处理中文文件名
                            # 同时提供filename和filename*两个参数以提高兼容性
                            encoded_filename = quote(filename)
                            part.add_header(
                                'Content-Disposition',
                                'attachment',
                                filename=('utf-8', '', filename),
                                # 备用ASCII文件名
                                filename_ascii=encoded_filename
                            )
                            msg.attach(part)
                    else:
                        print(f"⚠️ 附件不存在: {file_path}")
            
            # 连接SMTP服务器并发送邮件
            if self.use_tls:
                server = smtplib.SMTP(self.smtp_server, self.smtp_port)
                server.starttls()
            else:
                server = smtplib.SMTP_SSL(self.smtp_server, self.smtp_port)
            
            server.login(self.sender_email, self.sender_password)
            server.send_message(msg)
            server.quit()
            
            print(f"✅ 邮件发送成功: {subject}")
            print(f"   收件人: {', '.join(receiver_emails)}")
            return True
            
        except smtplib.SMTPAuthenticationError:
            print("❌ 邮件发送失败: SMTP认证失败，请检查邮箱和密码")
            return False
        except smtplib.SMTPException as e:
            print(f"❌ 邮件发送失败: SMTP错误 - {e}")
            return False
        except Exception as e:
            print(f"❌ 邮件发送失败: {e}")
            return False
    
    def send_simple_email(self, receiver_emails: List[str], subject: str, body: str) -> bool:
        """发送简单文本邮件
        
        Args:
            receiver_emails: 收件人邮箱列表
            subject: 邮件主题
            body: 邮件正文（纯文本）
        
        Returns:
            bool: 发送成功返回True，失败返回False
        """
        return self.send_email(receiver_emails, subject, body, None)
