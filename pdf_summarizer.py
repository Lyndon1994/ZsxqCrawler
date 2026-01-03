#!/usr/bin/env python3
"""
PDF文档自动总结与邮件发送脚本
功能：
1. 扫描指定目录中的PDF文件
2. 使用Azure OpenAI进行文档总结
3. 将总结结果通过邮件发送
"""
import os
import sys
import time
import argparse
from typing import List, Dict, Optional
from pathlib import Path
import json

# 导入依赖库
try:
    import tomllib
except ImportError:
    try:
        import tomli as tomllib
    except ImportError:
        print("⚠️ 需要安装tomli库来解析TOML配置文件")
        print("💡 请运行: pip install tomli")
        tomllib = None

try:
    from openai import AzureOpenAI
except ImportError:
    print("⚠️ 需要安装openai库来调用Azure OpenAI API")
    print("💡 请运行: pip install openai")
    AzureOpenAI = None

try:
    import PyPDF2
except ImportError:
    print("⚠️ 需要安装PyPDF2库来读取PDF文件")
    print("💡 请运行: pip install PyPDF2")
    PyPDF2 = None

from email_sender import EmailSender


class PDFSummarizer:
    """PDF文档总结器"""
    
    def __init__(self, config: dict):
        """初始化PDF总结器
        
        Args:
            config: 配置字典，包含azure_openai和email配置
        """
        self.config = config
        
        # Azure OpenAI配置
        openai_config = config.get('azure_openai', {})
        self.api_key = openai_config.get('api_key')
        self.endpoint = openai_config.get('endpoint')
        self.deployment_name = openai_config.get('deployment_name', 'gpt-4')
        self.api_version = openai_config.get('api_version', '2024-02-15-preview')
        self.system_prompt = openai_config.get('system_prompt', '你是一个专业的文档总结助手。')
        
        # 初始化Azure OpenAI客户端
        if AzureOpenAI and self.api_key and self.endpoint:
            self.client = AzureOpenAI(
                api_key=self.api_key,
                api_version=self.api_version,
                azure_endpoint=self.endpoint
            )
        else:
            self.client = None
            print("⚠️ Azure OpenAI未配置或库未安装")
        
        # 邮件配置
        email_config = config.get('email', {})
        smtp_server = email_config.get('smtp_server')
        smtp_port = email_config.get('smtp_port', 587)
        sender_email = email_config.get('sender_email')
        sender_password = email_config.get('sender_password')
        use_tls = email_config.get('use_tls', True)
        self.receiver_emails = email_config.get('receiver_email', '').split(',')
        self.receiver_emails = [email.strip() for email in self.receiver_emails if email.strip()]
        self.subject_template = email_config.get('subject_template', 'PDF文档总结: {filename}')
        
        # 初始化邮件发送器
        if smtp_server and sender_email and sender_password:
            self.email_sender = EmailSender(
                smtp_server=smtp_server,
                smtp_port=smtp_port,
                sender_email=sender_email,
                sender_password=sender_password,
                use_tls=use_tls
            )
        else:
            self.email_sender = None
            print("⚠️ 邮件配置不完整")
    
    def extract_text_from_pdf(self, pdf_path: str, max_pages: int = 50) -> str:
        """从PDF文件中提取文本
        
        Args:
            pdf_path: PDF文件路径
            max_pages: 最多读取的页数（避免文件太大）
        
        Returns:
            str: 提取的文本内容
        """
        if not PyPDF2:
            print("❌ PyPDF2库未安装，无法读取PDF")
            return ""
        
        try:
            text_content = []
            with open(pdf_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                total_pages = len(pdf_reader.pages)
                pages_to_read = min(total_pages, max_pages)
                
                print(f"📖 读取PDF: {os.path.basename(pdf_path)}")
                print(f"   总页数: {total_pages}, 读取页数: {pages_to_read}")
                
                for page_num in range(pages_to_read):
                    page = pdf_reader.pages[page_num]
                    text = page.extract_text()
                    if text:
                        text_content.append(text)
                
                full_text = '\n\n'.join(text_content)
                print(f"   提取字符数: {len(full_text)}")
                return full_text
                
        except Exception as e:
            print(f"❌ PDF读取失败: {e}")
            return ""
    
    def summarize_text(self, text: str, filename: str) -> Optional[tuple]:
        """使用Azure OpenAI总结文本
        
        Args:
            text: 要总结的文本
            filename: 文件名（用于提示）
        
        Returns:
            tuple: (总结结果, 模型信息字典)，失败返回None
        """
        if not self.client:
            print("❌ Azure OpenAI客户端未初始化")
            return None
        
        if not text or len(text.strip()) < 100:
            print("⚠️ 文本内容过短，无法总结")
            return None
        
        try:
            print(f"🤖 正在使用Azure OpenAI总结文档...")
            print(f"   模型: {self.deployment_name}")
            
            # 调用Azure OpenAI API
            response = self.client.chat.completions.create(
                model=self.deployment_name,
                messages=[
                    {"role": "system", "content": self.system_prompt},
                    {"role": "user", "content": f"请总结以下PDF文档的内容（文件名：{filename}）：\n\n{text}"}
                ]
            )
            
            summary = response.choices[0].message.content
            
            # 提取模型信息
            model_info = {
                'model': response.model,
                'prompt_tokens': response.usage.prompt_tokens,
                'completion_tokens': response.usage.completion_tokens,
                'total_tokens': response.usage.total_tokens,
            }
            
            print(f"✅ 总结完成")
            print(f"   字符数: {len(summary)}")
            print(f"   Token使用: {model_info['total_tokens']} (输入:{model_info['prompt_tokens']}, 输出:{model_info['completion_tokens']})")
            
            return summary, model_info
            
        except Exception as e:
            print(f"❌ AI总结失败: {e}")
            return None
    
    def create_email_body(self, filename: str, summary: str, pdf_path: str, model_info: dict = None) -> str:
        """创建邮件正文（HTML格式）
        
        Args:
            filename: PDF文件名
            summary: 总结内容
            pdf_path: PDF文件路径
            model_info: 模型信息字典
        
        Returns:
            str: HTML格式的邮件正文
        """
        file_size = os.path.getsize(pdf_path) / 1024  # KB
        file_size_str = f"{file_size:.1f} KB" if file_size < 1024 else f"{file_size/1024:.1f} MB"
        
        # 构建模型信息HTML
        model_info_html = ""
        if model_info:
            model_info_html = f"""
                <ul>
                    <li><strong>模型:</strong> {model_info.get('model', 'N/A')}</li>
                    <li><strong>Token使用:</strong> {model_info.get('total_tokens', 0):,} 
                        (输入: {model_info.get('prompt_tokens', 0):,}, 输出: {model_info.get('completion_tokens', 0):,})</li>
                </ul>
            """
        
        html = f"""
        <html>
        <head>
            <style>
                body {{ font-family: Arial, sans-serif; line-height: 1.6; color: #333; }}
                .header {{ background-color: #4CAF50; color: white; padding: 20px; text-align: center; }}
                .content {{ padding: 20px; background-color: #f9f9f9; }}
                .summary {{ background-color: white; padding: 15px; border-left: 4px solid #4CAF50; margin: 20px 0; white-space: pre-wrap; }}
                .info {{ color: #666; font-size: 14px; margin-top: 20px; }}
                .footer {{ text-align: center; padding: 10px; color: #999; font-size: 12px; }}
            </style>
        </head>
        <body>
            <div class="header">
                <h1>📄 PDF文档总结</h1>
            </div>
            <div class="content">
                <h2>文档信息</h2>
                <ul>
                    <li><strong>文件名:</strong> {filename}</li>
                    <li><strong>文件大小:</strong> {file_size_str}</li>
                    <li><strong>处理时间:</strong> {time.strftime('%Y-%m-%d %H:%M:%S')}</li>
                </ul>
                
                <h2>总结</h2>
                <div class="summary">
{summary.replace(chr(10), '<br>')}
                </div>
                
                <div class="info">
                    <p>💡 本总结由Azure OpenAI自动生成</p>
                    {model_info_html}
                    <p>📎 原始PDF文件已作为附件发送</p>
                </div>
            </div>
            <div class="footer">
                <p>知识星球 PDF 自动总结系统</p>
            </div>
        </body>
        </html>
        """
        return html
    
    def process_pdf(self, pdf_path: str, send_email: bool = True) -> bool:
        """处理单个PDF文件：提取、总结、发送邮件
        
        Args:
            pdf_path: PDF文件路径
            send_email: 是否发送邮件
        
        Returns:
            bool: 处理成功返回True
        """
        filename = os.path.basename(pdf_path)
        print(f"\n{'='*60}")
        print(f"🔄 处理PDF文件: {filename}")
        print(f"{'='*60}")
        
        # 1. 提取PDF文本
        text = self.extract_text_from_pdf(pdf_path)
        if not text:
            print("❌ PDF文本提取失败，跳过此文件")
            return False
        
        # 2. AI总结
        result = self.summarize_text(text, filename)
        if not result:
            print("❌ AI总结失败，跳过此文件")
            return False
        
        summary, model_info = result
        
        # 3. 发送邮件
        if send_email and self.email_sender and self.receiver_emails:
            subject = self.subject_template.format(filename=filename)
            body = self.create_email_body(filename, summary, pdf_path, model_info)
            
            success = self.email_sender.send_email(
                receiver_emails=self.receiver_emails,
                subject=subject,
                body=body,
                attachments=[pdf_path]
            )
            
            if success:
                print(f"✅ 处理完成并已发送邮件")
                return True
            else:
                print(f"⚠️ 总结完成但邮件发送失败")
                return False
        else:
            print(f"✅ 总结完成（未配置邮件或未开启发送）")
            print(f"\n总结内容:\n{summary}")
            return True
    
    def scan_and_process_pdfs(self, directory: str, send_email: bool = True, 
                             max_files: Optional[int] = None) -> Dict[str, int]:
        """扫描目录并处理所有PDF文件
        
        Args:
            directory: 要扫描的目录
            send_email: 是否发送邮件
            max_files: 最多处理的文件数（None表示不限制）
        
        Returns:
            dict: 处理统计信息
        """
        print(f"\n🔍 扫描目录: {directory}")
        
        # 查找所有PDF文件
        pdf_files = []
        for root, dirs, files in os.walk(directory):
            for file in files:
                if file.lower().endswith('.pdf'):
                    pdf_files.append(os.path.join(root, file))
        
        print(f"📚 找到 {len(pdf_files)} 个PDF文件")
        
        if not pdf_files:
            print("⚠️ 没有找到PDF文件")
            return {'total': 0, 'success': 0, 'failed': 0}
        
        # 限制处理数量
        if max_files:
            pdf_files = pdf_files[:max_files]
            print(f"📋 限制处理前 {max_files} 个文件")
        
        # 处理每个PDF
        stats = {'total': len(pdf_files), 'success': 0, 'failed': 0}
        
        for i, pdf_path in enumerate(pdf_files, 1):
            print(f"\n进度: [{i}/{len(pdf_files)}]")
            
            try:
                success = self.process_pdf(pdf_path, send_email)
                if success:
                    stats['success'] += 1
                else:
                    stats['failed'] += 1
            except Exception as e:
                print(f"❌ 处理出错: {e}")
                stats['failed'] += 1
            
            # 添加延迟，避免API调用过快
            if i < len(pdf_files):
                time.sleep(2)
        
        # 打印统计
        print(f"\n{'='*60}")
        print(f"📊 处理完成统计:")
        print(f"   总计: {stats['total']}")
        print(f"   成功: {stats['success']}")
        print(f"   失败: {stats['failed']}")
        print(f"{'='*60}")
        
        return stats


def load_config():
    """加载TOML配置文件"""
    if tomllib is None:
        print("❌ tomllib/tomli库未安装")
        return None

    config_paths = [
        "config.toml",
        "../config.toml",
        "../../config.toml"
    ]

    config_file = None
    for path in config_paths:
        if os.path.exists(path):
            config_file = path
            break

    if config_file is None:
        print("⚠️ 未找到config.toml配置文件")
        return None
    
    try:
        with open(config_file, 'rb') as f:
            config = tomllib.load(f)
        print("✅ 配置文件加载成功")
        return config
    except Exception as e:
        print(f"❌ 加载配置文件出错: {e}")
        return None


def main():
    """主函数"""
    parser = argparse.ArgumentParser(
        description='PDF文档自动总结与邮件发送脚本',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例用法:
  # 处理指定目录的所有PDF文件并发送邮件
  python pdf_summarizer.py -d ./output/databases/12345/downloads
  
  # 只总结不发送邮件
  python pdf_summarizer.py -d ./downloads --no-email
  
  # 限制只处理前5个文件
  python pdf_summarizer.py -d ./downloads -n 5
  
  # 处理单个PDF文件
  python pdf_summarizer.py -f ./test.pdf
        """
    )
    parser.add_argument('-d', '--directory', type=str,
                        help='要扫描的PDF文件目录')
    parser.add_argument('-f', '--file', type=str,
                        help='单个PDF文件路径')
    parser.add_argument('-n', '--max-files', type=int,
                        help='最多处理的文件数量')
    parser.add_argument('--no-email', action='store_true',
                        help='只总结不发送邮件')
    
    args = parser.parse_args()
    
    # 加载配置
    config = load_config()
    if not config:
        print("❌ 无法加载配置文件，请检查config.toml")
        return
    
    # 检查必要配置
    if not config.get('azure_openai', {}).get('api_key'):
        print("❌ 请先在config.toml中配置Azure OpenAI API密钥")
        return
    
    # 创建总结器
    summarizer = PDFSummarizer(config)
    
    # 确定是否发送邮件
    send_email = not args.no_email
    
    # 处理单个文件
    if args.file:
        if not os.path.exists(args.file):
            print(f"❌ 文件不存在: {args.file}")
            return
        summarizer.process_pdf(args.file, send_email)
        return
    
    # 处理目录
    if args.directory:
        if not os.path.exists(args.directory):
            print(f"❌ 目录不存在: {args.directory}")
            return
        summarizer.scan_and_process_pdfs(args.directory, send_email, args.max_files)
        return
    
    # 如果没有指定文件或目录，使用默认下载目录
    default_dir = config.get('download', {}).get('dir', 'downloads')
    if os.path.exists(default_dir):
        print(f"📂 使用默认下载目录: {default_dir}")
        summarizer.scan_and_process_pdfs(default_dir, send_email, args.max_files)
    else:
        print("❌ 请使用 -d 指定目录或 -f 指定文件")
        parser.print_help()


if __name__ == "__main__":
    main()
