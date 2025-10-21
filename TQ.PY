import os
import re
import requests
from bs4 import BeautifulSoup
from docx import Document
from docx.shared import Inches
from urllib.parse import urljoin, urlparse
import time
import random
from pathlib import Path  # 用于路径处理

class WebToWordConverter:
    def __init__(self, url, output_dir=None):
        """
        初始化网页转Word转换器
        :param url: 目标网页URL
        :param output_dir: 输出目录（默认：当前目录下的web2word_output）
        """
        self.url = url
        # 设置默认输出目录（当前脚本所在目录）
        self.base_dir = Path(__file__).parent.resolve()
        self.output_dir = Path(output_dir) if output_dir else self.base_dir / "web2word_output"
        self.image_dir = self.output_dir / "images"  # 图片保存目录
        
        # 初始化会话和请求头
        self.session = requests.Session()
        self.headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
            "Accept-Language": "zh-CN,zh;q=0.8,zh-TW;q=0.7,zh-HK;q=0.5,en-US;q=0.3,en;q=0.2",
            "Connection": "keep-alive"
        }
        
        self.document = Document()  # 创建Word文档对象
        self._prepare_directories()  # 准备输出目录

    def _prepare_directories(self):
        """创建输出目录（若不存在）"""
        try:
            self.output_dir.mkdir(parents=True, exist_ok=True)
            self.image_dir.mkdir(parents=True, exist_ok=True)
            print(f"输出目录已准备: {self.output_dir}")
        except PermissionError:
            raise Exception(f"权限不足，无法创建目录: {self.output_dir}")
        except Exception as e:
            raise Exception(f"创建目录失败: {str(e)}")

    def _get_soup(self):
        """获取并解析网页内容"""
        try:
            # 随机延迟避免反爬（1-3秒）
            time.sleep(random.uniform(1, 3))
            response = self.session.get(self.url, headers=self.headers, timeout=15)
            response.raise_for_status()  # 抛出HTTP错误（4xx/5xx）
            
            # 自动处理编码，解决中文乱码
            if response.encoding == "ISO-8859-1":
                response.encoding = response.apparent_encoding
            return BeautifulSoup(response.text, "html.parser")
        except requests.exceptions.HTTPError as e:
            print(f"HTTP错误: {str(e)}")
        except requests.exceptions.Timeout:
            print("请求超时")
        except requests.exceptions.ConnectionError:
            print("网络连接错误")
        except Exception as e:
            print(f"获取网页内容失败: {str(e)}")
        return None

    def _clean_text(self, text):
        """清理文本内容（去除多余空行、空格）"""
        if not text:
            return ""
        # 去除首尾空白和连续空行
        text = re.sub(r'\n\s*\n', '\n\n', text.strip())
        # 去除多余空格（保留单个空格）
        text = re.sub(r' +', ' ', text)
        return text

    def _download_image(self, img_src):
        """下载图片并返回本地保存路径"""
        if not img_src:
            return None
            
        try:
            # 处理相对路径
            img_url = urljoin(self.url, img_src)
            # 解析图片文件名
            parsed_url = urlparse(img_url)
            img_filename = os.path.basename(parsed_url.path)
            
            # 处理无扩展名的图片
            if not os.path.splitext(img_filename)[1].lower() in ['.jpg', '.jpeg', '.png', '.gif', '.bmp']:
                img_filename = f"img_{int(time.time())}.jpg"  # 生成随机文件名
            
            # 保存路径
            img_path = self.image_dir / img_filename
            
            # 若图片已存在，直接返回
            if img_path.exists():
                return img_path
            
            # 下载图片
            time.sleep(random.uniform(0.5, 1.5))
            response = self.session.get(img_url, headers=self.headers, timeout=10, stream=True)
            response.raise_for_status()
            
            # 保存图片
            with open(img_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=1024):
                    if chunk:
                        f.write(chunk)
            
            print(f"已下载图片: {img_filename}")
            return img_path
        except Exception as e:
            print(f"图片下载失败（{img_url}）: {str(e)}")
            return None

    def extract_and_save(self, content_selector=None, doc_filename="网页内容提取.docx"):
        """
        提取网页内容并保存为Word文档
        :param content_selector: CSS选择器，指定内容区域（如"div#content_views"）
        :param doc_filename: 输出Word文件名
        """
        print(f"开始处理: {self.url}")
        
        # 解析网页
        soup = self._get_soup()
        if not soup:
            print("无法继续处理，退出程序")
            return
        
        # 定位内容区域
        if content_selector:
            content_area = soup.select_one(content_selector)
            if not content_area:
                print(f"未找到匹配的内容区域: {content_selector}，将尝试提取整个页面")
                content_area = soup.body
        else:
            content_area = soup.body
        
        if not content_area:
            print("未找到可提取的内容区域")
            return
        
        # 提取标题（优先使用h1标签）
        title_tag = soup.find('h1')
        if title_tag and title_tag.get_text(strip=True):
            title = self._clean_text(title_tag.get_text())
            self.document.add_heading(title, level=0)  # 一级标题
            self.document.add_paragraph()  # 空行分隔

        # 遍历内容区域，提取文字和图片
        for element in content_area.descendants:
            # 忽略空标签和脚本标签
            if not element.name or element.name in ['script', 'style', 'nav', 'footer', 'aside']:
                continue
            
            # 处理标题标签（h2-h6）
            if element.name in ['h2', 'h3', 'h4', 'h5', 'h6']:
                text = self._clean_text(element.get_text())
                if text:
                    level = int(element.name[1])  # h2->2, h3->3...
                    self.document.add_heading(text, level=level)
                    self.document.add_paragraph()  # 空行
            
            # 处理段落标签
            elif element.name == 'p':
                text = self._clean_text(element.get_text())
                if text:
                    self.document.add_paragraph(text)
            
            # 处理图片标签
            elif element.name == 'img':
                img_src = element.get('src') or element.get('data-src')  # 处理延迟加载图片
                img_path = self._download_image(img_src) if img_src else None
                if img_path:
                    try:
                        # 插入图片（限制宽度为6英寸，适应Word页面）
                        self.document.add_picture(str(img_path), width=Inches(6))
                        # 添加图片说明（alt属性）
                        img_alt = self._clean_text(element.get('alt', ''))
                        if img_alt:
                            self.document.add_paragraph(f"图片说明：{img_alt}", style="Caption")
                        self.document.add_paragraph()  # 空行分隔
                    except Exception as e:
                        print(f"插入图片失败: {str(e)}")
        
        # 保存Word文档
        doc_path = self.output_dir / doc_filename
        try:
            self.document.save(doc_path)
            print(f"\n处理完成！Word文档已保存至: {doc_path}")
            print(f"图片保存目录: {self.image_dir}")
        except Exception as e:
            print(f"保存Word文档失败: {str(e)}")

if __name__ == "__main__":
    # 配置参数
    TARGET_URL = ""  # 目标网页URL
    CONTENT_SELECTOR = "div#content_views"  # CSDN文章内容区域选择器
    OUTPUT_DIR = None  # 自定义输出目录，如"/home/文档/网页提取"，None则使用默认路径
    DOC_FILENAME = ""  # 输出Word文件名

    # 执行提取
    try:
        converter = WebToWordConverter(TARGET_URL, output_dir=OUTPUT_DIR)
        converter.extract_and_save(content_selector=CONTENT_SELECTOR, doc_filename=DOC_FILENAME)
    except Exception as e:
        print(f"程序运行出错: {str(e)}")

