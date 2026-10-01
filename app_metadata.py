"""Author and application metadata, shared by GUI and reports."""
from pathlib import Path
import sys

APP_NAME = '文件去敏感化與範本清理工具'
VERSION = '3.1.0'
AUTHOR = 'Arthur Tao'
LICENSE_NAME = 'MIT'
BLOG_URL = 'https://taoyutsun.blogspot.com/'
FACEBOOK_URL = 'https://facebook.com/arthurtaoyutsun'
SOURCE_REPO_URL = 'https://github.com/taoyutsun/Document-Scan-Cleanup'
SOURCE_DIRECTORY = Path(sys.executable).parent if getattr(sys, 'frozen', False) else Path(__file__).resolve().parent
AUTHOR_DESCRIPTION = f'{APP_NAME} 由 {AUTHOR} 設計與維護，採 {LICENSE_NAME} 授權。歡迎使用與分享，並保留原作者與來源資訊。'
