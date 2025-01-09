import PyInstaller.__main__
import os

# 获取当前目录
current_dir = os.path.dirname(os.path.abspath(__file__))

PyInstaller.__main__.run([
    'app.py',
    '--onefile',
    '--windowed',
    '--icon=app.ico',
    '--add-data=templates;templates',
    '--hidden-import=pandas',
    '--hidden-import=openpyxl',
    '--name=Excel处理工具'
]) 