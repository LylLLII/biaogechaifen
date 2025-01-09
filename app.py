from flask import Flask, render_template, request, send_file, jsonify
import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import re
from io import BytesIO
import zipfile
import tempfile

app = Flask(__name__)

# 确保上传文件夹存在
UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process_file():
    if 'file' not in request.files:
        return jsonify({'error': '没有上传文件'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': '未选择文件'}), 400

    # 创建临时目录存储处理后的文件
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_input = os.path.join(temp_dir, file.filename)
        file.save(temp_input)

        try:
            output_files = process_excel(temp_input, temp_dir)
            
            memory_file = BytesIO()
            with zipfile.ZipFile(memory_file, 'w') as zf:
                for output_file in output_files:
                    zf.write(output_file, os.path.basename(output_file))
            
            memory_file.seek(0)
            return send_file(
                memory_file,
                mimetype='application/zip',
                as_attachment=True,
                download_name='处理结果.zip'
            )

        except Exception as e:
            return jsonify({'error': str(e)}), 500

def process_excel(input_file, output_dir):
    # 这里需要把原来的Excel处理逻辑完整地复制过来
    # 从原始文件 20241209省平台总表拆分成各家明细表优化后打包.py 中复制处理逻辑
    output_files = []
    
    # 复制原有的列宽定义
    column_widths = {
        '医疗机构编码': 6.25,
        '医疗机构名称': 10,
        '患者姓名': 4,
        '患者性别': 5.8,
        '险种类型': 8,
        '结算日期': 6.6,
        '医保目录名称': 4,
        '规则名称': 9,
        '疑似违规内容': 13.25,
        '疑似违规金额': 6,
        '初审意见': 15,
        '复审意见': 15,
        '终审意见': 15,
        '申诉意见': 6.8,
        '终审结论': 3.4,
        '扣款金额（元）': 5.7,
        '终审时间': 10,
        '二次反馈': 4,
        '备注': 6,
    }

    # 复制原有的列定义
    columns_to_keep = [
        '医疗机构编码', '医疗机构名称', '患者姓名', '患者性别', '险种类型',
        '结算日期', '医保目录名称', '规则名称', '疑似违规内容',
        '疑似违规金额', '初审意见', '申诉意见', '复审意见',
        '终审结论', '终审意见', '扣款金额（元）', '终审时间',
        '二次反馈', '备注'
    ]

    # 在这里复制原有的Excel处理逻辑
    # ... (从原始文件复制处理逻辑)
    
    return output_files

if __name__ == '__main__':
    # 使用生产级别的服务器运行
    from waitress import serve
    serve(app, host='0.0.0.0', port=8080) 