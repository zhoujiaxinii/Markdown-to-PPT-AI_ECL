from flask import Flask, request, jsonify
import os
import re
from src.utils.md_parser import MDParser
from src.utils.ppt_generator import PPTGenerator

app = Flask(__name__)

@app.route('/api/convert', methods=['POST'])
def convert():
    try:
        # 获取请求参数
        md_file = request.files.get('md_file')
        ppt_template = request.files.get('ppt_template')
        
        if not md_file or not ppt_template:
            return jsonify({'error': 'Missing required files'}), 400
        
        # 保存上传的文件
        md_path = os.path.join('temp', md_file.filename)
        ppt_path = os.path.join('temp', ppt_template.filename)
        
        os.makedirs('temp', exist_ok=True)
        md_file.save(md_path)
        ppt_template.save(ppt_path)
        
        # 解析markdown文件
        parser = MDParser(md_path)
        md_content = parser.parse()
        
        # 生成PPT
        generator = PPTGenerator(ppt_path, md_content)
        output_path = generator.generate()
        
        return jsonify({'success': True, 'output_path': output_path})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)