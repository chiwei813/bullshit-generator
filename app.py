# -*- coding: utf-8 -*-
from flask import Flask, render_template, request, send_file, jsonify
from paper_generator import PaperGenerator
import os
from datetime import datetime

app = Flask(__name__)
generator = PaperGenerator()

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate_paper():
    try:
        data = request.get_json()
        topic = data.get('topic')
        if not topic:
            return jsonify({'error': '請輸入研究主題'}), 400

        # 生成論文
        filepath = generator.generate_paper(topic)  # 現在這會返回檔案路徑
        
        # 使用生成的檔案路徑
        filename = os.path.basename(filepath)
        
        return jsonify({
            'status': 'success',
            'message': '論文生成完成',
            'filename': filename
        })
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/download/<filename>')
def download_file(filename):
    try:
        file_dir = os.path.dirname(os.path.abspath(__file__))
        return send_file(
            os.path.join(file_dir, filename),
            as_attachment=True,
            download_name=filename
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 404

@app.route('/progress')
def get_progress():
    return jsonify({'progress': generator.progress})

if __name__ == '__main__':
    app.run(debug=True)