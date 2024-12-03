from flask import Flask, send_from_directory
import os

app = Flask(__name__)

# 配置上传文件夹路径
app.config['UPLOAD_FOLDER'] = 'uploads/'

@app.route('/uploads/<filename>')
def download_excel(filename):
    print(f"Attempting to download file: {filename}")
    # 使用 send_from_directory 发送文件
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=4500, debug=True)