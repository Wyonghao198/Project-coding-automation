# requset 用于访问客户端发送给服务器的请求数据，获取方法(GET/POSE)
# render_template 用于接收模板，将参数传递给模板引擎Jinja2；制作动态生成网页
# flash 用于闪现消息
# request 用于发送HTTP请求
# redirect 用于生成一个 HTTP 重定向响应
# url_for 用于用于构建 URL
from flask import Flask, request, render_template, flash, redirect, url_for, send_file, make_response
import os
import pandas as pd
from io import BytesIO

app = Flask(__name__)  # flask基本框架
app.secret_key = 'supersecretkey'  # 设置一个密钥用于flash消息
UPLOAD_FOLDER = 'uploads/'  # 指定上传文件的存储目录
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER  # app.config 是一个字典对象，用于存储应用的配置

@app.route('/')
def index():
    return render_template('upload1.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    # 检查一下是否上传了文件，否则返回url
    if 'file' not in request.files:
        flash('No file part')
        return redirect(request.url)
    file = request.files['file']

    # 检查文件名是否为空，否则返回url
    if file.filename == '':
        flash('No selected file')
        return redirect(request.url)

    # 保存文件
    if file:
        # 获取上传文件的原始名
        filename = file.filename
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        print(f"File will be saved to: {filepath}")
        # 在这里，你可以读取上传的文件
        # Excel文件路径
        file_path1 = r'E:\0工作目的\python\项目实例\基础数据库.xlsm'

        # 读取第一个Excel文件，用于映射
        # 这里要自己调整基础数据库的标题行，设置相应的索引列
        df1 = pd.read_excel(
            file_path1,
            sheet_name=0,
            header=0,
            index_col='编号',
            engine='openpyxl'
        )

        # 创建一个从df1的'c'(项目名称)列到'b'（编码)列的映射
        c_to_b_map = df1.set_index('项目名称')['编码'].to_dict()

        # 这里模拟了一个“打开-读取”的过程，但实际上pandas会直接加载文件
        xls = pd.ExcelFile(filepath, engine='openpyxl')
        sheet_names = xls.sheet_names

        # 遍历所有工作表
        updated_sheets = {}
        for sheet_name in sheet_names:
            df2 = pd.read_excel(
                xls,
                sheet_name=sheet_name,
                header=3,
                index_col='编号')

            df2['编码'] = df2['项目名称'].map(c_to_b_map).fillna('需要手动填充')

            # 重置索引，并将'编号'列添加回DataFrame
            df2.reset_index(inplace=True)  # 重置索引
            df2.rename(columns={'index': '编号'}, inplace=True)  # 将默认索引列名'index'更改为'编号'

            # 将更新后的DataFrame保存到字典中
            updated_sheets[sheet_name] = df2

        # 使用 BytesIO 创建一个内存中的文件对象
        output = BytesIO()

        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for sheet_name, df in updated_sheets.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)

                # 将内存中的文件指针重置到开始位置
        output.seek(0)

        # 返回文件给客户端作为下载
        response = make_response(
            send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'))
        response.headers['Content-Disposition'] = 'attachment; filename="updated_database.xlsx"'
        return response

if __name__ == '__main__':
    if not os.path.exists(UPLOAD_FOLDER):
        os.makedirs(UPLOAD_FOLDER)
    app.run(host='0.0.0.0', port=1500, debug=True)

