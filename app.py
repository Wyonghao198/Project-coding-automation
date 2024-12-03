from flask import Flask, request, render_template, flash, redirect, url_for
import os
import pandas as pd

app = Flask(__name__)
app.secret_key = 'supersecretkey'  # 用于flash消息
UPLOAD_FOLDER = 'uploads/'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER


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
        #
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
        with pd.ExcelWriter(r'E:\0工作目的\python\项目实例\测试数据库18.xlsx', engine='openpyxl') as writer:
            for sheet_name, df in updated_sheets.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        print(updated_sheets)
        # 返回成功消息
        return f'File {filename} uploaded and read successfully'
    return 'Failed to upload file'

if __name__ == '__main__':
    if not os.path.exists(UPLOAD_FOLDER):
        os.makedirs(UPLOAD_FOLDER)
    app.run(host='0.0.0.0', port=1500, debug=True)

