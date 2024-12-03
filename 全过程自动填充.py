# requset 用于访问客户端发送给服务器的请求数据，获取方法(GET/POSE)
# render_template 用于接收模板，将参数传递给模板引擎Jinja2；制作动态生成网页
# flash 用于闪现消息
# request 用于发送HTTP请求
# redirect 用于生成一个 HTTP 重定向响应
# url_for 用于用于构建 URL
# send_file 将文件发送给客户端
# make_response 创建一个响应对象
# import os 处理文件和目录
# BytesIO 操作内存中的数据，临时数据

from flask import (Flask, request, render_template, flash, redirect, url_for,
                   send_file, make_response, send_from_directory)
import os
import pandas as pd
import numpy as np
from io import BytesIO

app = Flask(__name__)  # flask基本框架
app.secret_key = 'supersecretkey@157'  # 设置一个密钥用于flash消息
UPLOAD_FOLDER = 'uploads/'  # 指定上传文件的存储目录
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER  # app.config 是一个字典对象，用于存储应用的配置
app.config['UPLOAD_FOLDER2'] = 'static/pdfs/'  # 存放pdf阅读文件

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/page1')
def page1():
    return render_template('upload1.html')

@app.route('/page2')
def page2():
    return render_template('upload2.html')

@app.route('/page3')
def page3():
    redirect_url = url_for('uploaded_file3', filename='项目编码自动排序web应用使用手册.pdf')
    return redirect(redirect_url)
# 路由来提供PDF、xlsx文件
@app.route('/pdfs/<filename>')
def uploaded_file3(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER2'], filename, as_attachment=False)

@app.route('/upload1', methods=['POST'])
def upload_file1():
    # 检查一下是否上传了文件，否则返回url
    if 'file' not in request.files:
        flash('No file part')
        return redirect(request.url)
    file = request.files['file']

    # 检查文件名是否为空，否则返回url
    if file.filename == '':
        flash('No selected file')
        return redirect(request.url)

    # 获取上传文件的原始名
    filename = file.filename
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(filepath)
    print(f"File will be saved to: {filepath}")
    # 在这里，你可以读取上传的文件
    # Excel文件路径
    file_path1 = r'E:\0工作目的\python\项目实例\基础数据库.xlsm'

    # 读取Excel文件中的所有工作表
    xls1 = pd.read_excel(file_path1, sheet_name=None, engine='openpyxl')
    # 创建一个空的列表来存储所有工作表的DataFrame
    dfs = []

    # 遍历所有工作表，并将它们的数据添加到列表中
    for sheet_name, df in xls1.items():
        # 由于所有工作表的结构相同，我们可以直接添加它们到列表中
        dfs.append(df)

    # 使用pd.concat合并所有DataFrame，并忽略索引（因为我们可能不需要保留原始的工作表索引）
    df1 = pd.concat(dfs, ignore_index=True)
    df1.set_index('编号', inplace=True)

    # 创建一个从df1的'c'(项目名称)列到'b'（编码)列的映射
    c_to_b_map = df1.set_index('项目名称')['编码'].to_dict()

    # 这里模拟了一个“打开-读取”的过程，但实际上pandas会直接加载文件
    xls2 = pd.ExcelFile(filepath, engine='openpyxl')
    sheet_names = xls2.sheet_names

    # 遍历所有工作表
    updated_sheets = {}
    for sheet_name in sheet_names:
        df2 = pd.read_excel(
            xls2,
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
    response.headers['Content-Disposition'] = 'attachment; filename="updated_ZDTC.xlsx"'
    return response

@app.route('/upload2', methods=['POST'])
def upload_file2():
    # 检查一下是否上传了文件，否则返回url
    if 'file' not in request.files:
        flash('No file part')
        return redirect(request.url)
    file = request.files['file']

    # 检查文件名是否为空，否则返回url
    if file.filename == '':
        flash('No selected file')
        return redirect(request.url)

    # 获取上传文件的原始名
    filename = file.filename
    filepath3 = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(filepath3)
    print(f"File will be saved to: {filepath3}")
    # 在这里，你可以读取上传的文件
    # Excel文件路径

    # 这里模拟了一个“打开-读取”的过程，但实际上pandas会直接加载文件
    xls = pd.ExcelFile(filepath3, engine='openpyxl')
    sheets_dict = {sheet_name: pd.read_excel(xls, sheet_name=sheet_name, header=0) for sheet_name in
                   xls.sheet_names}
    frames = []

    # 遍历工作表字典，将每个工作表的数据添加到列表中
    # 变量名.append(数据)，列表结尾追加数据【追加单个数据】
    for sheet_name, df in sheets_dict.items():
        # 添加一个名为'工作表'的列，用于标识数据来自哪个工作表
        df['工作表'] = sheet_name
        frames.append(df)

    # 使用pd.concat()合并所有DataFrame
    # ignore_index=True创建一个新的、连续的整数索引
    combined_df = pd.concat(frames, ignore_index=True)

    # has_nulls = combined_df['编码'].isna().any()
    # print(f"编码列中有空值吗? {has_nulls}")

    # 如果有空值，并且想要将它们替换为一个特定值（比如'未知'或0），可以这样做：
    # 注意：选择替换值时要考虑数据类型和后续的数据处理需求
    replacement_value = '未知'  # 或者选择一个数字，如果编码列是数值类型的话
    combined_df['编码'] = combined_df['编码'].fillna(replacement_value)

    # 检查'编码'列的类型
    print("编码列的类型:", combined_df['编码'].dtype)

    # 用户输入起始计数值
    start_value = request.form.get('start_value')
    try:
        start_value = int(start_value)
    except ValueError:
        return "请输入一个有效的整数！"

    print("用户输入的起始计数值类型:", type(start_value))  # 检查输入值的类型

    # 为每个编码组内的行生成序号，从用户指定的值开始
    # df.groupby 按照'编码'来进行分组，然后用cumcount()对每个分组内部对元素进行编号
    combined_df['序号'] = combined_df.groupby('编码').cumcount() + start_value

    # 将序号格式化为三位字符串（如'001', '002'等）
    # 利用apply(应用一个函数)和lambda，将整数的序号转化为f-string
    combined_df['序号_格式化'] = combined_df['序号'].apply(lambda x: f'{x:03d}')

    # 转化字符串编码类型
    combined_df['编码'] = combined_df['编码'].astype(str)

    # 将原编码和序号合并，创建新的编码列
    combined_df['新编码'] = combined_df['编码'] + combined_df['序号_格式化']

    # 如果你确实需要更新原始编码列
    combined_df['编码'] = combined_df['新编码']

    # 删除辅助列（如果不再需要）
    combined_df.drop(['序号', '序号_格式化', '新编码'], axis=1, inplace=True)

    # 将含'/'整个字符串替换为单个'/'
    combined_df['编码'] = combined_df['编码'].apply(lambda x: '/' if '/' in x else x)

    # 使用apply()函数和lambda表达式，将编码列中包含'未知'两个字的字符串替换为NaN
    combined_df['编码'] = combined_df['编码'].apply(lambda x: np.nan if '未知' in x else x)

    # 现在，我们需要根据'工作表'列将combined_df拆回到原始的工作表结构中
    # 首先，创建一个空的字典来存储拆分后的DataFrame
    split_sheets_dict = {}

    # 遍历所有工作表名，并筛选出属于该工作表的数据
    for sheet_name in sheets_dict.keys():
        split_sheets_dict[sheet_name] = combined_df[combined_df['工作表'] == sheet_name].drop('工作表', axis=1)

    print(split_sheets_dict)
    # 使用 BytesIO 创建一个内存中的文件对象
    output = BytesIO()

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, df in split_sheets_dict.items():
            # 如果需要，可以在这里对df进行进一步的处理或清理
            df.to_excel(writer, sheet_name=sheet_name, index=False)

            # 将内存中的文件指针重置到开始位置
    output.seek(0)

    # 返回文件给客户端作为下载
    response = make_response(
        send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'))
    response.headers['Content-Disposition'] = 'attachment; filename="updated_ZDPX.xlsx"'
    return response

if __name__ == '__main__':
    if not os.path.exists(UPLOAD_FOLDER):
        os.makedirs(UPLOAD_FOLDER)
    app.run(host='0.0.0.0', port=1500, debug=True)

