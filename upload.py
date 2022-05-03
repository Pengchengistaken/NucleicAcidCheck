# Python3脚本，不适用于Python2
# !/usr/bin/envpython
# coding=utf-8
from bottle import route, run, template, request, static_file
import os
import time

# 此处需改为你的目录地址
import main

xlsx_path = 'E:\\workspace\\NucleicAcidCheck'  # 定义上传文件的保存路径

# 此处可扩充为完整HTML
uploadPage = '''
    <body id="tinymce" class="mce-content-body " data-id="content" contenteditable="true" spellcheck="false">
        <h1> 注意事项</h1>
        <h3> 
            <ol>
                <li>上传文件命名格式：班级名-20220504-20220504.xlsx，比如<span style="color: rgb(224, 62, 45);" data-mce-style="color: #e03e2d;">503-20220504-20220504.xlsx。</span></li>
                <li>上传文件较大，以及后台识别图像较久，点击上传后，请耐心等待（图片太多可能需要二三十分钟）返回<span style="color: rgb(224, 62, 45);" data-mce-style="color: #e03e2d;">下载文件</span>。</li>
                <li>出现下载文件即可点击下载查看识别后的结果。</li><li>识别结果仅供参考。</li>
            </ol>
            <form action="upload" method="POST" enctype="multipart/form-data">
                <input type="file" name="data" />
                <input type="submit" value="上传" />
            </form>
	    </h3>
    </body>

'''


@route('/upload')
def upload():
    return uploadPage


@route('/upload', method='POST')
def do_upload():
    upload_file = request.files.get('data')  # 获取上传的文件
    upload_file.save(xlsx_path, overwrite=True)  # overwrite参数是指覆盖同名文件
    if file_filter(upload_file.filename):
        if os.system('python3 main.py %s %s' % ('20220429', upload_file.filename)) == 0:
            output_file = 'output/check_' + upload_file.filename
            return u"<h1>识别成功，请点击<a href='/download/" + output_file + "'>下载文件</a></h1>"
        else:
            return u"<h1>出错了！请检查上传的文件或者联系管理员！</h1>"
    else:
        return u"<h1>出错了！请检查上传的文件或者联系管理员！</h1>"


@route('/download/<filename:path>')
def download(filename):
    return static_file(filename, root=xlsx_path, download=filename)

def file_filter(f):
    if f[-5:] in ['.xlsx']:
        return True
    return False

run(host='0.0.0.0', port=8899, debug=True)