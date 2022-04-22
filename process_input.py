import base64
import re
import xml.dom.minidom as xmldom
import os
import zipfile
import shutil
import xlrd


def isfile_exist(file_path):
    if not os.path.isfile(file_path):
        return False
    return True


def copy_change_file_name(file_path, new_type='.zip'):
    if not isfile_exist(file_path):
        return ''
    extend = os.path.splitext(file_path)[1]
    if extend != '.xlsx' and extend != '.xls':
        print('不是excel文件')
        return False

    file_name = os.path.basename(file_path)
    new_name = str(file_name.split('.')[0]) + new_type
    dir_path = os.path.dirname(file_path)  # 文件所在目录
    new_path = os.path.join(dir_path, new_name)  # 新文件路径

    if os.path.exists(new_path):
        os.remove(new_path)
    shutil.copyfile(file_path, new_path)
    return new_path  # 新的文件路径, zip文件


def unzip_file(zipfile_path):
    if not isfile_exist(zipfile_path):
        return False
    if os.path.splitext(zipfile_path)[1] != '.zip':
        print('不是zip文件')
        return False
    file_zip = zipfile.ZipFile(zipfile_path, 'r')
    file_name = os.path.basename(zipfile_path)  # 获取文件名
    zip_dir = os.path.join(os.path.dirname(zipfile_path), str(file_name.split('.')[0]))  # 文件所在目录
    # print("file_zip: " + file_zip.filename)
    # print("file_name: " + file_name)
    # print("zip_dir: " + zip_dir)
    for files in file_zip.namelist():
        file_zip.extract(files, zip_dir)  # 解压到指定目录
    file_zip.close()
    return True


# 读取解压后文件夹，打印图片路径
def read_img(zipfile_path):
    img_dict = dict()
    if not isfile_exist(zipfile_path):
        return False

    dir_path = os.path.dirname(zipfile_path)  # 文件所在目录
    file_name = os.path.basename(zipfile_path)  # 文件名
    pic_dir = 'xl' + os.sep + 'media'  # 解压的图片在media目录
    pic_path = os.path.join(dir_path, str(file_name.split('.')[0]), pic_dir)

    file_list = os.listdir(pic_path)
    for file in file_list:
        filepath = os.path.join(pic_path, file)
        # print(filepath)
        img_index = int(re.findall(r'image(\d+)\.', filepath)[0])
        img_base64 = get_img_base64(img_path=filepath)
        img_dict[img_index] = dict(img_index=img_index, img_path=filepath, img_base64=img_base64)
    return img_dict


def get_img_base64(img_path):
    if not isfile_exist(img_path):
        return ''
    with open(img_path, 'rb') as f:
        base64_date = base64.b64encode(f.read())
        s = base64_date.decode()
        return 'data:image/jpeg;base64,%s' % s


def get_img_pos_info(zip_file_path, img_dict, img_feature):
    os.path.dirname(zip_file_path)
    dir_path = os.path.dirname(zip_file_path)
    file_name = os.path.basename(zip_file_path)
    xml_dir = 'xl' + os.sep + 'drawings' + os.sep + 'drawing1.xml'
    xml_path = os.path.join(dir_path, str(file_name.split('.')[0]), xml_dir)
    image_info_dict = parse_xml(xml_path, img_dict, img_feature)
    return image_info_dict


def get_img_info(excel_file_path, img_feature):
    if img_feature not in ['img_index', 'img_path', 'img_base64']:
        raise Exception("图片返回参数错误", ['img_index', 'img_path', 'img_base64'])
    zip_file_path = copy_change_file_name(excel_file_path)
    if zip_file_path != '':
        if unzip_file(zip_file_path):
            img_dict = read_img(zip_file_path)
            image_info_dict = get_img_pos_info(zip_file_path, img_dict, img_feature)
            return image_info_dict
    return dict()


# 解析xml文件并获取图片位置
def parse_xml(file_name, img_dic, img_feature='img_path'):
    image_info = dict()
    dom_obj = xmldom.parse(file_name)
    element = dom_obj.documentElement

    def _f(subElementObj):
        for anchor in subElementObj:
            xdr_from = anchor.getElementsByTagName('xdr:from')[0]
            col = xdr_from.childNodes[0].firstChild.data  # 获取标签间的数据
            row = xdr_from.childNodes[2].firstChild.data
            embed = anchor.getElementsByTagName('xdr:pic')[0].getElementsByTagName('xdr:blipFill')[0]\
                .getElementsByTagName('a:blip')[0].getAttribute('r:embed')  # 获取属性
            image_info[int(row), int(col)] = img_dic.get(int(embed.replace('rId', '')), {}).get(img_feature)

    sub_twoCellAnchor = element.getElementsByTagName('xdr:twoCellAnchor')
    sub_oneCellAnchor = element.getElementsByTagName('xdr:oneCellAnchor')
    _f(sub_twoCellAnchor)
    _f(sub_oneCellAnchor)
    return image_info


def read_excel_info(file_path, img_col_index, img_feature='img_path'):
    """
    读取包含图片的excel，并返回列表
    :param file_path:
    :param img_col_index:
    :param img_feature:
    :return:
    """
    img_info_dict = get_img_info(file_path, img_feature)
    book = xlrd.open_workbook(file_path)
    sheet = book.sheet_by_index(0)
    head = dict()
    for i, v in enumerate(sheet.row(0)):
        head[i] = v.value
    info_list = []
    for row_num in range(sheet.nrows):
        d = dict()
        for col_num in range(sheet.ncols):
            if row_num == 0:
                continue
            if 'empty:' in sheet.cell(row_num, col_num).__str__():
                if col_num in img_col_index:
                    d[head[col_num]] = img_info_dict.get((row_num, col_num))
                else:
                    d[head[col_num]] = sheet.cell(row_num, col_num).value
            else:
                d[head[col_num]] = sheet.cell(row_num, col_num).value
        if d:
            info_list.append(d)
    return info_list
