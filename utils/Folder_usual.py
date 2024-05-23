# # -*- coding: utf-8 -*-

import os
import shutil
from fastapi import UploadFile
import zipfile
import os

def creat_folder(folder_name):
    FolderName = os.path.join(os.path.dirname(os.path.dirname(__file__)), folder_name)
    if not os.path.exists(FolderName):
        os.makedirs(FolderName)
        # print(f'{FolderName}创建文件夹成功')
        return FolderName
    else:
        # print(f'{FolderName}文件夹已存在')
        return FolderName


def remove_folder(path):
    try:
        shutil.rmtree(path)
        # print(f'已删除文件夹{path}')
        return True
    except OSError as e:
        # print(f'错误信息：{e}')
        return False

def Save_upload_file(file: UploadFile, folder_path: str, file_name):
    file_extension = os.path.splitext(file.filename)[1]  # 提取文件扩展名
    files_in_folder = os.listdir(folder_path)
    file_count = len(files_in_folder)
    file_path = os.path.join(folder_path, file_name)
    with open(file_path, 'wb') as f:
        f.write(file.file.read())
    return file_path

def zip_unzip(zip_file_name, target_dir):
    # 检查是否为ZIP文件
    is_zip = zipfile.is_zipfile(zip_file_name)
    if is_zip:
        # 创建目标目录（如果不存在）
        os.makedirs(target_dir, exist_ok=True)
        with zipfile.ZipFile(zip_file_name, 'r') as zf:
            for name in zf.infolist():
                try:
                    # 解码文件名
                    file_name_utf_16 = name.filename.encode('cp437').decode('gbk', errors='replace')
                    # 检查文件名是否以点开头，如果是则跳过
                    if not file_name_utf_16.startswith('.'):
                        # 构建文件的完整路径
                        full_file_path = os.path.join(target_dir, file_name_utf_16)
                        # 确保文件所在的目录存在
                        os.makedirs(os.path.dirname(full_file_path), exist_ok=True)
                        # 解压文件
                        with open(full_file_path, 'wb') as output_file:
                            with zf.open(name) as input_file:
                                output_file.write(input_file.read())
                except Exception as e:
                    # print(f"Error processing file {name.filename}: {e}")
                    pass
    os.remove(zip_file_name)
    return target_dir

def zip_folder(folder_path, zip_file_name):
    zip_file = zipfile.ZipFile(zip_file_name, 'w', zipfile.ZIP_DEFLATED)
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            zip_file.write(os.path.join(root, file), file)
    zip_file.close()
    return zip_file_name

def creat_function_folder(request, functionName):
    try:
        username = request.session.get('username')
    except:
        username = 'ceshi'
    save_path = os.path.join(os.path.dirname(__file__), "user_folder", f'{username}', functionName, 'input')
    try:
        os.makedirs(save_path)
    except FileExistsError:
        shutil.rmtree(os.path.dirname(save_path))
        os.makedirs(save_path)
    outputfile = os.path.join(os.path.dirname(save_path), 'output')
    try:
        os.makedirs(outputfile)
    except FileExistsError:
        pass
    return save_path, outputfile


if __name__ == '__main__':
    dir = "static/result/"
    print(creat_folder(dir))
