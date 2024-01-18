import os
import pathlib
import py7zr
import rarfile
import shutil
import tempfile
import zipfile

# 文件名计数器，在一个压缩包里可能有多个相同后缀的文档，解压后形成 学号-姓名.ext 学号-姓名-1.ext 等等
name_counter = {}


def main():
    # 请修改下面参数
    extract([], "2.zip", [".doc", ".docx"], "实习报告")


def extract(layers=[], filename=None, exts=[], dst=None, delete=False):
    """从指定的压缩文件中提取指定的文件类型。
    layers: list[str] 当前所处的层，即压缩包文件名列表
    filename: str 压缩包名
    exts: list[str] 文件类型列表
    dst: str 目标路径
    delete: 提取完毕后是否删除压缩包
    """

    print(filename)

    # 创建路径
    if dst and (not os.path.exists(dst)):
        os.mkdir(dst)

    # 添加到新的层，防止嵌套干扰
    newlayers = []
    for name in layers:
        newlayers.append(name)
    newlayers.append(filename)
    if filename.endswith(".7z"):
        extract_7z(newlayers, filename, exts, dst, delete)
    if filename.endswith(".rar"):
        extract_rar(newlayers, filename, exts, dst, delete)
    if filename.endswith(".zip"):
        extract_zip(newlayers, filename, exts, dst, delete)
    if delete:
        os.remove(filename)


# 获取文件名，如："/path/test.txt" => "test.txt"
def get_name(filename):
    return pathlib.Path(filename).name


# 获取文件名主干，如："/path/test.txt" => "test"
def get_stem(filename):
    return pathlib.Path(filename).stem


# 获取文件名后缀，如："/path/test.txt" => ".txt"
def get_suffix(filename):
    return pathlib.Path(filename).suffix


# 处理文件列表，列表中所有文件已经解压
def process_files(temppath, layers, files, exts, dst, delete):
    for file in files:
        src = os.path.join(temppath, file)
        ext = get_suffix(file)
        if ext in exts:
            newname = get_name(file)
            if len(layers) >= 2:
                stem = get_stem(layers[1])
                newname = stem + ext
                if newname in name_counter:
                    count = name_counter[newname]
                    newname = stem + "-" + str(count) + ext
                    name_counter[newname] = count + 1
                else:
                    name_counter[newname] = 1
            if dst:
                shutil.copyfile(src, os.path.join(dst, newname))
            else:
                shutil.copyfile(src, newname)
            os.remove(src)
        if ext in [".7z", ".rar", ".zip"]:
            extract(layers, src, exts, dst, True)


def extract_7z(layers, filename, exts, dst, delete):
    files = []
    temppath = tempfile.TemporaryDirectory().name
    with py7zr.SevenZipFile(filename) as sz:
        for file in sz.getnames():
            ext = get_suffix(file)
            if ext in exts + [".7z", ".rar", ".zip"]:
                files.append(file)
        sz.extract(path=temppath, targets=files)
    process_files(temppath, layers, files, exts, dst, delete)


def extract_rar(layers, filename, exts, dst, delete):
    files = []
    temppath = tempfile.TemporaryDirectory().name
    with rarfile.RarFile(filename) as rf:
        for file in rf.namelist():
            ext = get_suffix(file)
            if ext in exts + [".7z", ".rar", ".zip"]:
                files.append(file)
                rf.extract(file, temppath)
    process_files(temppath, layers, files, exts, dst, delete)


def extract_zip(layers, filename, exts, dst, delete):
    files = []
    temppath = tempfile.TemporaryDirectory().name
    with zipfile.ZipFile(filename, allowZip64=True, metadata_encoding="gbk") as zf:
        for file in zf.namelist():
            ext = get_suffix(file)
            if ext in exts + [".7z", ".rar", ".zip"]:
                files.append(file)
                zf.extract(file, temppath)
    process_files(temppath, layers, files, exts, dst, delete)


if __name__ == "__main__":
    main()
