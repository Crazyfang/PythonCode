#!/usr/bin/env python
# encoding: utf-8

import shutil
import os.path
import zipfile
import datetime
from PIL import Image, ImageFont, ImageDraw
import tkinter
import tkinter.filedialog


def zip_files(fpath, targetpath):
    """
    #将压缩包解压到指定目录下
    :param fpath:压缩文件
    :param targetpath:压缩目标路径
    :return
    """
    try:
        f = zipfile.ZipFile(fpath, "r")
        for file in f.namelist():
            f.extract(file, targetpath)
        print("解压成功")
        return True
    except Exception as e:
        print("解压出错！", e)
        return False


def ergodic_files(zippath, targetpath):
    """
    遍历指定目录，显示目录下的所有文件名
    :param zippath:解压路径
    :param targetpath:新图片生成的路径
    :return :
    """
    filenames = os.listdir(zippath)
    for filename in filenames:
        composite_image(zippath, filename, targetpath)
    return None


def composite_image(filepath, filename, targetpath):
    """
    对图片进行组合
    :param filepath: 读取图片的路径
    :param filename: 图片的名称
    :param targetpath: 新图片生成的路径
    :return None:
    """

    # 定义背景宽度像素
    width = 1183
    # 定义背景高度像素
    #high = 1676
    # 定义粘贴后的二维码大小
    ewmwidth = 680
    ewmhigh = 680
    # 二维码粘贴高度定位为580，小牌下移0；大牌下移221
    gd = 580
    tz = 221
    if v.get() == 1:
        tz = 0
    # 定义字体大小
    zt = 88

    # 打开底版图片
    im = Image.open('./template.jpg')
    # 打开二维码图片
    imin = Image.open(filepath + filename)
    # 缩放
    imin = imin.resize((ewmwidth, ewmhigh))
    # 粘贴图片
    im.paste(imin, (int((width-ewmwidth)/2), gd + tz))

    name = readName(filename)

    length = len(name)

    # 使用自定义的字体，第二个参数表示字符大小
    font = ImageFont.truetype('./HeiTi.TTF', zt)
    # 在图片上添加文字
    draw = ImageDraw.Draw(im)

    # 限制输入16个字符
    # 限制每行字数为8，即8个字换行
    hangzishu = 8
    if length <= hangzishu:  # 字符串很长的情况下
        an = (width - length * zt) / 2  # 判断字符串到图片左侧的距离
        draw.text((an, 1285 + tz), name, fill=(0, 0, 0), font=font)  # 文字写入
    elif (length > hangzishu) and (length <= hangzishu * 2):
        an1 = (width - hangzishu * zt) / 2  # 第一行
        an2 = (width - (length - hangzishu) * zt) / 2  # 第二行
        a1, a2 = cutName(name, hangzishu)
        draw.text((an1, 1252 + tz), a1, fill=(0, 0, 0), font=font)
        draw.text((an2, 1349 + tz), a2, fill=(0, 0, 0), font=font)
    else:
        # 人为控制字符长度
        lb2["text"] = "字数超限！"

    im.save(os.path.join(targetpath, name) + ".jpg", 'JPEG')


# 大于10个字符时分割中文，以1-8为一行，9开始为第二行
def cutName(name, hangzishu):
    return (name[:hangzishu], name[hangzishu:])


# 读取"_"前的名称
def readName(filename):
    return filename.split("_")[0]


# 创建目录
def mkdir(path):
    isExists = os.path.exists(path)
    # 判断结果
    if not isExists:
        # 如果不存在则创建目录
        # 创建目录操作函数
        os.makedirs(path)


# 删除解压后的临时文件夹中的内容
def del_dir(path):
    isExists = os.path.exists(path)
    # 判断结果
    if isExists:
        shutil.rmtree(path)


# 生成当前日期的字符串，作为生成文件夹
def mkdate():
    date = datetime.datetime.now()
    return date.strftime("%Y%m%d")


# 通过选择生成压缩文件的路径+名称
def choose_zipfile():

    filename = tkinter.filedialog.askopenfilename()

    if filename != '':
        lb2.config(text="您选择的文件是：" + filename)
        convert(filename)
    else:
        lb2.config(text="您没有选择任何文件")

def an_garcode(dir_names):

    for temp_name in os.listdir(dir_names):
        try:
            # 使用cp437对文件名进行解码还原
            new_name = temp_name.encode('cp437')
            # win下一般使用的是gbk编码
            new_name = new_name.decode("gbk")
            # 对乱码的文件名及文件夹名进行重命名
            os.rename(os.path.join(dir_names, temp_name), os.path.join(dir_names, new_name))
            # 传回重新编码的文件名给原文件名
            temp_name = new_name
        except:
            # 如果已被正确识别为utf8编码时则不需再编码
            pass


# 转化文件
def convert(fpath):
    global baseFrame

    zippath = "./tmp/"

    try:
        # 预处理，删除临时文件夹
        del_dir(zippath)
        # 创建临时文件夹
        mkdir(zippath)
        # 将压缩包压缩到指定临时文件夹
        zip_files(fpath, zippath)
        # 中文文件名处理
        # an_garcode(zippath)
        # 通过文件路径，创建生成文件夹
        tempfilename = os.path.split(fpath)[1]
        filename = os.path.splitext(tempfilename)[0]
        targetpath = os.path.join(mkdate(), filename)
        mkdir(targetpath)

        # 在指定路径下读取图片组合，并在指定路径下生成图片
        ergodic_files(zippath, targetpath)

    except Exception as e:
        lb2["text"] = str(e)
    else:
        lb2["text"] = "转换成功"
    finally:
        # 删除临时文件夹
        del_dir(zippath)

if __name__ == "__main__":
    # 启动舞台
    baseFrame = tkinter.Tk()
    baseFrame.title('一码通二维码生成小工具')


    # 新建radiobutton，区分模版大小
    mb = [('大', 0), ('小', 1)]
    v = tkinter.IntVar()
    # 默认选中大
    v.set(0)
    # for循环创建单选框
    i = 0   #列数
    for lan, num in mb:
        tkinter.Radiobutton(baseFrame, text=lan, value=num, variable=v).grid(row=0, column=i)
        i += 1

    # 新建选择文件
    lb1 = tkinter.Label(baseFrame, text="请选择压缩文件：")
    lb1.grid(row=1, column=0, stick=tkinter.W)

    btn1 = tkinter.Button(baseFrame, text="选择", command=choose_zipfile)
    btn1.grid(row=1, column=1, stick=tkinter.E)

    # 提示区域
    lb2 = tkinter.Label(baseFrame, text="")
    lb2.grid(row=2)

    # 启动主Frame
    baseFrame.mainloop()


