import os
from os import path

# 获取文件路径，获取文件名称列表
source = path.normpath(r"C:\Users\m\Desktop\\")
videoList = os.listdir(source)

# 只选择目录下的mkv文件
for Sname in videoList:
    if not Sname.endswith("mkv"):
        videoList.remove(Sname)

# 执行ffmpeg命令
for i in videoList:
    output = i[0:-4]
    cmd = r"ffmpeg -i X:\\xxx\\xxx\\xx\\%s -c:v copy -c:a aac Y:\\yyy\\yyy\\%s.mp4" %(i,output)
    os.system(cmd)
