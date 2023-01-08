import winsound
import cv2
import time
from cadtest3 import vtPnt,vtInt,vtFloat,vtObject,vtVariant
import win32com.client

wincad = win32com.client.Dispatch("AutoCAD.Application")
doc1 = wincad.ActiveDocument
mp1 = doc1.ModelSpace
mt1_insertpt1=vtFloat([0,0,0])

show_heigth = 30
show_width = 80

ascii_char = list("$@B%8&WM#*oahkbdpqwmZO0QLCJUYXzcvunxrjft/\|()1{}[]?-_+~<>i!lI;:,'. ")
# 生成一个ascii字符列表
char_len = len(ascii_char)
#取列表长度
vc = cv2.VideoCapture("C:/Users/Puck/Downloads/BadApple!!【1080P原版】.mp4")  # 加载视频

if vc.isOpened():  # 判断是否正常打开
    rval,frame = vc.read()
else:
    rval = False

frame_count = 0
outputList = []  # 初始化输出列表
while rval:  # 循环读取视频帧
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)  # 使用opencv转化成灰度图
    gray = cv2.resize(gray, (show_width, show_heigth))  # resize灰度图
    text = ""
    for pixel_line in gray:
        for pixel in pixel_line:  # 字符串拼接
            text += ascii_char[int(pixel / 256 * char_len)]
        text += "\n"
    outputList.append(text)
    frame_count = frame_count + 1
    if frame_count % 100 == 0:
        print("已处理" + str(frame_count) + "帧")
    rval, frame = vc.read()
print("处理完毕")
#outputList=outputList[10:]  #跳过视频开头空白帧

winsound.PlaySound("C:/Users/Puck/Downloads/BadApple.wav", flags=1)  #播放BGM，异步播放
for frame in outputList:
    mt1=mp1.AddMtext(mt1_insertpt1,20,frame)    #方法可查阅
    wincad.Update()
    mt1.Delete()
    time.sleep(0.01223)   #控制帧数，需要调试确定




