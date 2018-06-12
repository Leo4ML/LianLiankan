# -*- coding: utf-8 -*-
"""
Created on Tue Jun 12 16:30:49 2018

@author: Leo
"""

import metching
import cv2
import numpy as np
import win32api
import win32gui
import win32con
import win32com.client
from PIL import ImageGrab
import time


window = win32gui.FindWindow(None,'QQ游戏 - 连连看角色版')
shell = win32com.client.Dispatch("WScript.Shell")
shell.SendKeys('%')
win32gui.SetForegroundWindow(window)
pos = win32gui.GetWindowRect(window)
game_pos = (pos[0],pos[1])


scim = ImageGrab.grab()
scim.save('screen.png')
screen_image = cv2.imread('screen.png')
#游戏内容方块横纵个数 19*11  ,qq连连看的方块大小31*35 ，每个方框边缘像素宽4
#方块连接区域距离左边距离14像素，顶部距离181像素
game_x = game_pos[0] + 14
game_y = game_pos[1] + 181
all_square = []
#游戏内容方块横纵个数 19*11  ,qq连连看的方块大小31*35 
for x in range(0,19):
    for y in range(0,11):
        square = screen_image[game_y +y*35:game_y + (y+1)*35, game_x + x*31:game_x + (x+1)*31]
        all_square.append(square)
#注意，这里的小方块切除边缘，是先y轴，再x轴的
all_square_list = list(map(lambda square : square[4:32,4:28],all_square))

types=[]
empty_img = cv2.imread('empty.jpg')
types.append(empty_img)
for square in all_square_list:
    if not isImageExist(square, types):
        types.append(square)

def isImageExist(img, img_list):
    for existed_img in img_list:
        b = np.subtract(existed_img,img)
        if not np.any(b):
            return True
        else:
            continue
    return False
#这部分，将所有切割出的19*11个方块分类，加上空白方块有38种，编号0-37.然后将19*11个方块依次比较来
#做编号分类，将所有的方块，编码成数字矩阵。
record = []
line = []
for square in all_square_list:
    num = 0
    for type1 in types:
        res = cv2.subtract(square, type1)
        if not np.any(res):
            line.append(num)
            break
        num+=1
#这里line是二维数组，len（line）是它的行数，既纵列方块个数；
#for square in all_square_list  这里是纵向遍历的，所以后面要transpose，恢复成正常视觉顺序。
    if len(line) == 11:
        record.append(line)
        line=[]
print(record)

result = np.transpose(record)
#    return record


def autoRelease(result,game_x,game_y):
    for i in range(len(result)):
        for j in range(len(result[0])):
            # 以上两个for循环，定位第一个选中点
            if result[i][j] != 0:
                for m in range(0,len(result)):
                    for n in range(0,len(result[0])):
                        if result[m][n] != 0:
                            # 后两个for循环定位第二个选中点
                            if metching.canConnect(i,j,m,n,result):
                            # 执行消除算法并返回
                                result[i][j] = 0
                                result[m][n] = 0
                                print('可消除点：'+ str(i+1) + ',' + str(j+1) + '和' + str(m+1) + ',' + str(n+1))
                                x1 = game_x + j*31
                                y1 = game_y + i*35
                                x2 = game_x + n*31
                                y2 = game_y + m*35
                                win32api.SetCursorPos((x1 + 15,y1 + 18))
                                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, x1+15, y1+18, 0, 0)
                                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, x1+15, y1+18, 0, 0)
                                time.sleep(0.5)

                                win32api.SetCursorPos((x2 + 15, y2 + 18))
                                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, x2 + 15, y2 + 18, 0, 0)
                                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, x2 + 15, y2 + 18, 0, 0)
                                time.sleep(0.5)
                                return True
    return False




