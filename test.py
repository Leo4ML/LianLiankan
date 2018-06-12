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
#游戏内容方块横纵个数 19*11  ,qq连连看的方块大小31*35 ，每个方框边缘像素宽2
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
all_square_list = list(map(lambda square : square[5:26,5:30],all_square))

def isImageExist(img, img_list):
    for existed_img in img_list:
        b = np.subtract(existed_img,img)
        if not np.any(b):
            return True
        else:
            continue
    return False