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

game_x = game_pos[0] + 14
game_y = game_pos[1] + 180
all_square = []
#游戏内容方块横纵个数 19*11  ,qq连连看的方块大小32*36 
for x in range(0,19):
    for y in range(0,11):
        square = screen_image[game_y +y*36:game_y + (y+1)*36, game_x + x*32:game_x + (x+1)*32]
        all_square.append(square)
all_square_list = list(map(lambda square : square[5:26,5:30],all_square))