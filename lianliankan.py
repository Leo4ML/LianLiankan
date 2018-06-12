# -*- coding: utf-8 -*-
"""
Created on Mon Jun 11 15:27:03 2018

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
#from config import *


def getGameWindowPosition():
    window = win32gui.FindWindow(None,'QQ游戏 - 连连看角色版')
    while not window:
        print ('重新定位游戏窗口....')
        window = win32gui.FindWindow(IpClassName=None,IpWindowName='QQ游戏 - 连连看角色版')
#win32的bug，若要啊置顶窗口，必须先得发送个按键指令，这里发送alt，%表示    
    shell = win32com.client.Dispatch("WScript.Shell")
    shell.SendKeys('%')
    win32gui.SetForegroundWindow(window)
    pos = win32gui.GetWindowRect(window)
    print('定位游戏窗口位置：'+str(pos))
    return (pos[0],pos[1])

def getScreenImage():
    print('捕捉屏幕截图')
    scim = ImageGrab.grab()
    scim.save('screen.png')
    return cv2.imread('screen.png')

def getAllSquare(screen_image,game_pos):
    game_x = game_pos[0] + 9
    game_y = game_pos[1] + 180
    all_square = []
#游戏内容方块横纵个数 19*11  ,qq连连看的方块大小36*36 
    for x in range(0,19):
        for y in range(0,11):
            square = screen_image[game_y +y*35:game_y + (y+1)*35, game_x + x*35:game_x + (x+1)*35]
            all_square.append(square)
    return list(map(lambda square : square[5:26,5:30],all_square))
'''
#这里是对小方块的边缘处理，放防止干扰 
    new_all_square = []
    for square in all_square:
        s = square[5:26,5:30]
        new_all_square.append(s)
    return new_all_square
'''
def isImageExist(img, img_list):
    for existed_img in img_list:
        b = np.subtract(existed_img,img)
        if not np.any(b):
            return True
        else:
            continue
    return False

def getAllSquareTypes(all_square):
    types=[]
    empty_img = cv2.imread('empty.png')
    types.append(empty_img)
    for square in all_square:
        if not isImageExist(square, types):
            types.append(square)
    return types

def getAllSquareRecord(all_square_list, types):
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
#这里是方块纵向的个数
        if len(line) == 11:
            record.append(line)
            line=[]
    print(record)
    return record

def autoRelease(result, game_x,game_y):
    for i in range(0,len(result)):
        for j in range(0,len(result[0])):
            if result[i][j] !=0:
                for m in range(0,len(result)):
                    for n in range(0,len(result)):
                        if result[m][n] !=0:
                            if metching.canConnect(i,j,m,n,result):
                                result[i][j] = 0
                                result[m][n] = 0
                                print('可消除点：'+ str(i+1) + ',' + str(j+1) + '和' + str(m+1) + ',' + str(n+1))
                                x1 = game_x + j*36
                                y1 = game_y + i*36
                                x2 = game_x + n*36
                                y2 = game_y + m*36
                                win32api.SetCursorPos((x1 + 15,y1 + 18))
                                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, x1+15, y1+18, 0, 0)
                                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, x1+15, y1+18, 0, 0)
                                time.sleep(1)
                                return True
    return False


def autoRemove(squares,game_pos):
    game_x = game_pos[0] + 100
    game_y = game_pos[1] + 100
    # 判断是否消除完了？如果没有的话，点击重列后继续消除
    for i in range(0,105):
        autoRelease(squares,game_x,game_y)
        
if __name__ == '__main__':
    # 1、定位游戏窗体
    game_pos = getGameWindowPosition()
    time.sleep(1)
    # 2、从屏幕截图一张，通过opencv读取
    screen_image = getScreenImage()
    # 3、图像切片，把截图中的连连看切成一个一个的小方块，保存在一个数组中
    all_square_list = getAllSquare(screen_image,game_pos)
    # 4、切片处理后的图片，相同的作为一种类型，放在数组中。
    types = getAllSquareTypes(all_square_list)
    # 5、将切片处理后的图片，转换成相对应的数字矩阵。注意 拿到的数组是横纵逆向的，转置一下。
    result = np.transpose(getAllSquareRecord(all_square_list,types))
    # 6、执行自动消除
    autoRemove(result,game_pos)
    # 7、消除完成，释放资源。
    cv2.waitKey(0)
    cv2.destroyAllWindows()
        