import pyautogui
import time

# x,y = pyautogui.position()
# time.sleep(5)
# while True:
#     y = y+5
#     pyautogui.mouseDown(button='left')
#     pyautogui.moveTo(x,y)
#     pyautogui.mouseUp()


num = [50,100,150,200,250,300,350,400,450]
i = 1
while True:
    if i in num:
        print("我是"+str(i))
    i= i+1
