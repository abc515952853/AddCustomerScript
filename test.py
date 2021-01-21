import pyautogui
import time

x,y = pyautogui.position()
time.sleep(5)
while True:
    y = y+5
    pyautogui.mouseDown(button='left')
    pyautogui.moveTo(x,y)
    pyautogui.mouseUp()
