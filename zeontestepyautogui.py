import win32gui
import win32con
import time
import pyautogui
import os

IMAGES = r"C:\Users\andre\Desktop\Zeon"
screen_width, screen_height = pyautogui.size()
time.sleep(5)
bottom_left = (0, screen_height//2, screen_width//2, screen_height//2)
location = pyautogui.locateOnScreen(os.path.join(IMAGES, "btn_save.png"), confidence=0.6, region=bottom_left)
pyautogui.click(location)