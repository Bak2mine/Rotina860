import pyautogui
import time

screen_width, screen_height = pyautogui.size()
time.sleep(5)
pyautogui.moveTo(screen_width//2, 50)