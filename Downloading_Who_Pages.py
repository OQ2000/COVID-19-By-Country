from pynput.keyboard import Key, Controller, KeyCode, Listener
import clipboard
from time import sleep
import os
import sys
import random
from pyautogui import press, hotkey, typewrite, keyDown, keyUp

sleep(2)

def start_Program():
    i = 2
    x = 0
    slp = 2
    while x in range(61):
        press('enter')
        sleep(slp)
        hotkey('ctrl','s')
        sleep(1)
        press('enter')
        sleep(slp)
        hotkey('ctrl','w')
        sleep(slp)
        hotkey('shift','tab')
        i = i + 1
        x = x + 1

start_Program()