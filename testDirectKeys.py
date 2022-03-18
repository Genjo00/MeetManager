from directKeys import isKeyPressed
from directKeys import click, queryMousePosition, PressKey, ReleaseKey, ESC
import time

while True:
    #pt = queryMousePosition()
    #print(pt.x, pt.y)
    print(isKeyPressed(0x61))
    time.sleep(1)
