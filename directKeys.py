# Original code from CodeBullet 
# https://github.com/Code-Bullet/Piano-Tiles/blob/master/directKeys.py

import ctypes
import time

SendInput = ctypes.windll.user32.SendInput

ESC = 0x01
CTRL = 0x1D
ENTER = 0x1C
N = 0x31
W = 0x11
A = 0x1E
S = 0x1F
D = 0x20
K = 0x25
V = 0x2F

# C struct redefinitions
PUL = ctypes.POINTER(ctypes.c_ulong)


class KeyBdInput(ctypes.Structure):
    _fields_ = [("wVk", ctypes.c_ushort),
                ("wScan", ctypes.c_ushort),
                ("dwFlags", ctypes.c_ulong),
                ("time", ctypes.c_ulong),
                ("dwExtraInfo", PUL)]


class HardwareInput(ctypes.Structure):
    _fields_ = [("uMsg", ctypes.c_ulong),
                ("wParamL", ctypes.c_short),
                ("wParamH", ctypes.c_ushort)]


class MouseInput(ctypes.Structure):
    _fields_ = [("dx", ctypes.c_long),
                ("dy", ctypes.c_long),
                ("mouseData", ctypes.c_ulong),
                ("dwFlags", ctypes.c_ulong),
                ("time", ctypes.c_ulong),
                ("dwExtraInfo", PUL)]


class Input_I(ctypes.Union):
    _fields_ = [("ki", KeyBdInput),
                ("mi", MouseInput),
                ("hi", HardwareInput)]


class Input(ctypes.Structure):
    _fields_ = [("type", ctypes.c_ulong),
                ("ii", Input_I)]


from ctypes import windll, Structure, c_long, byref


class POINT(Structure):
    _fields_ = [("x", c_long), ("y", c_long)]


def queryMousePosition():
    pt = POINT()
    windll.user32.GetCursorPos(byref(pt))
    return pt
    # return { "x": pt.x, "y": pt.y}/


def click(x, y):
    # convert to ctypes pixels
    # x = int(x * 0.666)
    # y = int(y * 0.666)
    ctypes.windll.user32.SetCursorPos(x, y)
    ctypes.windll.user32.mouse_event(2, 0, 0, 0, 0)  # left down
    ctypes.windll.user32.mouse_event(4, 0, 0, 0, 0)  # left up

def rightClick(x, y):
    # convert to ctypes pixels
    # x = int(x * 0.666)
    # y = int(y * 0.666)
    ctypes.windll.user32.SetCursorPos(x, y)
    ctypes.windll.user32.mouse_event(8, 0, 0, 0, 0)  # right down
    ctypes.windll.user32.mouse_event(16, 0, 0, 0, 0)  # right up



def moveMouseTo(x, y):
    # convert to ctypes pixels
    # x = int(x * 0.666)
    # y = int(y * 0.666)
    ctypes.windll.user32.SetCursorPos(x, y)
    # ctypes.windll.user32.mouse_event(2, 0, 0, 0, 0)  # left down
    # ctypes.windll.user32.mouse_event(4, 0, 0, 0, 0)  # left up


def PressKey(hexKeyCode):
    extra = ctypes.c_ulong(0)
    ii_ = Input_I()
    ii_.ki = KeyBdInput(0, hexKeyCode, 0x0008, 0, ctypes.pointer(extra))
    x = Input(ctypes.c_ulong(1), ii_)
    ctypes.windll.user32.SendInput(1, ctypes.pointer(x), ctypes.sizeof(x))


def ReleaseKey(hexKeyCode):
    extra = ctypes.c_ulong(0)
    ii_ = Input_I()
    ii_.ki = KeyBdInput(0, hexKeyCode, 0x0008 | 0x0002, 0, ctypes.pointer(extra))
    x = Input(ctypes.c_ulong(1), ii_)
    ctypes.windll.user32.SendInput(1, ctypes.pointer(x), ctypes.sizeof(x))

def isKeyPressed(hexKeyCode):
    print (ctypes.windll.user32.GetKeyState(hexKeyCode))