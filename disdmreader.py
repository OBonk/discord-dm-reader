import win32com.client as comctl
import win32api
from win32con import *
from pywinauto import Desktop
from time import sleep
from win32api import GetMonitorInfo, MonitorFromPoint

monitor_info = GetMonitorInfo(MonitorFromPoint((0,0)))
#monitor_area = monitor_info.get("Monitor")
work_area = monitor_info.get("Work")
y = int(work_area[3]/2)
x = int((work_area[2]*3)/4)

def findDiscord(l):
    for win in l:
        if "discord" in win.window_text().lower():
            return win.window_text()
    return None

windows = Desktop(backend="uia").windows()
window = findDiscord(windows)
if window == None:
    print("Discord not open!")
    quit()
else:
    print(f"found {window}")
    input("Press enter to continue...")
    wsh = comctl.Dispatch("WScript.Shell")
    running = True
    while running:
        num = int(input("How many dms would you like to read? (0 to exit) "))
        window = findDiscord(windows)
        wsh.AppActivate(window)
        if num != 0:
            for i in range(num):
                wsh.SendKeys("%{DOWN}")
                sleep(0.7)
                win32api.mouse_event(MOUSEEVENTF_WHEEL, x, y, -100, 0)
        else:
            running = False
# Google Chrome window title

