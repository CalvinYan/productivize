# Original code graciously provided by Stack Overflow user David Hefferman

import csv
import sys
import time
import ctypes
import ctypes.wintypes

EVENT_OBJECT_FOCUS = 0x8005
WINEVENT_OUTOFCONTEXT = 0x0000

time_log = {} # The number of seconds spent on each window, using changes in window focus

num_updates = 0
last_window = '' # The name of the recently unfocused window
last_time = 0 # integer timestamp of the last change in focus

user32 = ctypes.windll.user32
ole32 = ctypes.windll.ole32

ole32.CoInitialize(0)

WinEventProcType = ctypes.WINFUNCTYPE(
    None, 
    ctypes.wintypes.HANDLE,
    ctypes.wintypes.DWORD,
    ctypes.wintypes.HWND,
    ctypes.wintypes.LONG,
    ctypes.wintypes.LONG,
    ctypes.wintypes.DWORD,
    ctypes.wintypes.DWORD
)

def callback(hWinEventHook, event, hwnd, idObject, idChild, dwEventThread, dwmsEventTime):
    global num_updates
    global last_window
    global last_time
    length = user32.GetWindowTextLengthA(hwnd)
    buff = ctypes.create_string_buffer(length + 1)
    user32.GetWindowTextA(hwnd, buff, length + 1)
    name = str(buff.value, 'windows-1252')
    if name != '':
        new_time = int(time.time())
        print(time.strftime('%X') + ' --> ' + name)
        if (last_time != 0):
            time_delta = new_time - last_time
            updateLog(time_log, last_window, time_delta)
            print(str(time_delta) + ' seconds have elapsed')
        last_time = new_time
        last_window = name
        if num_updates == 30: 
            writeData(time_log)
            quit()
        num_updates += 1

def readData(time_log):
    date_string = time.strftime('%d-%m-%Y')
    with open(date_string + '.csv', "r+") as dataFile:
        lines = csv.reader(dataFile)
        for line in lines:
            if len(line) > 0:
                print(line)
                time_log[line[0]] = int(line[1])

# Update the total record of time spent on each application
def writeData(time_log):
    log_list = [[key, val] for key, val in time_log.items()] # Convert dict to list
    log_list = sorted(log_list, key = lambda app:(-int(app[1]), app[0]))
    date_string = time.strftime('%d-%m-%Y')
    with open(date_string + '.csv', 'w+', newline = '') as dataFile:
        out = csv.writer(dataFile)
        out.writerows(log_list)

# Add the number of seconds spent using an application to its total duration, or add it if necessary
def updateLog(time_log, app_name, seconds):
    current_value = time_log[app_name] if app_name in time_log else 0
    time_log[app_name] = current_value + seconds

readData(time_log)

WinEventProc = WinEventProcType(callback)

user32.SetWinEventHook.restype = ctypes.wintypes.HANDLE
hook = user32.SetWinEventHook(
    EVENT_OBJECT_FOCUS,
    EVENT_OBJECT_FOCUS,
    0,
    WinEventProc,
    0,
    0,
    WINEVENT_OUTOFCONTEXT
)

if hook == 0:
    print('SetWinEventHook failed')
    sys.exit(1)

msg = ctypes.wintypes.MSG()

counter = 0
while user32.GetMessageW(ctypes.byref(msg), 0, 0, 0) != 0:
    user32.TranslateMessageW(msg)
    user32.DispatchMessageW(msg)

user32.UnhookWinEvent(hook)
ole32.CoUninitialize()