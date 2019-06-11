# Original code graciously provided by Stack Overflow user David Hefferman

import csv
import json
import os
import sys
import time
import ctypes
import ctypes.wintypes
import threading
from win32 import win32api, win32gui, win32process
import win32con
import psutil

# User settings
settings = {}

idle_thread = threading.Thread() # Thread to run the afk check
afk_state = False # Has the system not received an input for longer than the timeout?
last_input_previous = win32api.GetTickCount() # Previous value of GetLastInputInfo()

# The number of seconds representing the last second of the day - used for end-of-day autosave
LAST_SECOND = 23 * 3600 + 59 * 60 + 59
date_change_thread = threading.Thread()

time_log = {} # The number of seconds spent on each window, using changes in window focus

num_updates = 0
last_app = '' # The name of the application used prior to window switch
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


# Periodically check if the user's afk state (afk or not afk) has changed and adjust the time counter accordingly
def idle_check():
    global idle_thread
    global last_time
    global afk_state

    last_input = int(win32api.GetLastInputInfo() / 1000.0)
    if isAFK(last_input):
        # Did the state switch from not afk to afk?
        if not afk_state:
            afk_state = True
            # Add the window usage time prior to state switch
            usage_time = max(last_input - last_time, 0)
            print("You've become afk after %d seconds" % (usage_time))
            updateLog(time_log, getAppName(), last_window, usage_time)
    else:
        # Did the state switch from afk to not afk?
        if afk_state:
            afk_state = False
            # Set the last_time pointer to the current time to "skip over" idle time
            last_time = int(win32api.GetTickCount() / 1000)
            print("Welcome back!")
    # Since we're here, we might as well check if this is the last update cycle before the date changes
    # and call onDateChange(). Code reformatting may be necessary
    # hour_minute = time.strftime('%H:%M')
    # num_seconds = int(time.strftime('%S'))
    # if hour_minute == "23:59" and num_seconds + 10.0 >= 60:
    #     onDateChange(time_log, last_window, last_time, last_input)
    # Restart the timer for another update cycle
    idle_thread = threading.Timer(0.1, idle_check)
    idle_thread.start()

idle_thread = threading.Timer(0.1, idle_check)
idle_thread.start()

def callback(hWinEventHook, event, hwnd, idObject, idChild, dwEventThread, dwmsEventTime):
    global num_updates
    global last_app
    global last_window
    global last_time
    length = user32.GetWindowTextLengthW(hwnd)
    buff = ctypes.create_unicode_buffer(length + 1)
    status = user32.GetWindowTextW(hwnd, buff, length + 1)
    name = buff.value
    if not afk_state and name != '' and name != last_window and name == win32gui.GetWindowText(win32gui.GetForegroundWindow()) and getAppName() not in getSetting("appExclude", {}):
        new_time = int(win32api.GetTickCount() / 1000)
        print(time.strftime('%X') + ' --> ' + name)
        if (last_time != 0):
            time_delta = new_time - last_time
            updateLog(time_log, last_app, last_window, time_delta)
            print(str(time_delta) + ' seconds have elapsed')
        last_app = getAppName()
        last_time = new_time
        last_window = name
        if num_updates == 30: 
            onExit()
        num_updates += 1

# Update the AFK state
# Encapsulation is for the purpose of adding additional validation later, like if video is playing or not
def isAFK(last_input):
    current_time = int(win32api.GetTickCount() / 1000)
    # Is the user afk?
    return current_time - last_input > getSetting("afkTimeoutSeconds", 120)

def getAppName():
    _, processID = win32process.GetWindowThreadProcessId(win32gui.GetForegroundWindow())
    app_name = psutil.Process(processID).name()
    return app_name

def readData(time_log):
    path_string = os.getenv('LOCALAPPDATA') + '\\Productivize\\logs\\'
    date_string = time.strftime('%d-%m-%Y')
    # Python's default access modes don't allow us to open the file for reading, create it if it doesn't
    # exist, don't truncate it, and start from the top, so a combination of modes is required
    with open(path_string + date_string + '.csv', "a") as dataFile: # Create if not exist
        pass
    with open(path_string + date_string + '.csv', "r", encoding = "utf-8") as dataFile: # Read, don't truncate, start from top
        lines = csv.reader(dataFile)
        for line in lines:
            if len(line) > 0:
                [app_name, window_name, seconds] = line
                updateLog(time_log, app_name, window_name, int(seconds))

# Update the total record of time spent on each application-window combination
def writeData(time_log):
    log_list = []
    # Convert dict to list
    for app_name, windows in time_log.items():
        for window_name, seconds in windows.items():
            log_list.append([app_name, window_name, str(seconds)])
    log_list = sorted(log_list, key = lambda app:(-int(app[2]), app[1], app[0]))
    path_string = os.getenv('LOCALAPPDATA') + '\\Productivize\\logs\\'
    date_string = time.strftime('%d-%m-%Y')
    with open(path_string + date_string + '.csv', 'w', newline = '', encoding = 'utf-8') as dataFile:
        out = csv.writer(dataFile)
        out.writerows(log_list)

# Retrieve and parse settings.json
def readSettings():
    path_string = os.getenv('LOCALAPPDATA') + '\\Productivize\\'
    with open(path_string + 'settings.json', 'r') as settingsFile:
        return json.loads(settingsFile.read())

# Retrieve a value from the settings dictionary
def getSetting(key, default):
    return settings[key] if key in settings else default

# Add the number of seconds spent on an application-window combination to its total duration, or add it if necessary
def updateLog(time_log, app_name, window_name, seconds):
    if app_name in time_log:
        current_value = time_log[app_name][window_name] if window_name in time_log[app_name] else 0
        time_log[app_name][window_name] = current_value + seconds
    else:
        # Create a new dictionary corresponsing to the application in question      
        time_log[app_name] = {window_name: seconds}

# All tasks that must be performed before the date changes at 12:00 AM
def onDateChange():
    global time_log
    print("Yeetimus")
    last_input = int(win32api.GetLastInputInfo()/1000)
    current_time = int(win32api.GetTickCount() / 1000)
    # Account for the possibility of being in/recently out of afk
    afk_timeout_seconds = getSetting("afkTimeoutSeconds", 120)
    if current_time - last_time <= afk_timeout_seconds:
        updateLog(time_log, last_window, current_time - last_time)
    elif current_time - last_input <= afk_timeout_seconds:
        updateLog(time_log, last_window, current_time - last_input)
    writeData(time_log)
    time_log = {}


# All tasks that must be performed before the program closes
def onExit():
    writeData(time_log)
    idle_thread.cancel()
    date_change_thread.cancel()
    quit()

readData(time_log)

settings = readSettings()

# Set timer to call onDateChange on the last second of the day
current_seconds = int(time.strftime("%H")) * 3600 + int(time.strftime("%M")) * 60 + int(time.strftime("%S"))
date_change_thread = threading.Timer(LAST_SECOND - current_seconds, onDateChange)
date_change_thread.start()
WinEventProc = WinEventProcType(callback)

user32.SetWinEventHook.restype = ctypes.wintypes.HANDLE
hook_focus = user32.SetWinEventHook(
    win32con.EVENT_OBJECT_FOCUS,
    win32con.EVENT_OBJECT_FOCUS,
    0,
    WinEventProc,
    0,
    0,
    win32con.WINEVENT_OUTOFCONTEXT | win32con.WINEVENT_SKIPOWNPROCESS
)
hook_namechange = user32.SetWinEventHook(
    win32con.EVENT_OBJECT_NAMECHANGE,
    win32con.EVENT_OBJECT_NAMECHANGE,
    0,
    WinEventProc,
    0,
    0,
    win32con.WINEVENT_OUTOFCONTEXT | win32con.WINEVENT_SKIPOWNPROCESS
)

if hook_focus == 0 or hook_namechange == 0:
    print('SetWinEventHook failed')
    sys.exit(1)

msg = ctypes.wintypes.MSG()

counter = 0
while user32.GetMessageW(ctypes.byref(msg), 0, 0, 0) != 0:
    user32.TranslateMessageW(msg)
    user32.DispatchMessageW(msg)

user32.UnhookWinEvent(hook_focus)
user32.UnhookWinEvent(hook_namechange)
ole32.CoUninitialize()