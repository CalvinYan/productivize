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
import PySimpleGUI as sg
import unicodedata

# The productivize display window
window = None

# User settings
settings = {}

idle_thread = threading.Thread() # Thread to run the afk check
afk_state = False # Has the system not received an input for longer than the timeout?

# The day of the month at runtime start - used for end-of-day autosave
current_date = time.strftime('%d')

time_log = {} # The number of seconds spent on each window, using changes in window focus

# State variables
last_app = '' # The name of the application used prior to window switch
last_window = win32gui.GetWindowText(win32gui.GetForegroundWindow()) # The name of the recently unfocused window
last_time = int(win32api.GetTickCount() / 1000) # integer timestamp of the last change in focus

# The word(s) used to filter the table's contents
window_filter = ''

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
            updateLog(time_log, last_app, last_window, usage_time)
    else:
        # Did the state switch from afk to not afk?
        if afk_state:
            afk_state = False
            # Set the last_time pointer to the current time to "skip over" idle time
            last_time = int(win32api.GetTickCount() / 1000)
            print("Welcome back!")
    # Since we're here, we might as well check if the date has changed and and call onDateChange()
    if time.strftime('%d') != current_date:
        onDateChange()
    # Restart the timer for another update cycle
    idle_thread = threading.Timer(0.1, idle_check)
    idle_thread.start()

idle_thread = threading.Timer(0.1, idle_check)
idle_thread.start()

def callback(hWinEventHook, event, hwnd, idObject, idChild, dwEventThread, dwmsEventTime):
    global last_app
    global last_window
    global last_time
    length = user32.GetWindowTextLengthW(hwnd)
    buff = ctypes.create_unicode_buffer(length + 1)
    status = user32.GetWindowTextW(hwnd, buff, length + 1)
    name = buff.value
    app_name = getAppName()
    if not afk_state and name != '' and name != last_window and name == win32gui.GetWindowText(win32gui.GetForegroundWindow()) and app_name not in getSetting("appExclude", {}):
        new_time = int(win32api.GetTickCount() / 1000)
        print(time.strftime('%X') + ' --> ' + name)
        if (last_time != 0):
            time_delta = new_time - last_time
            updateLog(time_log, last_app, last_window, time_delta)
            print(str(time_delta) + ' seconds have elapsed')
        last_app = app_name
        last_time = new_time
        last_window = name
        # TCL can't display emojis so we need to clean up the window name
        last_window = ''.join([letter for letter in name if ord(letter) < 65536])

# Update the AFK state
# Encapsulation is for the purpose of adding additional validation later, like if video is playing or not
def isAFK(last_input):
    # Does current window contain an afk-exluded keyword, like "YouTube"?
    for keyword in getSetting('afkExclude', []):
        if containsIgnoreCase(last_window, keyword):
            return False
    current_time = int(win32api.GetTickCount() / 1000)
    # Is the user afk?
    return current_time - last_input > getSetting("afkTimeoutSeconds", 120)

def getAppName():
    global afk_state
    try:
        _, processID = win32process.GetWindowThreadProcessId(win32gui.GetForegroundWindow())
        app_name = psutil.Process(processID).name()
        return app_name
    except:
        # This tends to happen if the computer is going to sleep, so afk is safe to assume -- 
        # even it it's some other reason, the afk will be undone too quickly to do damage
        afk_state = True
        return None
    

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

# Convert time_log dict to a list of each window sorted by decreasing usage time
def sortDataByWindow(time_log, display, compute_sum=False, filter=True):
    log_list = []
    time_sum = 0
    # Convert dict to list
    for app_name, windows in time_log.items():
        for window_name, seconds in windows.items():
            # Apply window name filter
            if not filter or window_filter == '' or containsIgnoreCase(window_name, window_filter):
                log_list.append([app_name, window_name, seconds])
                time_sum += seconds
    log_list = sorted(log_list, key = lambda app:(-int(app[2]), app[1], app[0]))
    time_sum = max(time_sum, 1) # Avoid divide by zero
    # Pretty-format times and add additional data
    if display:
        log_list = [[app_name, window_name, timeString(seconds), '%.2f%%' % (seconds/time_sum*100)] for app_name, window_name, seconds in log_list]
    return log_list if not compute_sum else (log_list, time_sum)

# Sort apps in time_log dict by decreasing usage time and then sort by window within each app
def sortDataByApp(time_log, display):
    app_list = []
    # Convert dict to list
    for app_name, windows in time_log.items():
        # Extract a list of all windows run by the same application and sort the component windows
        window_list = [[window_name, seconds] for window_name, seconds in windows.items()]
        window_list = sorted(window_list, key = lambda window: (-window[1], window[0]))
        app_list.append([app_name, window_list])
    app_list = sorted(app_list, key = lambda app:sum([window[1] for window in app[1]]), reverse = True)
    print(app_list)
    # log_list = sorted(log_list, key = lambda app:(-int(app[2]), app[1], app[0]))

def NKFD(string):
    # Two calls to normalize is necessary for corner cases apparently
    return unicodedata.normalize('NFKD', unicodedata.normalize('NFKD', string).casefold())

def containsIgnoreCase(string, substring):
    return NKFD(substring) in NKFD(string)

# Convert seconds to hours, minutes, seconds for GUI display
def timeString(seconds):
    time_string = ''
    hours = int(seconds / 3600)
    minutes = int(seconds % 3600 / 60)
    seconds = seconds % 60
    if hours > 0:
        time_string = f'{hours}h {minutes}m {seconds}s'
    elif minutes > 0:
        time_string = f'{minutes}m {seconds}s'
    else:
        time_string = f'{seconds}s'
    return time_string

# Update the total record of time spent on each application-window combination
def writeData(time_log):
    log_list = sortDataByWindow(time_log, display=False, filter=False)
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
    global last_time
    global current_date
    # Save currently active window to log
    if not afk_state:
        seconds = int(win32api.GetTickCount() / 1000) - last_time
        updateLog(time_log, last_app, last_window, seconds)
    writeData(time_log)
    time_log = {}
    # Reset date, timer
    current_date = time.strftime('%d')
    last_time = int(win32api.GetTickCount() / 1000)
    updateDisplay(window,time_log)

# Modify sg display elements to reflect changes in time_log
def updateDisplay(window, time_log):
    global last_time
    # Save currently active window to log
    if not afk_state:
        seconds = int(win32api.GetTickCount() / 1000) - last_time
        updateLog(time_log, last_app, last_window, seconds)
        last_time += seconds
    # Might want to think of a way to get values and time_sum in one function call
    values, time_sum = sortDataByWindow(time_log, display=True, compute_sum=True)
    window.FindElement('__data__').Update(values)
    window.FindElement('__time_sum__').Update(f'Total: {timeString(time_sum)}')
    window.FindElement('__time_subtotal__').Update('Subtotal: 0s')

# Modify displayed time subtotal only. Separate from updateDisplay since we need the table to remain unchanged as we read it
def updateSubtotal(selected_rows):
    print(selected_rows)
    data = sortDataByWindow(time_log, display=False)
    # Extract corresponding rows
    selected_data = [data[row] for row in selected_rows]
    subtotal = sum([seconds for _, _, seconds in selected_data])
    print("Subtotal:", subtotal)
    window.FindElement('__time_subtotal__').Update(f'Subtotal: {timeString(subtotal)}')

# All tasks that must be performed before the program closes
def onExit():
    # Save currently active window to log
    if not afk_state:
        seconds = int(win32api.GetTickCount() / 1000) - last_time
        updateLog(time_log, last_app, last_window, seconds)
    writeData(time_log)
    idle_thread.cancel()
    print("Goodbye!")
    quit()

readData(time_log)

settings = readSettings()

last_app = getAppName()

data, time_sum = sortDataByWindow(time_log, display=True, compute_sum=True)
if (len(data) == 0):
    # SimpleGUI's table doesn't like to be empty initially so add a placeholder
    data = [['python3.exe', 'Productivize', '0s', '0.00%']]
headings = ['Application Name', 'Window Name', 'Time Spent', 'Percent Share']
window_layout = [[sg.Table(values=data, headings=headings, auto_size_columns=True, max_col_width = 50, num_rows = 20, enable_events = True, key='__data__')], [sg.Text(f'Total: {timeString(time_sum)}', auto_size_text=False, key='__time_sum__'), sg.Text('Subtotal: 0s', auto_size_text=False, key='__time_subtotal__')], [sg.Button(button_text='Update'), sg.Button(button_text='Clear'), sg.Text('Window name filter:'), sg.Input(do_not_clear=True, key='__filter__'), sg.Button(button_text='Filter'), sg.Button(button_text='Reset Filter')]]
window = sg.Window('Productivize').Layout(window_layout)

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

while True:
    event, values = window.Read()
    if event == 'Update':
        updateDisplay(window, time_log)
    if event == 'Clear':
        time_log = {}
        last_time = int(win32api.GetTickCount() / 1000)
        updateDisplay(window, time_log)
    if event == 'Filter':
        window_filter = values['__filter__']
        updateDisplay(window, time_log)
    if event == 'Reset Filter':
        window_filter = ''
        window.FindElement('__filter__').Update('')
        updateDisplay(window, time_log)
    if event == '__data__':
        selected_rows = values['__data__']
        # Calculate new time subtotal
        updateSubtotal(selected_rows)
    if event is None or event == 'Exit':
        onExit()
    print(event, values)
    if user32.PeekMessageW(ctypes.byref(msg), 0, 0, 0) != 0:
        # user32.TranslateMessageW(msg)
        user32.DispatchMessageW(msg)

user32.UnhookWinEvent(hook_focus)
user32.UnhookWinEvent(hook_namechange)
ole32.CoUninitialize()