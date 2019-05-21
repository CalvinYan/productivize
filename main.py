# Productivize runs as a background desktop app, checking the name of the foreground application
# every second, and generating a report of the distribution of time spent using each application
# on command.

import win32
import win32gui, win32process
import time
import psutil

# Time spent for each app
time_dict = {}

while True:
    pid = win32process.GetWindowThreadProcessId(win32gui.GetForegroundWindow())
    print(pid)
    print(psutil.Process(pid[-1]).name())
    process_name = win32gui.GetWindowText(win32gui.GetForegroundWindow())
    # print(process_name)
    if process_name not in time_dict:
        time_dict[process_name] = 0
    time_dict[process_name] += 1
    print(time_dict)
    time.sleep(1)