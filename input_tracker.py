import os, sys, openpyxl, pygetwindow
from datetime import date
from time import time, sleep
from pynput import mouse, keyboard
#
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(os.path.realpath(__file__))
    return os.path.join(base_path, relative_path)
#
workbook = openpyxl.Workbook()
worksheet = workbook.active
worksheet.title = 'Events'
worksheet['A1'] = 'Timestamp'
worksheet['B1'] = 'Event Device'
worksheet['C1'] = 'Event Details'
worksheet['D1'] = 'Window Name'
workbook.save(resource_path(f'{date.today()}-inputs.xlsx'))
#
def log(event):
    wb = openpyxl.load_workbook(resource_path(f'{date.today()}-inputs.xlsx'))
    ws = wb.active
    row = ws.max_row + 1
    ws[f'A{row}'] = time()
    if event['device'] == 'mouse':
        ws[f'B{row}'] = 'Mouse'
    elif event['device'] == 'keyboard':
        ws[f'B{row}'] = 'Keyboard'
    ws[f'C{row}'] = event['details']
    ws[f'D{row}'] = event['window']
    wb.save(resource_path(f'{date.today()}-inputs.xlsx'))
print(f'log function loaded')
#
def on_move(x, y):
    pass
    # event = {}
    # event['device'] = 'mouse'
    # event['details'] = f'Mouse pointer moved to {x},{y}'
    # event['window'] = pygetwindow.getActiveWindowTitle()
    # log(event)
print(f'on_move function loaded')
#
def on_click(x, y, button, pressed):
    event = {}
    event['device'] = 'mouse'
    if pressed:
        event['details'] = f'Mouse button {button} pressed at {x},{y}'
        print(f'Mouse button {button} pressed at {x},{y}')
    elif not pressed:
        event['details'] = f'Mouse button {button} released at {x},{y}'
        print(f'Mouse button {button} released at {x},{y}')
    event['window'] = pygetwindow.getActiveWindowTitle()
    log(event)
print(f'on_click function loaded')
#
def on_scroll(x, y, dx, dy):
    event = {}
    event['device'] = 'mouse'
    if dy < 0:
        event['details'] = f'Scrolled down at {x},{y}'
        print(f'Scrolled down at {x},{y}')
    else:
        event['details'] = f'Scrolled up at {x},{y}'
        print(f'Scrolled up at {x},{y}')
    event['window'] = pygetwindow.getActiveWindowTitle()
    log(event)
print(f'on_scroll function loaded')
#
def on_press(key):
    event = {}
    event['device'] = 'keyboard'
    try:
        event['details'] = f'Key {key.char} pressed'
        print(f'Key {key.char} pressed')
    except AttributeError:
        event['details'] = f'Key {key} pressed'
        print(f'Key {key} pressed')
    event['window'] = pygetwindow.getActiveWindowTitle()
    log(event)
print(f'on_press function loaded')
#
def on_release(key):
    global m_listener, k_listener
    if key == keyboard.Key.pause:
        event = {}
        event['device'] = 'keyboard'
        event['details'] = 'Pause/Break pressed, exiting script'
        event['window'] = pygetwindow.getActiveWindowTitle()
        log(event)
        sleep(3)
        m_listener.stop()
        return False
print(f'on_release function loaded')
#
# Collect events until released
global m_listener
m_listener = mouse.Listener(on_move=on_move,on_click=on_click,on_scroll=on_scroll)
m_listener.start()
print(f'm_listener started')
print(m_listener.is_alive())
#
# Collect events until released
global k_listener
k_listener = keyboard.Listener(on_press=on_press,on_release=on_release)
k_listener.start()
print(f'k_listener started')
print(k_listener.is_alive())