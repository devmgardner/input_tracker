import os, sys, openpyxl, pygetwindow
from datetime import date
from time import time
from pynput import mouse, keyboard
#
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(os.path.realpath(__file__))
    return os.path.join(base_path, relative_path)
#
wb = openpyxl.Workbook()
ws = wb.active
ws.title = 'Events'
ws['A1'] = 'Timestamp'
ws['B1'] = 'Event Device'
ws['C1'] = 'Event Details'
ws['D1'] = 'Window Name'
wb.save(resource_path(f'{date.today()}-inputs.xlsx'))
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
#
def on_move(x, y):
    pass
    # event = {}
    # event['device'] = 'mouse'
    # event['details'] = f'Mouse pointer moved to {x},{y}'
    # event['window'] = pygetwindow.getActiveWindowTitle()
    # log(event)
#
def on_click(x, y, button, pressed):
    event = {}
    event['device'] = 'mouse'
    if pressed:
        event['details'] = f'Mouse button {button} pressed at {x},{y}'
    elif not pressed:
        event['details'] = f'Mouse button {button} released at {x},{y}'
    event['window'] = pygetwindow.getActiveWindowTitle()
    log(event)
#
def on_scroll(x, y, dx, dy):
    event = {}
    event['device'] = 'mouse'
    if dy < 0:
        event['details'] = f'Scrolled down at {x},{y}'
    else:
        event['details'] = f'Scrolled up at {x},{y}'
    event['window'] = pygetwindow.getActiveWindowTitle()
    log(event)
#
def on_press(key):
    event = {}
    event['device'] = 'keyboard'
    try:
        event['details'] = f'Key {key.char} pressed'
    except AttributeError:
        event['details'] = f'Key {key} pressed'
    event['window'] = pygetwindow.getActiveWindowTitle()
    log(event)
#
def on_release(key):
    if key == keyboard.Key.pause:
        event = {}
        event['device'] = 'keyboard'
        event['details'] = 'Pause/Break pressed, exiting script'
        event['window'] = pygetwindow.getActiveWindowTitle()
        log(event)
        # Stop program
        return False
#
# Collect events until released
with mouse.Listener(
        on_move=on_move,
        on_click=on_click,
        on_scroll=on_scroll) as m_listener:
    m_listener.join()
#
# Collect events until released
with keyboard.Listener(
        on_press=on_press,
        on_release=on_release) as k_listener:
    k_listener.join()