import os, sys, openpyxl
from date import date
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
wb.save(resource_path(f'{date.today()}-inputs.xlsx'))
#
def on_move(x, y):
    print('Pointer moved to {0}'.format(
        (x, y)))
#
def on_click(x, y, button, pressed):
    print('{0} at {1}'.format(
        'Pressed' if pressed else 'Released',
        (x, y)))
    if not pressed:
        # Stop listener
        return False
#
def on_scroll(x, y, dx, dy):
    print('Scrolled {0} at {1}'.format(
        'down' if dy < 0 else 'up',
        (x, y)))
#
# Collect events until released
with mouse.Listener(
        on_move=on_move,
        on_click=on_click,
        on_scroll=on_scroll) as m_listener:
    m_listener.join()
#