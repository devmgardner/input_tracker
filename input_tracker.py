import os, sys, openpyxl, pygetwindow, re, json
from datetime import date, timedelta
from time import time, sleep
from pynput import mouse, keyboard
# set up resource_path for compilation purposes
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(os.path.realpath(__file__))
    return os.path.join(base_path, relative_path)
# initialize today's workbook and set up Events tab
workbook = openpyxl.Workbook()
worksheet = workbook.active
worksheet.title = 'Events'
worksheet['A1'] = 'Timestamp'
worksheet['B1'] = 'Event Device'
worksheet['C1'] = 'Event Details'
worksheet['D1'] = 'Window Name'
# make them all bold
for cell in worksheet['A1:D1']:
    cell[0].font = openpyxl.styles.Font(bold=True)
workbook.save(resource_path(f'{date.today()}-inputs.xlsx'))
# log events as they happen, instead of trying to load everything into RAM
def log(event):
    # load the workbook we initialized in the beginning
    wb = openpyxl.load_workbook(resource_path(f'{date.today()}-inputs.xlsx'))
    ws = wb.active
    # get the max row, and increment
    row = ws.max_row + 1
    # the first cell in the row is the timestamp rounded to 4 decimal places
    ws[f'A{row}'] = round(time(),4)
    # determine device for event
    if event['device'] == 'mouse':
        ws[f'B{row}'] = 'Mouse'
    elif event['device'] == 'keyboard':
        ws[f'B{row}'] = 'Keyboard'
    # event details passed from each function
    ws[f'C{row}'] = event['details']
    # track windows for productivity tracking
    ws[f'D{row}'] = event['window']
    # save the workbook
    wb.save(resource_path(f'{date.today()}-inputs.xlsx'))
# print(f'log function loaded')
# this is the main reason i built this tool: to look at metrics for productivity
def metrics_1():
    # load the workbook
    wb = openpyxl.load_workbook(resource_path(f'{date.today()}-inputs.xlsx'))
    # create the Windows sheet to track time spent in each window overall
    windows_sheet = wb.create_sheet()
    windows_sheet.title = 'Windows'
    windows_sheet['A1'] = 'Window Name'
    windows_sheet['A1'].font = openpyxl.styles.Font(bold=True)
    windows_sheet['B1'] = 'Time Spent'
    windows_sheet['B1'].font = openpyxl.styles.Font(bold=True)
    # load the Events sheet as well to pull data
    events_sheet = wb['Events']
    # i was originally going to iterate over the window names themselves, but chose a different route
    # windows = []
    # for cell in events_sheet[f'D2:D{events_sheet.max_row}']:
    #     if not cell[0].value in windows:
    #         windows.append(cell[0].value)
    # create a dictionary to store the row value for each window name
    window_names_and_rows = {}
    # iterate through all rows in the events sheet
    for row in range(2,events_sheet.max_row+1):
        # setting some variables to clean up code
        window_name = events_sheet[f'D{row}'].value
        next_window_name = events_sheet[f'D{row+1}'].value
        prev_window_name = events_sheet[f'D{row-1}'].value
        # check if we're at the start of a new window name
        if window_name != prev_window_name:
            top_row = row
        # check if the next row is different, if so process the time spent
        if window_name != next_window_name:
            bottom_row = row
            # the time spent is equal to the bottom row's timestamp minus the top row's timestamp
            time_spent = events_sheet[f'A{bottom_row}'].value-events_sheet[f'A{top_row}'].value
            # if the window isn't already listed, add a new row with this window
            if window_name not in [cell[0].value for cell in windows_sheet[f'A2:A{windows_sheet.max_row}']]:
                # get the max row, to add to the next row
                maxrow = windows_sheet.max_row+1
                windows_sheet[f'A{maxrow}'] = window_name
                windows_sheet[f'B{maxrow}'] = time_spent
                window_names_and_rows[window_name] = maxrow
            # otherwise, just add the time_spent values together
            else:
                windows_sheet[f'B{window_names_and_rows[window_name]}'] = round(float(windows_sheet[f'B{window_names_and_rows[window_name]}'].value) + time_spent,4)
    # save the workbook, obviously
    wb.save(resource_path(f'{date.today()}-inputs.xlsx'))
# second metrics function for tracking how many times each key/button was pressed
def metrics_2():
    # load the workbook
    wb = openpyxl.load_workbook(resource_path(f'{date.today()}-inputs.xlsx'))
    # create the Buttons sheet to track how many times each button/key was utilized throughout the script
    ws = wb.create_sheet()
    ws.title = 'Buttons'
    ws['A1'] = 'Button/Key Name'
    ws['B1'] = 'Times Used'
    ws['A1'].font = openpyxl.styles.Font(bold=True)
    ws['B1'].font = openpyxl.styles.Font(bold=True)
    # load the Events sheet as well to pull data
    events_sheet = wb['Events']
    # load the list of all possible events tracked, and read as lines
    with open(resource_path('keys.txt'),'r') as fhand:
        lines = [line.strip() for line in fhand.readlines()]
    # iterate through the lines
    for line in lines:
        row = lines.index(line) + 2
        ws[f'A{row}'] = line
        # couldn't figure out the equals sign, couldn't be bothered to keep trying
        if line == '=':
            ws[f'A{row}'] = 'Equals'
        # set all values to 0 to start
        ws[f'B{row}'] = '0'
    # get a list of all the events from the Events sheet
    events = [event[0].value for event in events_sheet[f'C2:C{events_sheet.max_row}']]
    # make a list of all the different possible events, for indexing purposes
    all_events = [re.sub('\n','',event[0].value).lower() for event in ws[f'A2:A{ws.max_row}']]
    # iterate through all the events, checking for each type and incrementing the value
    for event in events:
        # check for mouse scroll down
        if event.startswith(f'Scrolled down'):
            ws[f'B{all_events.index("scrolldown")+2}'] = int(ws[f'B{all_events.index("scrolldown")+2}'].value)+1
        # check for mouse scroll up
        elif event.startswith(f'Scrolled up'):
            ws[f'B{all_events.index("scrollup")+2}'] = int(ws[f'B{all_events.index("scrollup")+2}'].value)+1
        # check all 3 mouse buttons
        elif event.startswith(f'Mouse button') and 'pressed' in event:
            actual = re.match(f'Mouse button Button\.(.*) pressed',event).group(1)
            if actual == 'left':
                ws[f'B{all_events.index("leftmouse")+2}'] = int(ws[f'B{all_events.index("leftmouse")+2}'].value)+1
            elif actual == 'right':
                ws[f'B{all_events.index("rightmouse")+2}'] = int(ws[f'B{all_events.index("rightmouse")+2}'].value)+1
            elif actual == 'middle':
                ws[f'B{all_events.index("middlemouse")+2}'] = int(ws[f'B{all_events.index("middlemouse")+2}'].value)+1
        # check for all character keys
        elif event.startswith(f'Key') and 'Key.' not in event:
            actual = re.match(f'Key (.*) pressed',event).group(1)
            # this has to be handled differently because = is annoying
            if actual == '=':
                ws[f'B{all_events.index("equals")+2}'] = int(ws[f'B{all_events.index("equals")+2}'].value)+1
            # this has to be handled differently because escape character
            elif actual == '\\':
                ws['B'+str(all_events.index('\\')+2)] = int(ws['B'+str(all_events.index('\\')+2)].value)+1
            # everything else is nice and easy
            else:
                ws[f'B{all_events.index(actual.lower())+2}'] = int(ws[f'B{all_events.index(actual.lower())+2}'].value)+1
        # check for 'special' keys
        elif event.startswith(f'Key') and 'Key.' in event:
            actual = re.match(f'Key Key\.(.*) pressed',event).group(1)
            # load json data file ot match special keys to their event names
            data = json.load(open(resource_path('keys.json')))
            ws[f'B{all_events.index(data[actual])+2}'] = int(ws[f'B{all_events.index(data[actual])+2}'].value)+1
    # save the workbook
    wb.save(resource_path(f'{date.today()}-inputs.xlsx'))
# this function creates a nice summary sheet and displays most used window and most used keys/buttons
def formatting():
    # load the workbook
    wb = openpyxl.load_workbook(resource_path(f'{date.today()}-inputs.xlsx'))
    # create the Buttons sheet to track how many times each button/key was utilized throughout the script
    ws = wb.create_sheet()
    ws.title = 'Summary'
    wb.move_sheet('Summary',offset=-3)
    # load the other sheets we need
    windows_sheet = wb['Windows']
    buttons_sheet = wb['Buttons']
    # make a list of all the times (still in seconds) from the Windows sheet
    times = [float(cell[0].value) for cell in windows_sheet[f'B2:B{windows_sheet.max_row}']]
    # grab the index of the longest time value, and add 2
    ind = times.index(max(times))+2
    # grab the window name and the window time
    max_time_window = windows_sheet[f'A{ind}'].value
    max_time = windows_sheet[f'B{ind}'].value
    # insert some column headers and make them Bold
    ws['A1'] = 'Most Used Window'
    ws['B1'] = 'Time Spent In Window'
    ws['A1'].font = openpyxl.styles.Font(bold = True)
    ws['B1'].font = openpyxl.styles.Font(bold = True)
    ws['A2'] = max_time_window
    # we want it pretty
    ws['B2'] = timedelta(seconds=max_time)
    # convert seconds to timedelta because i love pretty formatting
    for row in range(2,windows_sheet.max_row+1):
        windows_sheet[f'B{row}'] = timedelta(seconds=float(windows_sheet[f'B{row}'].value))
    # grab the buttons and their associated counters
    buttons = []
    times = []
    for row in range(2,buttons_sheet.max_row+1):
        buttons.append(buttons_sheet[f'A{row}'].value)
        times.append(int(buttons_sheet[f'B{row}'].value))
    # set up more headers
    ws['A5'] = 'Most Used Buttons'
    ws['B5'] = 'Counter'
    ws['A5'].font = openpyxl.styles.Font(bold = True)
    ws['B5'].font = openpyxl.styles.Font(bold = True)
    # iterate through the 10 most used buttons, remove them from the buttons and times lists, and append them to the sheet
    for i in range(6,16):
        ind = times.index(max(times))
        button = buttons.pop(ind)
        counter = times.pop(ind)
        ws[f'A{i}'] = button
        ws[f'B{i}'] = counter
    # save the workbook
    wb.save(resource_path(f'{date.today()}-inputs.xlsx'))
#    
def on_move(x, y):
    pass
    # i originally had the tracker tracking mouse movements, but holy shit that's a LOT of data and my terrible work PC couldn't handle it processing events every time I moved the mouse
    # event = {}
    # event['device'] = 'mouse'
    # event['details'] = f'Mouse pointer moved to {x},{y}'
    # event['window'] = pygetwindow.getActiveWindowTitle()
    # log(event)
# print(f'on_move function loaded')
# define what happens each time a mouse button is pressed
def on_click(x, y, button, pressed):
    # create an empty dictionary to store the payload for the log function
    event = {}
    # set the device name
    event['device'] = 'mouse'
    # if the mouse button is pressed...
    if pressed:
        event['details'] = f'Mouse button {button} pressed at {x},{y}'
        # print(f'Mouse button {button} pressed at {x},{y}')
    # ...or if the mouse button is released
    elif not pressed:
        event['details'] = f'Mouse button {button} released at {x},{y}'
        # print(f'Mouse button {button} released at {x},{y}')
    # grab the window title for productivity tracking
    event['window'] = pygetwindow.getActiveWindowTitle()
    # log the event
    log(event)
# print(f'on_click function loaded')
# what happens each time the mouse wheel is scrolled
def on_scroll(x, y, dx, dy):
    event = {}
    event['device'] = 'mouse'
    if dy < 0:
        event['details'] = f'Scrolled down at {x},{y}'
        # print(f'Scrolled down at {x},{y}')
    else:
        event['details'] = f'Scrolled up at {x},{y}'
        # print(f'Scrolled up at {x},{y}')
    event['window'] = pygetwindow.getActiveWindowTitle()
    log(event)
# print(f'on_scroll function loaded')
# what happens each time a key is pressed
def on_press(key):
    event = {}
    event['device'] = 'keyboard'
    # check for numpad keys
    if hasattr(key, 'vk') and 96 <= key.vk <= 105:
        if key.vk == 96:
            event['details'] = f'Key numpad0 pressed'
        if key.vk == 97:
            event['details'] = f'Key numpad1 pressed'
        if key.vk == 98:
            event['details'] = f'Key numpad2 pressed'
        if key.vk == 99:
            event['details'] = f'Key numpad3 pressed'
        if key.vk == 100:
            event['details'] = f'Key numpad4 pressed'
        if key.vk == 101:
            event['details'] = f'Key numpad5 pressed'
        if key.vk == 102:
            event['details'] = f'Key numpad6 pressed'
        if key.vk == 103:
            event['details'] = f'Key numpad7 pressed'
        if key.vk == 104:
            event['details'] = f'Key numpad8 pressed'
        if key.vk == 105:
            event['details'] = f'Key numpad9 pressed'
    else:
        try:
            event['details'] = f'Key {key.char} pressed'
            # print(f'Key {key.char} pressed')
        except AttributeError:
            event['details'] = f'Key {key} pressed'
            # print(f'Key {key} pressed')
    event['window'] = pygetwindow.getActiveWindowTitle()
    log(event)
# print(f'on_press function loaded')
# what happens each time a key is released
def on_release(key):
    # literally all we care about is if it's the Pause/Break key, because that's the key that kills the script
    # grab the global variables
    global m_listener, k_listener
    if key == keyboard.Key.pause:
        event = {}
        event['device'] = 'keyboard'
        event['details'] = 'Pause/Break pressed, exiting script'
        event['window'] = pygetwindow.getActiveWindowTitle()
        log(event)
        # log the event, sleep for 3 seconds (I cannot stress enough: MY WORK PC IS GARBAGE. I'M RUNNING AN A6-5400K APU FOR CAD WORK????)
        sleep(3)
        # stop the mouse listener
        m_listener.stop()
        # process the metrics
        metrics_1()
        metrics_2()
        formatting()
        # return False to kill the keyboard listener
        return False
# print(f'on_release function loaded')
#
# this is where the listeners are actually defined and started
# this took me an unreasonable amount of time to figure out, because i kept having issues (see my SO post) with one listener running but not the other, or neither of them running at all
global m_listener, k_listener
with mouse.Listener(on_move=on_move,on_click=on_click,on_scroll=on_scroll) as m_listener:
    with keyboard.Listener(on_press=on_press,on_release=on_release) as k_listener:
        m_listener.join()
        k_listener.join()