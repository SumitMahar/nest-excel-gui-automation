import time
import pyautogui as pg

pg.PAUSE = 1
pg.FAILSAFE = True

# this function is for refreshing the market watch
def ref_market_watch(rx=147, ry=230, lx=209, ly=645):
	pg.rightClick(rx, ry, duration=.3)
	pg.leftClick(lx, ly, duration=.4)

right_arrow_count = 3
# this is for excel gui automation
def change_name_box_to(cell, x=764, y=153):
    pg.doubleClick(x, y, duration=.1)
    # these x and y values are for upper left name box position
    # they might be different for another laptop or desktop display
    pg.press('enter')
    pg.click(x, y, duration=.1)
    pg.write(cell)
    pg.press('enter')

def copy_and_paste_columns():
    # copies the two specific columns 
    pg.keyDown('shift')
    time.sleep(.1)
    pg.keyDown('ctrl')
    time.sleep(.1)
    pg.press(['right','down'])
    pg.keyUp('shift')
    pg.press('c')
    pg.keyUp('ctrl')

    global right_arrow_count
    pg.press('right', presses=right_arrow_count)
    right_arrow_count += 3

    # this is for pasting the values only
    pg.hotkey('alt', 'e', 's', 'v')
    pg.press('enter')

# main loop
interval = 15

print('Starting the programe...\nPlease minimize this window.')
time.sleep(8) # 8 seconds of pause for minimizing the window.
while interval < 120:
    try:
        print(f'Refreshing the Market-Watch : {time.asctime()}', end='\n')
        ref_market_watch()
        change_name_box_to('f1')
        copy_and_paste_columns()
        interval += 15
        time.sleep(5) # 900 sec : 15 minutes
    except:
        print('Error..! Please setup the Excel and Nest windows accordingly.')
        time.sleep(20)