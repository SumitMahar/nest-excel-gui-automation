import time
import pyautogui as pg

pg.PAUSE = 1
pg.FAILSAFE = True

# this function is for refreshing the market watch
def ref_market_watch(rx=147, ry=230, lx=209, ly=645):
	pg.rightClick(rx, ry, duration=.1)
	pg.leftClick(lx, ly, duration=.1)

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
    '''
    "does not work in some systems
    an alternative code is written below"
    
    pg.keyDown('shiftleft')
    time.sleep(.2)
    pg.keyDown('ctrlleft')
    time.sleep(.2)
    pg.press(['right','down'])
    pg.keyUp('shift')
    time.sleep(.1)
    pg.press('c')
    pg.keyUp('ctrl')
    '''

    # Alternative to above code
    # we will have to specify the x and y values
    # of the first column
    # to know x and y values use the pyautogui.displayMousePosition() function
    x = 1000
    y = 185
    pg.click(x, y)
    pg.dragTo(x+60, y)
    pg.hotkey('ctrl', 'c')

    # this might be even more simpler if you define a name for
    # the columns that you want to copy from excel sheet, 
    # for defining name select the columns right click and then click on define name
    # then instead of writing A1 in the name box we will simply write defined name for
    # the columns to copy and they will be auto selected and the we will only have to copy and paste the values

    global right_arrow_count
    pg.press('right', presses=right_arrow_count)
    right_arrow_count += 3

    # this is for pasting the values only
    pg.hotkey('ctrl', 'v')
    time.sleep(.1)
    pg.press(['ctrl', 'v'])
    pg.press('escape')

# main loop
interval = 15

print('Starting the programe...\nPlease minimize this window.')
time.sleep(5) # 8 seconds of pause for minimizing the window.
while interval < 240:
    try:
        print(f'Refreshing the Market-Watch : {time.asctime()}', end='\n')
        ref_market_watch()
        change_name_box_to('A1')
        copy_and_paste_columns()
        interval += 15
        time.sleep(5) # 900 sec : 15 minutes
    except Exception as e:
        print('Error..! Please setup the Excel and Nest windows accordingly.')
        print('Exception was:', e)
        pg.keyUp('shift')
        pg.keyUp('ctrl')
        time.sleep(2)
        print('R.A.Count: ', right_arrow_count)
        break
