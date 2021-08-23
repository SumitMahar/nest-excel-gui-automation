import time
import pyautogui as pg

right_arrow_count = 3

def ref_market_watch():
	pg.rightClick(147, 230, duration=.3)
	pg.leftClick(209, 645, duration=.4)

def change_name_box_to(cell):
    pg.doubleClick(764, 153, duration=.1)
    pg.press('enter')
    pg.click(764, 153, duration=.1)
    pg.write(cell)
    pg.press('enter')
