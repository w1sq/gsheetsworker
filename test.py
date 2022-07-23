import pyautogui as pg
from time import sleep

# print(pg.position())
i = 10
while i:
    pg.click(1680, 1030) #joinlobby
    sleep(0.4)
    pg.typewrite("228")
    sleep(0.6)
    pg.click(855, 600) #ok
    sleep(1)
    pg.click(955, 580) #incorrect password ok
    sleep(0.8)
    i -= 1