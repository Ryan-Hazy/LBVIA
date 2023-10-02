import time
import pandas as pd
from openpyxl import Workbook, load_workbook
import pyautogui as pag
import os

path = r'C:\Users\SHARON\Desktop\Deposits\LBVIA\In Progress.xlsx'
wb = load_workbook(path, data_only=True)
ws = wb['Sheet1']

os.startfile(r'C:\Users\SHARON\Desktop\LBVIA UB')
time.sleep(3)

pag.press('enter')
time.sleep(1)

pag.press(['alt'])
pag.press('right', presses=3)
pag.press('down', presses=16)
pag.press('enter')
time.sleep(.5)

pag.click(1320,300)
time.sleep(.1)
ini = 3


while ws["A"+str(ini)].data_type == "s":
    acct = str(ws["A"+str(ini)].value)
    pag.typewrite(acct)
    time.sleep(.2)
    pag.press('tab')
    amt = str(ws["C"+str(ini)].value)
    pag.typewrite(amt)
    time.sleep(.2)
    pag.press('enter')
    time.sleep(.2)
    ini = int(ini) + 1

print("done")