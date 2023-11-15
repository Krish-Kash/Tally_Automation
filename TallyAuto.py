import pandas as pd
import pyautogui as pg 
import time

excel_path='C:\\Users\\karri\\OneDrive\\Documents\\Desktop\\Internship Project\\TallPy\\Tally_Connections\\BASE FILE FOR DEMO TALLY ENTRY.xlsx'

df=pd.read_excel(excel_path,sheet_name="Sheet1")
party_name=df["Party_Name"].values
category=df["Under"].values
amount=df["Opening Balance"].values

zipped=zip(party_name,category,amount)

time.sleep(3)

pg.press("c")
time.sleep(1)

pg.typewrite("Ledger")
time.sleep(1)

pg.press("enter")
time.sleep(1)

#Start of the loop

for (a,b,c) in zipped:
    pg.typewrite(a)

    time.sleep(1)
    pg.press("enter")
    pg.press("enter")

    time.sleep(1)

    pg.typewrite(b)
    time.sleep(1)
    
    pg.press("enter")
    pg.press("enter")
    pg.press("enter")
    pg.press("enter")
    pg.press("enter")
    pg.press("enter")
    pg.press("enter")
    pg.press("enter")
    pg.press("enter")
    pg.press("enter")

    time.sleep(1)

    pg.typewrite(str(c))
    time.sleep(1)


    pg.press("enter")
    pg.press("enter")

    pg.hotkey("ctrl","a")
    time.sleep(1)

    pg.press("escape")
    time.sleep(1)

    pg.press("enter")
    time.sleep(1)

    pg.press("enter")
    time.sleep(1)
