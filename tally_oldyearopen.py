
import pyautogui
import time
import pandas as pd
# 1. Change File location as per your requirement
df=pd.read_excel('E:\\python\\tally_py.xlsx')
time.sleep(20)
pyautogui.hotkey('alt', 'tab')


time.sleep(5)
#2. For looking multiyear file as stored in excel file @ Point no. 1
for i in df.itertuples():
    pyautogui.press("f3")
    #3. Select Company is the function in first page after log in to the Tally software
    pyautogui.write("Select Company")
    pyautogui.press("enter")
    time.sleep(3)
    #4. Specify Path is the function in second page after selecting Select Company as per point no. 3 in the Tally software
    pyautogui.write("Specify Path")
    pyautogui.press("enter")
    time.sleep(3)
    pyautogui.write("As per your File stored in Tally, in my computer it is stored in folder named previous so I write "Previous:"")
    pyautogui.press("enter")
    time.sleep(3)
    pyautogui.write(i.year)
    pyautogui.press("enter")
    pyautogui.write("username")
    pyautogui.press("enter")
    pyautogui.write("password")
    pyautogui.press("enter")
    time.sleep(20)
