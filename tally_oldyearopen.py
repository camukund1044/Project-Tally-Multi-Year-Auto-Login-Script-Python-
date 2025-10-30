
import pyautogui
import time
import pandas as pd
# Change File location as per your requirement
df=pd.read_excel('E:\\python\\tally_py.xlsx')
time.sleep(20)
pyautogui.hotkey('alt', 'tab')


time.sleep(5)
for i in df.itertuples():
    pyautogui.press("f3")
    pyautogui.write("Select Company")
    #pyautogui.press("up")
    #pyautogui.press("up")
    pyautogui.press("enter")
    time.sleep(3)
    pyautogui.write("Specify Path")
    #pyautogui.press("up")
    #pyautogui.press("up")
    #pyautogui.press("up")
    pyautogui.press("enter")
    time.sleep(3)
    pyautogui.write("Previous:")
    pyautogui.press("enter")
    time.sleep(3)
    #pyautogui.press("enter")
    pyautogui.write(i.year)
    pyautogui.press("enter")
    #pyautogui.press("enter")
    pyautogui.write("username")
    pyautogui.press("enter")
    pyautogui.write("password")
    #pyautogui.write(i.Date.strftime('%d-%m-%Y'))
    pyautogui.press("enter")
    time.sleep(20)
