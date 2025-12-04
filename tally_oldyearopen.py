
import pyautogui
import time
import pandas as pd
# 1. Change File location as per your requirement
df=pd.read_excel('E:\\python\\tally_py.xlsx')
time.sleep(20)
# 1.1 This will change to next tab from python script. You should ensure that the opened tally software should be in next tab to python programme.
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
    #5. User name of tally file. I assumed that you have same user name and password across all the years.
    pyautogui.write("username")
    pyautogui.press("enter")
    pyautogui.write("password")
    pyautogui.press("enter")
    #6. Wait till the file is loading. You may increase decrease the time based on size of your file, lower time will generate errors.
    time.sleep(20)
