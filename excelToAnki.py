### takes an excel spreadsheet input by the user and from starting word index to ending word index
### adds these words to a new word mode anki program on the left of the screen.
### Input/front of card comes from column in sheet named "Card Front" back comes from "Card Back"

import pandas
import pyperclip
import pyautogui
import keyboard

excelSheet = "../Chinese Homework/VocabExcel/HSK6.xlsx"
# excelSheet = "C:/Users/adamz/Desktop/Documents/Personal/Polish/Excel Anki/Duolingo Vocab 1.xlsx"
df = pandas.read_excel(excelSheet)

chinese = df['Chinese'].values
pinyin = df['Pinyin'].values
english = df['English'].values
lesson = df['Lesson'].values
entryType = df['Character'].values

lessonNumber = int(input("please input HSK6 lesson number: "))
type = str(input("please input 'False' for word input or 'True' for character input: "))

for i in range(len(chinese)):
    if lesson[i] == lessonNumber and type == str(entryType[i]):
        if keyboard.is_pressed('q'):
            quit()
        pyperclip.copy(chinese[i])
        # paste into anki
        pyautogui.moveTo(550,200)
        pyautogui.click()
        pyautogui.keyDown("ctrlleft")
        pyautogui.press("v")
        pyautogui.keyUp("ctrlleft")

        pyperclip.copy(english[i])
        # paste into anki
        pyautogui.moveTo(550,300)
        pyautogui.click()
        pyautogui.keyDown("ctrlleft")
        pyautogui.press("v")
        pyautogui.keyUp("ctrlleft")

        pyperclip.copy(pinyin[i])
        # paste into anki
        pyautogui.moveTo(550,400)
        pyautogui.click()
        pyautogui.keyDown("ctrlleft")
        pyautogui.press("v")
        pyautogui.keyUp("ctrlleft")

        # click enter
        pyautogui.moveTo(550,1110)
        pyautogui.click()
        