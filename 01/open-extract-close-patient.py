import pyautogui
import cv2
import numpy as np
import time


def click_by_mouse_on(position, move_x=0, move_y=0):
    # take a screenshot of the entire screen
    screenshot = pyautogui.screenshot()

    # save the screenshot as a file for to read
    screenshot.save("screenshot.png")

    # Load the screenshot
    screenshotNP = np.array(pyautogui.screenshot())

    # Convert to grayscale
    gray = cv2.cvtColor(screenshotNP, cv2.COLOR_BGR2GRAY)

    # Load the template image
    template = cv2.imread(position, cv2.IMREAD_GRAYSCALE)

    # Find the text on the screenshot using template matching
    result = cv2.matchTemplate(gray, template, cv2.TM_CCOEFF_NORMED)
    y, x = np.unravel_index(result.argmax(), result.shape)

    # Move the mouse to the desire area
    x = x + move_x
    y = y + move_y

    # Click on the center of the text
    width, height = template.shape[::-1]
    pyautogui.moveTo(x + width / 2, y + height / 2)

    # Simulate a right-click at the current mouse position
    pyautogui.click()


def read_cpr_nr_from_excel():
    import pandas as pd

    # Read the excel file
    df = pd.read_excel('cprs.xlsx')

    # Extract the second column without the first column
    cprs = df.iloc[:, 0].values

    return cprs


def input_cpr_nr(cpr_nr):
    click_by_mouse_on('02-navn-or-cprNr.jpg', 250, 0)

    import pyperclip
    # Copy a variable value to the clipboard
    pyperclip.copy(cpr_nr)

    import keyboard
    # Press and hold the Ctrl key
    keyboard.press('ctrl')

    # Press and release the C key
    keyboard.press('v')
    keyboard.release('v')

    # Release the Ctrl key
    keyboard.release('ctrl')


cpr_nrs = read_cpr_nr_from_excel()
cpr_nr = cpr_nrs[0]

# for cpr_nr in cpr_nrs:
click_by_mouse_on('01-patientopslag.jpg')
# Delay the click for 1 second
time.sleep(1)

input_cpr_nr(cpr_nr)
# Delay the click for 1 second
time.sleep(1)

click_by_mouse_on('03-find-patient.jpg')
# Delay the click for 1 second
time.sleep(1)

click_by_mouse_on('04-vaelg.jpg')
# Delay the click for 1 second
time.sleep(1)

click_by_mouse_on('05-if-aabn-journal.jpg')
# Delay the click for 1 second
time.sleep(1)

click_by_mouse_on('06-laboratoriesvar.jpg')
# Delay the click for 1 second
time.sleep(2)

# TODO: call extract-data-from-columns

click_by_mouse_on('08-close-by-x.jpg', 75, 0)
# Delay the click for 1 second
time.sleep(1)