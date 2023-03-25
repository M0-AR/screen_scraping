import time

import keyboard as keyboard
import numpy as np
import pyautogui
import cv2

screenshot = pyautogui.screenshot()

# save the screenshot as a file for to read
screenshot.save("number.png")

# Load the screenshot
screenshotNP = np.array(pyautogui.screenshot())

# Convert to grayscale
gray = cv2.cvtColor(screenshotNP, cv2.COLOR_BGR2GRAY)

# Load the template image
template = cv2.imread('print.jpg', cv2.IMREAD_GRAYSCALE)

# Find the text on the screenshot using template matching
result = cv2.matchTemplate(gray, template, cv2.TM_CCOEFF_NORMED)
y, x = np.unravel_index(result.argmax(), result.shape)

# Click on the center of the text
width, height = template.shape[::-1]
pyautogui.moveTo(x + width / 2, y + height / 2)

# Move the mouse to the right by 100 pixels
pyautogui.move(200, 0)

# Simulate a left-click at the current mouse position
pyautogui.click()

# Hold up the left mouse button at the current position
pyautogui.mouseDown(button='left')

# Move the mouse cursor to the right and down by 50 pixels each time
x, y = pyautogui.position()
while x < 500 and y < 500:
    pyautogui.moveRel(50, 50)
    x, y = pyautogui.position()

# Drag the mouse cursor to the new position to select text
# pyautogui.drag(x + 200, y, button='left')
pyautogui.drag(x + 500, y + 500, button='left')

time.sleep(1)

# Press and hold the Ctrl key
keyboard.press('ctrl')

# Press and release the C key
keyboard.press('c')
keyboard.release('c')

# Release the Ctrl key
keyboard.release('ctrl')


import win32com.client as win32

# Create an instance of the Excel application
excel = win32.gencache.EnsureDispatch('Excel.Application')

# Make Excel visible (optional)
excel.Visible = True

# Open a new workbook
workbook = excel.Workbooks.Add()

# Get the active sheet
sheet = workbook.ActiveSheet

# Paste the contents of the clipboard
sheet.Paste()

# Save the workbook with the name "hi"
workbook.SaveAs('hi')

# Close the workbook
workbook.Close()

# Quit Excel
excel.Quit()
