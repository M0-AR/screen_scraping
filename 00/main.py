import cv2
import numpy as np
import pyautogui
import time

# -------------------------------- Select all Columns -------------------------------
# take a screenshot of the entire screen
screenshot = pyautogui.screenshot()

# save the screenshot as a file for to read
screenshot.save("number.png")

# Load the screenshot
screenshotNP = np.array(pyautogui.screenshot())

# Convert to grayscale
gray = cv2.cvtColor(screenshotNP, cv2.COLOR_BGR2GRAY)

# Load the template image
template = cv2.imread('date.jpg', cv2.IMREAD_GRAYSCALE)

# Find the text on the screenshot using template matching
result = cv2.matchTemplate(gray, template, cv2.TM_CCOEFF_NORMED)
y, x = np.unravel_index(result.argmax(), result.shape)

# Click on the center of the text
width, height = template.shape[::-1]
pyautogui.moveTo(x + width / 2, y + height / 2)

# # Simulate a right-click at the current mouse position
# pyautogui.click()

# Hold down the left mouse button at the current position
pyautogui.mouseDown(button='left')

# Delay the click for 1 second
time.sleep(1)

# Move the mouse to the right slowly and select an area simultaneously
x2 = x + 2000  # example end x coordinate
y2 = y + height / 2  # example end y coordinate
pyautogui.moveTo(x2, y2, duration=0.25)  # move to the next position slowly

# Hold up the left mouse button at the current position
pyautogui.mouseUp(button='left')

# Get the size of the primary screen
screen_width, screen_height = pyautogui.size()

# Calculate the center coordinates of the screen
center_x, center_y = screen_width // 2, screen_height // 2

# Move the mouse cursor to the center of the screen
pyautogui.moveTo(center_x, center_y)

# # Simulate a right-click at the current mouse position
pyautogui.rightClick()

# Move the mouse down by 90 pixels
pyautogui.move(30, 90)

# Simulate a left-click at the current mouse position
pyautogui.click()

# Delay the click for 1 second
time.sleep(1)


# -------------------------- See All Result -------------------------------------
# take a screenshot of the entire screen
screenshot = pyautogui.screenshot()

# save the screenshot as a file for to read
screenshot.save("number.png")

# Load the screenshot
screenshotNP = np.array(pyautogui.screenshot())

# Convert to grayscale
gray = cv2.cvtColor(screenshotNP, cv2.COLOR_BGR2GRAY)

# Load the template image
template = cv2.imread('02-all-result.jpg', cv2.IMREAD_GRAYSCALE)

# Find the text on the screenshot using template matching
result = cv2.matchTemplate(gray, template, cv2.TM_CCOEFF_NORMED)
y, x = np.unravel_index(result.argmax(), result.shape)

# Click on the center of the text
width, height = template.shape[::-1]
pyautogui.moveTo(x + width / 2, y + height / 2)

# Simulate a left-click at the current mouse position
pyautogui.click()

# Delay the click for 1 second
time.sleep(4)

# -------------------------------- Scroll Down ----------------------------------------
# Center then click on white place in Datoer some raekker
# take a screenshot of the entire screen
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

# Delay the click for 1 second
time.sleep(1)

# Get the dimensions of the screen
screen_width, screen_height = pyautogui.size()

# Calculate the center point of the screen
center_x = int(screen_width / 2)
center_y = int(screen_height / 2)

# Move the mouse to the center point of the screen
pyautogui.moveTo(center_x, center_y)

# Scroll to the right bottom
pyautogui.scroll(-10000)

# Delay the click for 1 second
time.sleep(3)

# Scroll to the right bottom
pyautogui.scroll(-10000)

# Delay the click for 1 second
time.sleep(1)

# -------------------------------------- Scroll to the right ----------------------------------
# take a screenshot of the entire screen
screenshot = pyautogui.screenshot()

# save the screenshot as a file for to read
screenshot.save("number.png")

# Load the screenshot
screenshotNP = np.array(pyautogui.screenshot())

# Convert to grayscale
gray = cv2.cvtColor(screenshotNP, cv2.COLOR_BGR2GRAY)

# Load the template image
template = cv2.imread('scroll-bar-bottom-left.jpg', cv2.IMREAD_GRAYSCALE)

# Find the text on the screenshot using template matching
result = cv2.matchTemplate(gray, template, cv2.TM_CCOEFF_NORMED)
y, x = np.unravel_index(result.argmax(), result.shape)

# Click on the center of the text
width, height = template.shape[::-1]
pyautogui.moveTo(x + width / 2, y + height / 2)

# hold down the left mouse button and move the cursor
# to the right until a certain sign is found in the current screenshot

# Load the screenshot as an image
screenshot = cv2.imread('number.png')

# Load the template image of the sign to search for
sign_template = cv2.imread('scroll-bar-bottom-right.jpg')

# Find the sign in the screenshot using OpenCV's template matching algorithm
result = cv2.matchTemplate(screenshot, sign_template, cv2.TM_CCOEFF_NORMED)
y, x = np.unravel_index(result.argmax(), result.shape)

# Get the current cursor position
cursor_x, cursor_y = pyautogui.position()

# Hold down the left mouse button at the current position
pyautogui.mouseDown(button='left')

# Move the cursor to the right until the sign is found
while x > cursor_x:
    # pyautogui.moveTo(cursor_x + 500, cursor_y, duration=0.1)
    pyautogui.moveTo(cursor_x + 1000, cursor_y, duration=0.1)
    cursor_x, _ = pyautogui.position()
    result = cv2.matchTemplate(screenshot, sign_template, cv2.TM_CCOEFF_NORMED)
    y, x = np.unravel_index(result.argmax(), result.shape)

# Release the left mouse button
pyautogui.mouseUp(button='left')
# Delay the click for 1 second
time.sleep(2)
# ----------------------------------- Select the whole data -------------------------------------
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

time.sleep(1)


# ----------------------------- Save to Excel ------------------------------
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
