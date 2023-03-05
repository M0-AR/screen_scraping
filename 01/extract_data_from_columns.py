import cv2
import pyperclip
import numpy as np
import pyautogui
import time

all_data = ''
copied_text = ''


def save_data_to_excel(cpr_nr):
    import os
    import shutil
    import win32com.client as win32

    # Clear the gen_py folder
    try:
        shutil.rmtree(os.path.join(os.environ['LOCALAPPDATA'], 'Temp', 'gen_py'))
    except:
        pass

    # Create an instance of the Excel application
    excel = win32.gencache.EnsureDispatch('Excel.Application')

    # Make Excel visible (optional)
    excel.Visible = True

    # Open a new workbook
    workbook = excel.Workbooks.Add()

    # Get the active sheet
    sheet = workbook.ActiveSheet

    # Split the all_data string into separate lines
    all_data_lines = all_data.split('\n')

    # Paste each line in a separate row in the first column
    for i, line in enumerate(all_data_lines):
        sheet.Cells(i + 1, 1).Value = line

    # Save the workbook with the name 'cprNumber'
    workbook.SaveAs(cpr_nr)

    # Close the workbook
    workbook.Close()

    # Quit Excel
    excel.Quit()


def go_one_page_back():
    import keyboard

    # Press and hold the Ctrl key
    keyboard.press('ctrl')

    # Press and release the C key
    keyboard.press('left')
    keyboard.release('left')

    # Release the Ctrl key
    keyboard.release('ctrl')

    time.sleep(4)


def select_all_columns_and_copy():
    # take a screenshot of the entire screen
    screenshot = pyautogui.screenshot()

    # save the screenshot as a file for to read
    screenshot.save("screenshot.png")

    # Load the screenshot
    screenshotNP = np.array(pyautogui.screenshot())

    # Convert to grayscale
    gray = cv2.cvtColor(screenshotNP, cv2.COLOR_BGR2GRAY)

    # Load the template image
    template = cv2.imread('07-alle-raekker.jpg', cv2.IMREAD_GRAYSCALE)

    # Find the text on the screenshot using template matching
    result = cv2.matchTemplate(gray, template, cv2.TM_CCOEFF_NORMED)
    y, x = np.unravel_index(result.argmax(), result.shape)

    # Click on the center of the text
    width, height = template.shape[::-1]
    pyautogui.moveTo(x + width / 2, y + height / 2)

    # # Simulate a right-click at the current mouse position
    ## pyautogui.click()

    # Move the mouse right by 150 pixels and down by 90 pixels
    y = y + 90
    pyautogui.move(150, 90)

    # Hold down the left mouse button at the current position
    pyautogui.mouseDown(button='left')

    # Delay the click for 1 second
    time.sleep(1)

    # Move the mouse to the right slowly and select an area simultaneously
    x2 = x + 2350  # example end x coordinate
    y2 = y + height / 2  # example end y coordinate
    pyautogui.moveTo(x2, y2, duration=0.3)  # move to the next position slowly

    # Hold up the left mouse button at the current position
    pyautogui.mouseUp(button='left')

    # # Simulate a right-click at the current mouse position
    pyautogui.rightClick()

    # Move the mouse down by 90 pixels
    pyautogui.move(20, 40)

    # Simulate a left-click at the current mouse position
    pyautogui.click()

    # Delay the click for 1 second
    time.sleep(1)


def extract_data_from_columns(cpr_nr):
    while True:
        select_all_columns_and_copy()

        # get the copied text from clipboard
        global copied_text
        copied_text = pyperclip.paste()

        # if data already exist that mean we are in the last page
        global all_data
        if copied_text in all_data:
            break

        # collect the whole data in data variable
        all_data = copied_text + '\n' + all_data + '\n'

        go_one_page_back()

    save_data_to_excel(cpr_nr)
    copied_text = ''
    all_data = ''
