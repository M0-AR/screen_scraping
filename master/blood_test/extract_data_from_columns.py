import cv2
import pyperclip
import numpy as np
import time
from global_functions import move_mouse_on, click_by_mouse_on
import pyautogui
pyautogui.FAILSAFE = False

def go_one_page_back():
    import keyboard

    # Press and hold the Ctrl key
    keyboard.press('ctrl')

    # Press and release the C key
    keyboard.press('left')
    keyboard.release('left')

    # Release the Ctrl key
    keyboard.release('ctrl')

    time.sleep(3)


def select_all_columns_then_copy():
    # Load the screenshot
    screenshotNP = np.array(pyautogui.screenshot())

    # Convert to grayscale
    gray = cv2.cvtColor(screenshotNP, cv2.COLOR_BGR2GRAY)

    # Load the template image
    template = cv2.imread('master/images/blood_test/01-alle-raekker.jpg', cv2.IMREAD_GRAYSCALE)

    # Find the text on the screenshot using template matching
    result = cv2.matchTemplate(gray, template, cv2.TM_CCOEFF_NORMED)
    y, x = np.unravel_index(result.argmax(), result.shape)

    # Click on the center of the text
    width, height = template.shape[::-1]
    pyautogui.moveTo(x + width / 2, y + height / 2)

    # Move the mouse right by 150 pixels and down by 90 pixels
    y = y + 90
    pyautogui.move(150, 50)

    # Hold down the left mouse button at the current position
    pyautogui.mouseDown(button='left')

    # Delay the click for 1 second
    time.sleep(1)

    # Move mouse to this location 
    move_mouse_on('master/images/blood_test/03-end-of-x.jpg', move_x=-3)

    # Delay the click for 1 second
    time.sleep(1)

    # Hold up the left mouse button at the current position
    pyautogui.mouseUp(button='left')

    # Simulate a right-click at the current mouse position
    pyautogui.rightClick()

    # Delay the click for 2 second
    time.sleep(2)

    # Click on the GUI element
    click_by_mouse_on('master/images/blood_test/04-copy.jpg')

    # Delay the click for 2 second
    time.sleep(2)


def extract_blood_test_data_from_columns():
    all_data = ''

    while True:
        select_all_columns_then_copy()

        # get the copied text from clipboard
        copied_text = pyperclip.paste()

        # if data already exist that mean we are in the last page
        if copied_text in all_data:
            break

        # collect the whole data in data variable
        all_data = copied_text + '\n' + all_data + '\n'

        go_one_page_back()

        time.sleep(5)

    return all_data