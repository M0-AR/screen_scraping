import numpy as np
import pyperclip
import pyautogui
import time
import keyboard
from master.global_functions import click_by_mouse_on, select_and_copy_data_from_table


def extract_rekv_numbers(text):
    import re

    # Define the regular expression pattern for the numbers
    pattern = r'\b[0-9]{8}\b'

    # Use the pattern to search for the numbers in the text
    matches = re.findall(pattern, text)

    # Return the matched numbers as a list
    return matches


# def select_and_copy_data_from_table(top_left_selection,top_right_selection, move_x=0, move_y=0):
#     # Move the mouse to image position
#     click_by_mouse_on(top_left_selection, move_x, move_y, confidence=0.9)
#
#     # Delay for 1 second
#     time.sleep(1)
#
#     # Get the position of the 03-end-of-selection.jpg image
#     x, y = pyautogui.locateCenterOnScreen(top_right_selection, confidence=0.9)
#
#     # Move the mouse to the 03-end-of-selection.jpg image position while holding down the left mouse button
#     pyautogui.mouseDown(button='left')
#     pyautogui.moveTo(x, y, duration=1)
#
#     scroll_most_down()
#
#     # Get the current position of the mouse cursor
#     x, _ = pyautogui.position()
#
#     # Get the position of the 04-08-scroll-down-until-sign-found.jpg image
#     _, y = pyautogui.locateCenterOnScreen('images/pato_bank/04-scroll-down-until-sign-found.jpg', confidence=0.9)
#
#     pyautogui.moveTo(x, y, duration=1)
#
#     pyautogui.mouseUp(button='left')
#
#     # Press and hold the Ctrl key
#     keyboard.press('ctrl')
#
#     # Press and release the C key
#     keyboard.press('c')
#     keyboard.release('c')
#
#     # Release the Ctrl key
#     keyboard.release('ctrl')
#
#     # Delay for 1 second
#     time.sleep(1)
#
#     # Get the text from the clipboard
#     data = pyperclip.paste()
#
#     # Delay for 1 second
#     time.sleep(1)
#     return data


def scroll_most_down():
    # Scroll down until red-line found
    while True:
        # Check if the red line image is on the screen
        if pyautogui.locateOnScreen('images/pato_bank/04-scroll-down-until-sign-found.jpg', confidence=0.9):
            break  # exit the loop if the image is found
        else:
            # Scroll down by 100 pixels
            pyautogui.scroll(-100)


def scroll_most_up():
    # Scroll up until the image is found
    while not pyautogui.locateOnScreen('images/pato_bank/03-end-of-selection.jpg', confidence=0.9):
        pyautogui.scroll(1000)
        time.sleep(0.5)


rekv_number_data = select_and_copy_data_from_table(up_left_corner_position='images/pato_bank/06-top_left_selection-rekv-nr.jpg',
                                                   up_right_corner_position='images/pato_bank/07-top_right_selection-rekv-nr.jpg',
                                                   end_scroll_position='images/pato_bank/08-end-of-scroll-rekv-nr.jpg')
print(rekv_number_data)
# click_by_mouse_on('images/pato_bank/01-patobank.jpg')
# # Delay the click for 1 second
# time.sleep(2)

# pato_bank_text = select_and_copy_data_from_table(top_left_selection='images/pato_bank/02-cgi-logo.jpg',
#                                                  top_right_selection='images/pato_bank/03-end-of-selection.jpg',
#                                                  move_y=300)

# rekv_nrs = extract_rekv_numbers(pato_bank_text)
# print(rekv_numbers)
#
# scroll_most_up()

# --------- Locate the Rekv.nr then click on it ------#
# Search for the image file
# rekv_number_location = pyautogui.locateOnScreen('number.png', confidence=0.90)
#
# if rekv_number_location is not None:
#     print('Found the image at: ', rekv_number_location)
#     x, y = pyautogui.center(rekv_number_location)
#     # pyautogui.click(x, y)
#     pyautogui.moveTo(x, y, duration=1)
#
# else:
#     print('Image not found on the screen.')
# ------------------------------------------------------#

# import pyautogui
#
# # Find the center of the rekv.jpg image
# rekv_location = pyautogui.locateCenterOnScreen('images/pato_bank/05-rekv-column.jpg', confidence=0.9)
#
# if rekv_location is not None:
#     x = rekv_location.x
#     y = rekv_location.y
#
#     # Scroll until the rekv number is on screen
#     pyautogui.moveTo(x, y)
#     pyautogui.scroll(-100)
#
#     # Define the search region around the center of the rekv.jpg image
#     search_region = (x - 500, y + 500, 1000, 1000)  # left, top, width, height
#
#     # Find the location of the number and click it
#     number_location = pyautogui.locateOnScreen('number.png',
#                                                region=search_region, confidence=0.9)
#     if number_location is not None:
#         number_center = pyautogui.center(number_location)
#         pyautogui.moveTo(number_center.x, number_center.y, duration=1)
#         # pyautogui.click(number_center)
#     else:
#         print("Number not found.")
# else:
#     print("Rekv image not found.")


# import time
# import win32api
# from PIL import ImageGrab
#
#
# def capture_screenshot(cursor_pos, radius_x=50, radius_y=50):
#     x, y = cursor_pos
#     left = max(0, x - radius_x)
#     top = max(0, y - radius_y)
#     right = x + radius_x
#     bottom = y + radius_y
#
#     screenshot = ImageGrab.grab(bbox=(left, top, right, bottom))
#     screenshot.save("number.png")
#
#
# while True:
#     cursor_pos = win32api.GetCursorPos()
#     capture_screenshot(cursor_pos, radius_x=100)
#     time.sleep(5)


# import time
# from pywinauto import Desktop
# from pywinauto.keyboard import send_keys
#
# desktop = Desktop(backend="uia")
#
# numbers = ['17750558', '16022811', '16020762', '03306002', '00303258', '97307229', '95311121', '94315488', '93321478', '91308865', '87006228']
#
# for number in numbers:
#     try:
#         element = desktop.window(title=number, control_type="Text")
#         if element.exists():
#             element.click()
#             time.sleep(1)
#         else:
#             print(f"Number {number} not found. Scrolling down.")
#             send_keys("{PGDN}")  # Send Page Down key
#             time.sleep(1)
#     except Exception as e:
#         print(f"Error: {e}")


# import pyautogui
# import time
# import win32api
# from PIL import ImageGrab
# import pytesseract
# import re
#
# # Set the path to the Tesseract OCR engine executable
# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
# # Update the path if necessary
#
# def capture_screenshot(cursor_pos, radius_x=50, radius_y=50):
#     x, y = cursor_pos
#     left = max(0, x - radius_x)
#     top = max(0, y - radius_y)
#     right = x + radius_x
#     bottom = y + radius_y
#
#     screenshot = ImageGrab.grab(bbox=(left, top, right, bottom))
#     return screenshot
#
# def find_number_in_screenshot(pattern, screenshot):
#     text = pytesseract.image_to_string(screenshot)
#     matches = re.findall(pattern, text)
#     return matches[0] if matches else None
#
# # Find the center of the 05-rekv-column.jpg image
# rekv_location = pyautogui.locateCenterOnScreen('images/pato_bank/05-rekv-column.jpg', confidence=0.9)
#
# if rekv_location is not None:
#     x = rekv_location.x
#     y = rekv_location.y
#
#     # Scroll until the rekv number is on screen
#     pyautogui.moveTo(x, y)
#     pyautogui.scroll(-100)
#
#     # Define the search region around the center of the rekv.jpg image
#     search_region = (x - 500, y + 500, 1000, 1000)  # left, top, width, height
#
# # Define the regular expression pattern for the numbers
# target_number = '17750558'
# # pattern = r'\b[0-9]{8}\b'
# found_number = False
#
# while not found_number:
#     cursor_pos = win32api.GetCursorPos()
#     screenshot = capture_screenshot(cursor_pos, radius_x=100)
#     screenshot.save('number.png')
#
#     # rekv_number = find_number_in_screenshot(pattern, screenshot)
#     rekv_number = find_number_in_screenshot(target_number, screenshot)
#
#     if rekv_number is not None:
#         print(f'Found the number {rekv_number}')
#         rekv_number_location = pyautogui.locateOnScreen('number.png', confidence=0.9)
#
#         if rekv_number_location is not None:
#             x, y = pyautogui.center(rekv_number_location)
#             pyautogui.moveTo(x, y, duration=1)
#             pyautogui.click(x, y)
#             rekv_number_data = select_and_copy_data_from_table(top_left_selection='images/pato_bank/06-top_left_selection-rekv-nr.jpg',
#                                                                top_right_selection='images/pato_bank/07-top_right_selection-rekv-nr.jpg')
#             print(rekv_number_data)
#             found_number = True
#         else:
#             print('Image not found on the screen. Continuing to scroll down.')
#             pyautogui.scroll(-3)
#             new_cursor_pos = (cursor_pos[0], cursor_pos[1] + 50)
#             pyautogui.moveTo(new_cursor_pos[0], new_cursor_pos[1], duration=1)
#             time.sleep(1)
#     else:
#         print('Number not found on the screen. Continuing to scroll down.')
#         pyautogui.scroll(-3)
#         new_cursor_pos = (cursor_pos[0], cursor_pos[1] + 50)
#         pyautogui.moveTo(new_cursor_pos[0], new_cursor_pos[1], duration=1)
#         time.sleep(1)
