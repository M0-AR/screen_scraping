from master.global_functions import click_by_mouse_on
import pyautogui
import time

def extract_rekv_numbers(text):
    import re

    # Define the regular expression pattern for the numbers
    pattern = r'\b[0-9]{8}\b'

    # Use the pattern to search for the numbers in the text
    matches = re.findall(pattern, text)

    # Return the matched numbers as a list
    return matches


def select_and_copy_data_from_table_with_scroll():
    pass


# Move the mouse to image position
click_by_mouse_on('images/pato_bank/02-cgi-logo.jpg', 0, 300)

# Delay for 1 second
time.sleep(1)

# Get the position of the 03-end-of-selection.jpg image
x, y = pyautogui.locateCenterOnScreen('images/pato_bank/03-end-of-selection.jpg', confidence=0.9)

# Move the mouse to the 03-end-of-selection.jpg image position while holding down the left mouse button
pyautogui.mouseDown(button='left')
pyautogui.moveTo(x, y, duration=1)

# Scroll down until red-line found
while True:
    # Check if the red line image is on the screen
    if pyautogui.locateOnScreen('images/pato_bank/04-scroll-down-until-sign-found.jpg', confidence=0.9):
        break  # exit the loop if the image is found
    else:
        # Scroll down by 100 pixels
        pyautogui.scroll(-100)

# Get the current position of the mouse cursor
x, _ = pyautogui.position()

# Get the position of the 04-scroll-down-until-sign-found.jpg image
_, y = pyautogui.locateCenterOnScreen('images/pato_bank/04-scroll-down-until-sign-found.jpg', confidence=0.9)

pyautogui.moveTo(x, y, duration=1)

pyautogui.mouseUp(button='left')

# rekv_numbers = extract_rekv_numbers()