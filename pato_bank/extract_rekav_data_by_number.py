import numpy as np
import pyperclip
import pyautogui
import time
import keyboard
from master.global_functions import click_by_mouse_on


def extract_rekv_numbers(text):
    import re

    # Define the regular expression pattern for the numbers
    pattern = r'\b[0-9]{8}\b'

    # Use the pattern to search for the numbers in the text
    matches = re.findall(pattern, text)

    # Return the matched numbers as a list
    return matches


def select_and_copy_data_from_table():
    # Move the mouse to image position
    click_by_mouse_on('images/pato_bank/02-cgi-logo.jpg', 0, 300)

    # Delay for 1 second
    time.sleep(1)

    # Get the position of the 03-end-of-selection.jpg image
    x, y = pyautogui.locateCenterOnScreen('images/pato_bank/03-end-of-selection.jpg', confidence=0.9)

    # Move the mouse to the 03-end-of-selection.jpg image position while holding down the left mouse button
    pyautogui.mouseDown(button='left')
    pyautogui.moveTo(x, y, duration=1)

    scroll_most_down()

    # Get the current position of the mouse cursor
    x, _ = pyautogui.position()

    # Get the position of the 04-08-scroll-down-until-sign-found.jpg image
    _, y = pyautogui.locateCenterOnScreen('images/pato_bank/04-08-scroll-down-until-sign-found.jpg', confidence=0.9)

    pyautogui.moveTo(x, y, duration=1)

    pyautogui.mouseUp(button='left')

    # Press and hold the Ctrl key
    keyboard.press('ctrl')

    # Press and release the C key
    keyboard.press('c')
    keyboard.release('c')

    # Release the Ctrl key
    keyboard.release('ctrl')

    # Delay for 1 second
    time.sleep(1)

    # Get the text from the clipboard
    pato_bank_text = pyperclip.paste()

    # Delay for 1 second
    time.sleep(1)

    rekv_nrs = extract_rekv_numbers(pato_bank_text)
    return rekv_nrs


def scroll_most_down():
    # Scroll down until red-line found
    while True:
        # Check if the red line image is on the screen
        if pyautogui.locateOnScreen('images/pato_bank/04-08-scroll-down-until-sign-found.jpg', confidence=0.9):
            break  # exit the loop if the image is found
        else:
            # Scroll down by 100 pixels
            pyautogui.scroll(-100)


def scroll_most_up():
    # Scroll up until the image is found
    while not pyautogui.locateOnScreen('images/pato_bank/03-end-of-selection.jpg', confidence=0.9):
        pyautogui.scroll(1000)
        time.sleep(0.5)


# click_by_mouse_on('images/pato_bank/01-patobank.jpg')
# # Delay the click for 1 second
# time.sleep(2)

# rekv_numbers = select_and_copy_data_from_table()
# print(rekv_numbers)

# scroll_most_up()

# from PIL import Image, ImageDraw, ImageFont
# import pyautogui
#
# # Convert number to image
# number = "20016491"
# image = Image.new("RGB", (100, 50), color=(255, 255, 255))
# draw = ImageDraw.Draw(image)
# font = ImageFont.truetype("arial.ttf", 20)
# draw.text((10, 10), number, font=font, fill=(0, 0, 0))
# image.save("number.png")

# Locate and click on image
# position = pyautogui.locateOnScreen("number.png")
# if position is not None:
#     x, y = pyautogui.center(position)
#     pyautogui.click(x, y)
# else:
#     print("Image not found on screen")




# Two solutions text or image in order to click it when find it
# import time
# import cv2
# import numpy as np
#
# text = "20016491"
# font = cv2.FONT_HERSHEY_SIMPLEX
# font_scale = 1
# thickness = 2
# size = cv2.getTextSize(text, font, font_scale, thickness)[0]
# width = size[0]
# height = size[1] * 2  # Double the height of the image
# image = np.zeros((height, width, 3), dtype=np.uint8)
# cv2.putText(image, text, (0, height - size[1]), font, font_scale, (255, 255, 255), thickness)
# cv2.imwrite("number.png", image)

# # Search for the image file
# rekv_number_location = pyautogui.locateOnScreen('number.png', confidence=0.30)
#
# if rekv_number_location is not None:
#     print('Found the image at: ', rekv_number_location)
#     x, y = pyautogui.center(rekv_number_location)
#     # pyautogui.click(x, y)
#     pyautogui.moveTo(x, y, duration=1)
#
# else:
#     print('Image not found on the screen.')

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
#                                                region=search_region, confidence=0.4)
#     if number_location is not None:
#         number_center = pyautogui.center(number_location)
#         pyautogui.moveTo(number_center.x, number_center.y, duration=1)
#         # pyautogui.click(number_center)
#     else:
#         print("Number not found.")
# else:
#     print("Rekv image not found.")

import pytesseract
import cv2

# Take a screenshot and save it as 'screenshot.png'
# screenshot = pyautogui.screenshot('screenshot.png') # TODO: uncomment

# Load the image
image = cv2.imread('screenshot.png')

# Convert the image to grayscale
gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

# Apply thresholding to make the text more visible
thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV | cv2.THRESH_OTSU)[1]

# Use pytesseract to extract the text from the image
text = pytesseract.image_to_string(thresh)

# Print the extracted text
print(text)

# Search for the image of the number '20016491' in the screenshot and get its position
number_location = pyautogui.locateOnScreen('number.png', grayscale=True, confidence=0.9)

# Get the center of the number image
x, y = pyautogui.center(number_location)

# Click on the center of the number image
pyautogui.moveTo(x, y, duration=1)

# # Get the position of the text "20016491" in the screenshot
# position = pyautogui.locateOnScreen('screenshot.png', grayscale=True, confidence=0.9, text='20016491')
#
# pyautogui.moveTo(position.x, position.y, duration=1)
#
# # Print the position
# print(position)