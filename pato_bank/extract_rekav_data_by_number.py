from master.global_functions import click_by_mouse_on, select_and_copy_data_from_table
import pyautogui
import time
import win32api
from PIL import ImageGrab
import pytesseract
from pytesseract import Output
import re

# Set the path to the Tesseract OCR engine executable
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"


# Update the path if necessary

def extract_rekv_numbers(text):
    import re

    # Define the regular expression pattern for the numbers
    pattern = r'\b[0-9]{8}\b'

    # Use the pattern to search for the numbers in the text
    matches = re.findall(pattern, text)

    # Return the matched numbers as a list
    return matches


def scroll_most_up(sign_to_end_scroll):
    # Scroll up until the image is found
    while not pyautogui.locateOnScreen(sign_to_end_scroll, confidence=0.9):
        pyautogui.scroll(1000)
        time.sleep(0.5)


def capture_screenshot(cursor_pos, radius_x=50, radius_y=50):
    x, y = cursor_pos
    left = max(0, x - radius_x)
    top = max(0, y - radius_y)
    right = x + radius_x
    bottom = y + radius_y

    screenshot = ImageGrab.grab(bbox=(left, top, right, bottom))
    return screenshot, (left, top, right, bottom)


def find_number_in_screenshot(pattern, screenshot):
    data = pytesseract.image_to_data(screenshot, output_type=Output.DICT)
    n_boxes = len(data['level'])
    for i in range(n_boxes):
        text = data['text'][i]
        if re.fullmatch(pattern, text):
            x, y, w, h = data['left'][i], data['top'][i], data['width'][i], data['height'][i]
            return text, (x, y, w, h)
    return None, None


# click_by_mouse_on('images/pato_bank/01-patobank.jpg')
# # Delay the click for 1 second
# time.sleep(2)

# pato_bank_text = select_and_copy_data_from_table(up_left_corner_position='images/pato_bank/02-cgi-logo.jpg',
#                                                  up_right_corner_position='images/pato_bank/03-end-of-selection.jpg',
#                                                  end_scroll_position='images/pato_bank/04-scroll-down-until-sign-found.jpg',
#                                                  move_y=300)
#
# rekv_nrs = extract_rekv_numbers(pato_bank_text)
# print(rekv_nrs)
#
# scroll_most_up('images/pato_bank/02-cgi-logo.jpg')
# pyautogui.click()

rekv_nrs = ['17750558', '16022811', '16020762', '03306002', '00303258', '97307229', '95311121', '94315488', '93321478',
            '91308865', '87006228']

time.sleep(1)

for target_number in rekv_nrs:
    target_number = '87006228'
    # Find the center of the 05-rekv-column.jpg image
    rekv_location = pyautogui.locateCenterOnScreen('images/pato_bank/05-rekv-column.jpg', confidence=0.9)
    pyautogui.moveTo(rekv_location.x, rekv_location.y)

    found_number = False

    while not found_number:
        cursor_pos = win32api.GetCursorPos()
        screenshot, screenshot_coords = capture_screenshot(cursor_pos, radius_x=100, radius_y=1000)

        rekv_number, number_bbox = find_number_in_screenshot(target_number, screenshot)

        if rekv_number is not None:
            print(f'Found the number {rekv_number}')
            if number_bbox is not None:
                x, y, w, h = number_bbox
                screenshot_x, screenshot_y = screenshot_coords[:2]
                adjusted_x, adjusted_y = x + w // 2 + screenshot_x, y + h // 2 + screenshot_y
                pyautogui.moveTo(adjusted_x, adjusted_y, duration=1)
                pyautogui.click(adjusted_x, adjusted_y)

                time.sleep(2)
                rekv_number_data = select_and_copy_data_from_table(
                    up_left_corner_position='images/pato_bank/06-top_left_selection-rekv-nr.jpg',
                    up_right_corner_position='images/pato_bank/07-top_right_selection-rekv-nr.jpg',
                    end_scroll_position='images/pato_bank/08-end-of-scroll-rekv-nr.jpg')

                print(rekv_number_data)
                scroll_most_up('images/pato_bank/09-return-rekv-nr.jpg')
                click_by_mouse_on('images/pato_bank/09-return-rekv-nr.jpg')
                time.sleep(2)
                found_number = True
            else:
                print('Image not found on the screen. Continuing to scroll down.')
                pyautogui.scroll(-75)
                new_cursor_pos = (cursor_pos[0], cursor_pos[1])
                pyautogui.moveTo(new_cursor_pos[0], new_cursor_pos[1], duration=1)
                time.sleep(1)
        else:
            print('Number not found on the screen. Continuing to scroll down.')
            pyautogui.scroll(-75)
            new_cursor_pos = (cursor_pos[0], cursor_pos[1])
            pyautogui.moveTo(new_cursor_pos[0], new_cursor_pos[1], duration=1)
            time.sleep(1)
