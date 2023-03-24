import time
import pyautogui
import pyperclip

from master.global_functions import click_by_mouse_on, press_ctrl_c, save_data_to_excel


def extract_miba_data(patient_path):
    # Wait for 1 second
    time.sleep(1)

    # Click on the 'MIBA' button
    click_by_mouse_on('images/miba/01-miba.jpg')

    # Wait for 1 second
    time.sleep(3)

    # Click on the 'All results' button
    click_by_mouse_on('images/miba/02-all-result.jpg', confidence=0.7)

    # Wait for 1 second
    time.sleep(1)

    # Click on
    click_by_mouse_on('images/miba/03-select-all.jpg', -50, 0)

    # Wait for 1 second
    time.sleep(1)

    # Click on
    click_by_mouse_on('images/miba/04-display-all-result.jpg')

    # Wait for 1 second
    time.sleep(1)

    # Click on
    click_by_mouse_on('images/miba/05-up-left-corner.jpg')

    # Wait for 1 second
    time.sleep(1)

    # Simulate a right-click at the current mouse position
    pyautogui.rightClick()
    # Delay the click for 2 second
    time.sleep(2)

    # Click on
    click_by_mouse_on('images/miba/06-select-all.jpg')
    # Wait for 1 second
    time.sleep(1)

    # Copy the selected data to the clipboard
    press_ctrl_c()

    # Get the text from the clipboard and return it
    all_data = pyperclip.paste()

    save_data_to_excel(patient_path, 'miba', all_data)
    # Delay the click for 3 second
    time.sleep(3)

    # Click on the 'x' button to close miba
    click_by_mouse_on('images/miba/08-close-miba.jpg')

    # Delay the click for 3 second
    time.sleep(3)
