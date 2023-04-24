import os
import re
import time

import pandas as pd
import pyautogui
import pyperclip
from master.global_functions import click_by_mouse_on, move_mouse_on, save_data_to_excel, write_dataframe_to_excel


def main_extract_medicin_data(path_to_save_data):
    """
    Extracts medicine data from an application, copies the data to the clipboard,
    and saves it to an Excel file.

    Args:
        path_to_save_data (str): The path where the Excel file will be saved.
    """
    try:
        # Click on the first image and wait for 2 seconds
        click_by_mouse_on('images/medicin/01-medicin.jpg')
        time.sleep(2)

        # Click on the second image and wait for 2 seconds
        click_by_mouse_on('images/medicin/02-unselect.jpg')
        time.sleep(2)

        # Click on the third image, move 60 pixels on the x-axis, and wait for 1 second
        click_by_mouse_on('images/medicin/03-flere.jpg', move_x=60, confidence=0.95)
        time.sleep(1)

        # Click on the fourth image and wait for 1 second
        click_by_mouse_on('images/medicin/04-select-all.jpg', confidence=0.95)
        time.sleep(1)

        # Move the mouse to the position of the fifth image and wait for 1 second
        move_mouse_on('images/medicin/05-right-click.jpg')
        time.sleep(1)

        # Simulate a right-click at the current mouse position and wait for 2 seconds
        pyautogui.rightClick()
        time.sleep(2)

        # Click on the sixth image and wait for 1 second
        click_by_mouse_on('images/medicin/06-show-result-in-side-panel.jpg', confidence=0.90)
        time.sleep(2)

        # Click on the seventh image and wait for 15 seconds
        click_by_mouse_on('images/medicin/07-copy-all.jpg', confidence=0.90)
        time.sleep(15)

        # Get the text from the clipboard and save it to an Excel file
        medicin_data = pyperclip.paste()

        # using re.DOTALL flag to enable matching across multiple lines
        regex = r"Receptdetaljer.*?Effektueringsdetaljer"

        # extract all matches_text of the pattern
        matches_text = re.findall(regex, medicin_data, re.DOTALL)

        # extract medication name and date
        medications = []
        for text in matches_text:
            text_list = text.split("\n")
            for line in text_list:
                match = re.search(r'^(.*?)(\d{2}-\d{2}-\d{4})', line)
                if match:
                    med_name = match.group(1).strip()
                    med_date = match.group(2)
                    medications.append({"Medication": med_name, "Date": med_date})

        # create dataframe from medications list
        df = pd.DataFrame(medications)

        write_dataframe_to_excel(df, os.path.join(path_to_save_data, 'medicin.xlsx'))
        # save_data_to_excel(path_to_save_data, 'medicin', medicin_data, simulate_ctrl_v=True)
    except Exception as e:
        print(f"Failed to extract medicin data: {e} {path_to_save_data}")