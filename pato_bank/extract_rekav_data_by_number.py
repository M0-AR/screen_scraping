import os

import pandas as pd
from master.global_functions import click_by_mouse_on, select_and_copy_data_from_table, save_data_to_excel
import pyautogui
import time
import win32api
from PIL import ImageGrab
import re
import pytesseract
from pytesseract import Output

# Set the path to the Tesseract OCR engine executable
# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
# Update the path if necessary

keywords = [
    "Mat.nr.	Beskrivelse af materiale/prøve",
    "Diagnoser",
    "Konklusion",
    "Mikroskopi",
    "Andre undersøgelser",
    "Makroskopi",
    "Kliniske oplysninger"
]

# Define the regular expression patterns for the numbers
patterns = [
    r'\b[0-9]{8}\b',
    r'\b[0-9]{2}[a-zA-Z]{2}[0-9]{4}\b',
    r'\b[0-9]{2}[a-zA-Z]{3}[0-9]{6}\b',
    r'\b[0-9]{10}\b',
    r'\b[0-9]{3}-[0-9]{5}\b',
]


def extract_rekv_number(text):
    import re

    global patterns
    # Iterate through the patterns and return the first match found
    for pattern in patterns:
        match = re.search(pattern, text)
        if match:
            return match.group()

    # If no matches were found, return None
    return None


def extract_and_append(text, df):
    """
    Extracts sections from the given text based on the specified keywords and appends them to the given DataFrame.

    :param text: The input text to extract sections from.
    :param df: The DataFrame to append the extracted sections to.
    :return: The updated DataFrame with the extracted sections appended.
    """
    # Add an ending keyword to stop at the beginning of the undesired text
    ending_keyword = "Dato/Tid"

    sections = {}

    global keywords
    for i, keyword in enumerate(keywords):
        start_index = text.find(keyword)
        if start_index != -1:
            start_index += len(keyword)
            if i < len(keywords) - 1:
                end_index = text.find(keywords[i + 1])
            else:
                # Find the ending_keyword if it exists in the text
                end_index = text.find(ending_keyword)
                if end_index == -1:
                    end_index = len(text)
            sections[keyword] = text[start_index:end_index].strip()

    df = df.append(sections, ignore_index=True)
    return df


def split_blocks(input_text):
    """
    Splits input_text into blocks based on the pattern of date.

    Args:
    - input_text (str): A string containing the text to split.

    Returns:
    - A list of strings containing the text blocks.
    """
    # Split the text into lines
    lines = input_text.splitlines()

    # Regular expression pattern to match dates
    date_pattern = re.compile(r'\d{2}\.\d{2}\.\d{4}')

    # Initialize a list to store the blocks
    blocks = []
    current_block = []

    # Iterate over the lines
    for line in lines:
        # If the line starts with a date, add the current block to the blocks list and start a new block
        if date_pattern.match(line):
            if current_block:
                blocks.append('\n'.join(current_block))
            current_block = [line]
        else:
            current_block.append(line)

    # Add the last block to the list
    if current_block:
        blocks.append('\n'.join(current_block))

    blocks.pop(0)
    return blocks


def check_second_line(block):
    """
    This function help the 'i' in process_block function to have
    the right length to all Blocks
    :param block:
    :return:
    """
    lines = block.split('\n')
    rekv_number = extract_rekv_numbers(lines[1])
    if rekv_number:
        # The second line is not exist, add a new line
        lines.insert(1, '\n')
    return '\n'.join(lines)


def process_block(block):
    """
    Process each block and returns a dictionary containing the processed data.

    Args:
    - block (str): A string containing the text block to process.

    Returns:
    - A dictionary containing the processed data.
    """
    block = check_second_line(block)

    lines = block.splitlines()

    i = 0
    for index, line in enumerate(lines):
        number = extract_rekv_number(line)
        if number:
            i = index
            break

    modtaget = lines[0].split('\t')[0]
    serviceyder = lines[0].split('\t')[1] + ' \n ' + '\n'.join(lines[1:2]) + ' \n ' + \
                  '\n'.join(lines[i:i + 1]).split('\t')[0]
    rekv_nr = '\n'.join(lines[i:i + 1]).split('\t')[1]
    kategori = '\n'.join(lines[i:i + 1]).split('\t')[5]
    diagnoser = '\t'.join('\n'.join(lines[i + 1:]).split('\t'))

    return {
        "Modtaget": modtaget,
        "Serviceyder": serviceyder,
        "Rekv.nr.": rekv_nr,
        "Kategori": kategori,
        "Diagnoser": diagnoser,
    }


def extract_rekv_numbers(text):
    import re

    global patterns
    # Define the regular expression patterns for the numbers
    pattern_8_digits = patterns[0]
    pattern_8_digits_letters = patterns[1]
    pattern_11_digits_letters = patterns[2]
    pattern_10_digits = patterns[3]
    pattern_9_digits_dash = patterns[4]

    # Use the patterns to search for the numbers in the text
    matches_8_digits = re.findall(pattern_8_digits, text)
    matches_8_digits_letters = re.findall(pattern_8_digits_letters, text)
    matches_11_digits_letters = re.findall(pattern_11_digits_letters, text)
    matches_10_digits = re.findall(pattern_10_digits, text)
    matches_9_digits_dash = re.findall(pattern_9_digits_dash, text)

    # Combine all matched numbers into a single list
    matches = matches_8_digits + matches_8_digits_letters + matches_11_digits_letters + matches_10_digits + matches_9_digits_dash

    return matches


import pandas as pd


def process_input_text(input_text):
    """
    Processes the input text and returns a pandas DataFrame.

    Args:
    - input_text (str): A string containing the input text.

    Returns:
    - A pandas DataFrame containing the processed data.
    """
    # Split the input text into blocks
    blocks = split_blocks(input_text)

    # Process each block and store the results in a list of dictionaries
    results = [process_block(block) for block in blocks]

    # Create a DataFrame from the list of dictionaries
    df = pd.DataFrame(results)

    return df


def sort_numbers_by_text_order(numbers, text):
    return sorted(numbers, key=lambda number: text.index(number))


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


def extract_pato_bank_data(path_to_save):
    click_by_mouse_on('images/pato_bank/01-patobank.jpg')
    # Delay the click for 1 second
    time.sleep(2)

    pato_bank_text = select_and_copy_data_from_table(up_left_corner_position='images/pato_bank/02-cgi-logo.jpg',
                                                     up_right_corner_position='images/pato_bank/03-end-of-selection.jpg',
                                                     end_scroll_position='images/pato_bank/04-scroll-down-until-sign-found.jpg',
                                                     move_y=300)

    # Create a DataFrame from the list of dictionaries
    df = process_input_text(pato_bank_text)
    #
    # Check if the file exists and remove it if it does
    if os.path.exists('file1.xlsx'):
        os.remove('file1.xlsx')
    # # Write the DataFrame to an Excel file
    df.to_excel('file1.xlsx', index=False)

    # documents_path = os.path.expanduser("~\Documents")
    # save_data_to_excel(documents_path, 'pato_bank_text', pato_bank_text)

    rekv_nrs = extract_rekv_numbers(pato_bank_text)
    sorted_rekv_nrs = sort_numbers_by_text_order(rekv_nrs, pato_bank_text)

    scroll_most_up('images/pato_bank/02-cgi-logo.jpg')
    pyautogui.click()

    time.sleep(1)

    # Create an empty DataFrame with the desired columns
    df = pd.DataFrame(columns=keywords)

    for target_number in sorted_rekv_nrs:
        # Find the center of the 05-rekv-column.jpg image
        rekv_location = pyautogui.locateCenterOnScreen('images/pato_bank/05-rekv-column.jpg', confidence=0.9)
        pyautogui.moveTo(rekv_location.x, rekv_location.y)

        found_number = False

        while not found_number:
            cursor_pos = win32api.GetCursorPos()
            screenshot, screenshot_coords = capture_screenshot(cursor_pos, radius_x=150, radius_y=1000)

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

                    # Add the extracted sections as a new row in the DataFrame
                    df = extract_and_append(rekv_number_data, df)

                    scroll_most_up('images/pato_bank/09-return-rekv-nr.jpg')
                    click_by_mouse_on('images/pato_bank/09-return-rekv-nr.jpg')
                    time.sleep(2)
                    found_number = True
                else:
                    print('Image not found on the screen. Continuing to scroll down.')
                    pyautogui.scroll(-150)
                    new_cursor_pos = (cursor_pos[0], cursor_pos[1])
                    pyautogui.moveTo(new_cursor_pos[0], new_cursor_pos[1], duration=1)
                    time.sleep(1)
            else:
                print(f'{target_number} not found on the screen. Continuing to scroll down.')
                pyautogui.scroll(-150)
                new_cursor_pos = (cursor_pos[0], cursor_pos[1])
                pyautogui.moveTo(new_cursor_pos[0], new_cursor_pos[1], duration=1)
                time.sleep(1)

    # Check if the file exists and remove it if it does
    if os.path.exists('file2.xlsx'):
        os.remove('file2.xlsx')
    df.to_excel('file2.xlsx', index=False)

    # Read the Excel files
    file1 = pd.read_excel('file1.xlsx')
    file2 = pd.read_excel('file2.xlsx')

    # Concatenate the DataFrames
    df = pd.concat([file1, file2], ignore_index=True)

    # Move non-empty values to the top for all columns
    for column in df.columns:
        df[column] = df[column].replace('', pd.NA)  # Replace empty strings with NA values
        non_empty_rows = df[column].dropna().reset_index(drop=True)  # Get non-empty rows
        empty_rows = df[column].isna().sum()  # Count empty rows
        df[column] = pd.concat([non_empty_rows, pd.Series([pd.NA] * empty_rows, dtype='object')], ignore_index=True)

    # Delete all the data after row 11
    df = df.iloc[:11]  # TODO len(sorted_rekv_nrs) - 1

    # Write the modified DataFrame to a new Excel file
    df.to_excel(path_to_save + '/pato_bank.xlsx', index=False)
