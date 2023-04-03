import os
import re
import time
import pandas as pd
import pyautogui
import pytesseract
from pytesseract import Output
from PIL import ImageGrab
import win32api
from typing import Tuple

from master.global_functions import click_by_mouse_on, select_and_copy_data_from_table

# Set the path to the Tesseract OCR engine executable
# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
# Update the path if necessary

REKV_NUM_PATTERNS = [
    r'\b[0-9]{8}\b',
    r'\b[0-9]{2}[a-zA-Z]{2}[0-9]{4}\b',
    r'\b[0-9]{2}[a-zA-Z]{3}[0-9]{6}\b',
    r'\b[0-9]{10}\b',
    r'\b[0-9]{3}-[0-9]{5}\b',
]

KEYWORDS = [
    "Mat.nr.	Beskrivelse af materiale/prøve",
    "Diagnoser",
    "Konklusion",
    "Mikroskopi",
    "Andre undersøgelser",
    "Makroskopi",
    "Kliniske oplysninger"
]


def extract_rekv_number(text):
    """Extract the REKV number from the given text.

    Args:
        text (str): The text to search for a REKV number.

    Returns:
        str: The first REKV number found, or None if no match is found.
    """
    for pattern in REKV_NUM_PATTERNS:
        match = re.search(pattern, text)
        if match:
            return match.group()
    return None


def extract_sections(text):
    """
    Extracts sections from the given text based on the specified keywords.

    :param text: The input text to extract sections from.
    :return: A dictionary containing the extracted sections with keywords as keys.
    """
    # Add an ending keyword to stop at the beginning of the undesired text
    ending_keyword = "Dato/Tid"

    sections = {}

    for i, keyword in enumerate(KEYWORDS):
        start_index = text.find(keyword)
        if start_index != -1:
            start_index += len(keyword)
            if i < len(KEYWORDS) - 1:
                end_index = text.find(KEYWORDS[i + 1])
            else:
                # Find the ending_keyword if it exists in the text
                end_index = text.find(ending_keyword)
                if end_index == -1:
                    end_index = len(text)
            sections[keyword] = text[start_index:end_index].strip()

    return sections


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


def ensure_second_line(block: str) -> str:
    """
    Ensures that the second line of the given block has a newline character
    to help the 'i' in process_block function to have the right length to all Blocks.

    :param block: The block to be checked.
    :return: The block with the second line updated if necessary.
    """
    block_lines = block.split('\n')
    extracted_numbers = extract_rekv_numbers_from_text(block_lines[1])
    if extracted_numbers is not None:
        # If the second line doesn't contain a number, insert a newline character
        block_lines.insert(1, '\n')
    return '\n'.join(block_lines)


def process_block(block):
    """
    Process a block of text and return a dictionary containing the processed data.

    Args:
    - block (str): A string containing the text block to process.

    Returns:
    - A dictionary containing the processed data.
    """
    block = ensure_second_line(block)

    block_lines = block.splitlines()

    i = 0
    for index, line in enumerate(block_lines):
        number = extract_rekv_number(line)
        if number:
            i = index
            break

    modtaget = block_lines[0].split('\t')[0]
    serviceyder = block_lines[0].split('\t')[1] + ' \n ' + '\n'.join(block_lines[1:2]) + ' \n ' + \
                  '\n'.join(block_lines[i:i + 1]).split('\t')[0]
    rekv_nr = '\n'.join(block_lines[i:i + 1]).split('\t')[1]
    kategori = '\n'.join(block_lines[i:i + 1]).split('\t')[5]
    diagnoser = '\t'.join('\n'.join(block_lines[i + 1:]).split('\t'))

    return {
        "Modtaget": modtaget,
        "Serviceyder": serviceyder,
        "Rekv.nr.": rekv_nr,
        "Kategori": kategori,
        "Diagnoser": diagnoser,
    }


def extract_rekv_numbers_from_text(text):
    """
    Extract all the REKV numbers from the given text.

    Args:
    - text (str): The text to search for REKV numbers.

    Returns:
    - A list of all the REKV numbers found in the text.
    """
    import re

    # Use the patterns to search for the numbers in the text
    matches = [match.group() for pattern in REKV_NUM_PATTERNS for match in re.finditer(pattern, text)]

    return matches


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
    """
    Sorts the given list of numbers based on the order in which they appear in the given text.

    Args:
        numbers (List[str]): A list of numbers to sort.
        text (str): The text to search for the numbers.

    Returns:
        A new list of the same numbers sorted based on their appearance in the text.
    """
    return sorted(numbers, key=lambda number: text.index(number))


def capture_screenshot(cursor_position: Tuple[int, int], width: int = 100, height: int = 100):
    """
    Captures a screenshot around the given cursor position with the specified width and height.

    Args:
        cursor_position: A tuple containing the (x, y) coordinates of the cursor position.
        width: The width of the screenshot to capture.
        height: The height of the screenshot to capture.

    Returns:
        A tuple containing the captured screenshot image and the bounding box of the captured area.
    """
    x, y = cursor_position
    left = max(0, x - width)
    top = max(0, y - height)
    right = x + width
    bottom = y + height

    screenshot_image = ImageGrab.grab(bbox=(left, top, right, bottom))
    return screenshot_image, (left, top, right, bottom)


def find_number_in_screenshot(pattern, screenshot):
    """
    Finds the first occurrence of a number matching the given pattern in a screenshot.

    Args:
    - pattern (str): A regular expression pattern to match against the number.
    - screenshot (PIL.Image): An image of the screenshot to search.

    Returns:
    - A tuple containing the text of the matched number and the bounding box coordinates of the number.
    - If no match is found, returns (None, None).
    """
    data = pytesseract.image_to_data(screenshot, output_type=Output.DICT)
    n_boxes = len(data['level'])
    for i in range(n_boxes):
        text = data['text'][i]
        if re.fullmatch(pattern, text):
            x, y, w, h = data['left'][i], data['top'][i], data['width'][i], data['height'][i]
            return text, (x, y, w, h)
    return None, None


def click_pato_bank():
    """
    Click on the Pato Bank logo to start the process.
    """
    click_by_mouse_on('images/pato_bank/01-patobank.jpg')
    time.sleep(2)


def extract_table_text():
    """
    Extract text from the Pato Bank table.

    Returns:
        str: Extracted text from the table.
    """
    return select_and_copy_data_from_table(
        up_left_corner_position='images/pato_bank/02-cgi-logo.jpg',
        up_right_corner_position='images/pato_bank/03-end-of-selection.jpg',
        end_scroll_position='images/pato_bank/04-scroll-down-until-sign-found.jpg',
        move_y=300
    )


def create_dataframe_from_text(text):
    """
    Create a DataFrame from the extracted text.

    Args:
        text (str): Extracted text from the table.

    Returns:
        pd.DataFrame: DataFrame created from the extracted text.
    """
    return process_input_text(text)


def remove_existing_file(file_path):
    """
    Remove existing file if it exists.

    Args:
        file_path (str): The path to the file that should be removed.
    """
    if os.path.exists(file_path):
        os.remove(file_path)


def write_dataframe_to_excel(df, file_path, index=False):
    """
    Write the DataFrame to an Excel file.

    Args:
        df (pd.DataFrame): DataFrame to be written to the Excel file.
        file_path (str): The path to the file where the DataFrame should be written.
        index (bool, optional): Whether to write row names (index). Defaults to False.
    """
    df.to_excel(file_path, index=index)


def extract_and_sort_rekv_numbers(text):
    """
    Extract and sort REKV numbers from the text.

    Args:
        text (str): Extracted text from the table.

    Returns:
        list: Sorted REKV numbers.
    """
    rekv_nrs = extract_rekv_numbers_from_text(text)
    return sort_numbers_by_text_order(rekv_nrs, text)


def create_empty_dataframe(columns):
    """
    Create an empty DataFrame with the desired columns.

    Args:
        columns (list): A list of column names for the DataFrame.

    Returns:
        pd.DataFrame: An empty DataFrame with the desired columns.
    """
    return pd.DataFrame(columns=columns)


def find_and_move_to_rekv_column():
    """
    Find the center of the 05-rekv-column.jpg image.

    Returns:
        tuple: (x, y) coordinates of the center of the 05-rekv-column.jpg image.
    """
    rekv_location = pyautogui.locateCenterOnScreen('images/pato_bank/05-rekv-column.jpg', confidence=0.9)
    pyautogui.moveTo(rekv_location.x, rekv_location.y)
    return rekv_location


def handle_rekv_number(target_number, cursor_pos, screenshot, screenshot_coords):
    """
    Handle the target REKV number by clicking on it and extracting the corresponding data from the table.

    Args:
        target_number (int): The target REKV number.
        cursor_pos (tuple): The current cursor position (x, y).
        screenshot (Image): The screenshot taken around the cursor position.
        screenshot_coords (tuple): The (x, y) coordinates of the top-left corner of the screenshot.

    Returns:
        pd.DataFrame: DataFrame containing the extracted data for the target REKV number.
    """
    rekv_number, number_bbox = find_number_in_screenshot(target_number, screenshot)
    if rekv_number is not None:
        print(f'Found the number {rekv_number}')
        if number_bbox is not None:
            # Click on the target number and extract the corresponding data from the table
            x, y, w, h = number_bbox
            screenshot_x, screenshot_y = screenshot_coords[:2]
            adjusted_x, adjusted_y = x + w // 2 + screenshot_x, y + h // 2 + screenshot_y
            pyautogui.moveTo(adjusted_x, adjusted_y, duration=1)
            pyautogui.click(adjusted_x, adjusted_y)

            time.sleep(2)
            rekv_number_data = select_and_copy_data_from_table(
                up_left_corner_position='images/pato_bank/06-top_left_selection-rekv-nr.jpg',
                up_right_corner_position='images/pato_bank/07-top-right-selection-rekv-nr.jpg',
                end_scroll_position='images/pato_bank/08-end-of-scroll-rekv-nr.jpg'
            )

            # Extract data as a new row in the DataFrame
            return extract_sections(rekv_number_data)

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
    return None


def clean_and_organize_dataframe(df, rekv_nrs):
    """
    Clean and organize the DataFrame by moving non-empty values to the top, deleting noisy data,
    and replacing instances of _x000D_ with a space.

    Args:
        df (pd.DataFrame): The DataFrame to be cleaned and organized.
        sorted_rekv_nrs (list): A list of sorted REKV numbers.

    Returns:
        pd.DataFrame: The cleaned and organized DataFrame.
    """

    # Move non-empty values to the top for all columns
    for column in df.columns:
        df[column] = df[column].replace('', pd.NA)  # Replace empty strings with NA values
        non_empty_rows = df[column].dropna().reset_index(drop=True)  # Get non-empty rows
        empty_rows = df[column].isna().sum()  # Count empty rows
        df[column] = pd.concat([non_empty_rows, pd.Series([pd.NA] * empty_rows, dtype='object')],
                               ignore_index=True)

    # Delete all noisy data
    df = df.iloc[:len(rekv_nrs)]

    # Replace all instances of _x000D_ with a space in all cells of the dataframe df
    df = df.replace('_x000D_', ' ', regex=True)

    return df


def main_extract_pato_bank_data(path_to_save):
    """
    Main function to extract data from Pato Bank and save it to an Excel file.
    Args:
    path_to_save (str): The path to the directory where the resulting Excel file should be saved.
    """

    click_pato_bank()
    pato_bank_text = extract_table_text()
    df = create_dataframe_from_text(pato_bank_text)

    remove_existing_file('file1.xlsx')
    write_dataframe_to_excel(df, 'file1.xlsx')

    sorted_rekv_nrs = extract_and_sort_rekv_numbers(pato_bank_text)

    click_pato_bank()

    df = create_empty_dataframe(KEYWORDS)

    for target_number in sorted_rekv_nrs:
        find_and_move_to_rekv_column()

        found_number = False
        while not found_number:
            cursor_pos = win32api.GetCursorPos()
            screenshot, screenshot_coords = capture_screenshot(cursor_pos, width=150, height=1000)

            rekv_number_data = handle_rekv_number(target_number, cursor_pos, screenshot, screenshot_coords)

            if rekv_number_data is not None:
                df = df.append(rekv_number_data, ignore_index=True)
                found_number = True

        click_pato_bank()

    remove_existing_file('file2.xlsx')
    write_dataframe_to_excel(df, 'file2.xlsx')

    file1 = pd.read_excel('file1.xlsx')
    file2 = pd.read_excel('file2.xlsx')

    df = pd.concat([file1, file2], ignore_index=True)
    df = clean_and_organize_dataframe(df, sorted_rekv_nrs)

    write_dataframe_to_excel(df, os.path.join(path_to_save, 'pato_bank.xlsx'))