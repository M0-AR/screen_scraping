import time
import cv2
import keyboard
import numpy as np
import pyautogui
import pyperclip

import pyperclip
import keyboard

import os
import shutil


def move_file(source_path: str, destination_path: str) -> None:
    """
    Move a file from the source path to the destination path.

    :param source_path: A string representing the path to the source file.
    :param destination_path: A string representing the path to the destination directory or file.
    :raises TypeError: If the input parameters are not strings.
    :raises ValueError: If either the source path or destination path is invalid or empty.
    """
    # Validate input parameters
    if not isinstance(source_path, str) or not isinstance(destination_path, str):
        raise TypeError("Both the source_path and destination_path parameters must be strings.")

    if not os.path.exists(source_path):
        raise ValueError(f"The source path '{source_path}' does not exist.")

    if not destination_path or not os.path.exists(os.path.dirname(destination_path)):
        raise ValueError(f"The destination path '{destination_path}' is invalid or empty.")

    # Move the file to the destination path
    try:
        shutil.move(source_path, destination_path)
    except Exception as e:
        raise ValueError(f"Failed to move file from '{source_path}' to '{destination_path}': {e}")

    return None


def input_text_in_field(text: str, image_path: str, x: int, y: int) -> None:
    """
    Inputs a text into a designated field on a GUI.

    :param text: A string of text.
    :param image_path: A string containing the file path of the GUI element to click.
    :param x: An integer representing the x-coordinate of the click location on the GUI element.
    :param y: An integer representing the y-coordinate of the click location on the GUI element.
    :raises TypeError: If the input parameters are of the incorrect type.
    :raises ValueError: If the cpr_number parameter is not a 10-digit string.
    """
    # Validate input parameters
    if not isinstance(text, str):
        raise TypeError("text parameter must be a string.")

    if not isinstance(image_path, str):
        raise TypeError("image_path parameter must be a string.")

    if not isinstance(x, int) or not isinstance(y, int):
        raise TypeError("x and y parameters must be integers.")

    # Click on the GUI element
    click_by_mouse_on(image_path, x, y)

    # Copy the CPR number to the clipboard
    pyperclip.copy(text)

    # Paste the CPR number from the clipboard using keyboard shortcuts
    try:
        keyboard.press_and_release('ctrl+v')
    except keyboard.KeyboardError:
        print("Error: Could not paste CPR number using keyboard shortcuts.")

    return None


def move_mouse_on(
        position: str,
        move_x: int = 0,
        move_y: int = 0,
) -> None:
    """Find an image or text on the screen and move a mouse

    Args:
        position (str): The image or text to be searched on the screen.
        move_x (int, optional): The amount to move the mouse cursor on the X-axis
            after finding the image or text. Defaults to 0.
        move_y (int, optional): The amount to move the mouse cursor on the Y-axis
            after finding the image or text. Defaults to 0.

    Returns:
        None: The function does not return a value.

    Raises:
        Exception: If the image or text cannot be located on the screen.

    """
    # Take a screenshot of the entire screen
    screenshot = pyautogui.screenshot()
    # Load the screenshot as a NumPy array
    screenshot_np = np.array(screenshot)
    # Convert the screenshot to grayscale
    gray = cv2.cvtColor(screenshot_np, cv2.COLOR_BGR2GRAY)
    # Load the template image
    template = cv2.imread(position, cv2.IMREAD_GRAYSCALE)
    # Find the text on the screenshot using template matching
    result = cv2.matchTemplate(gray, template, cv2.TM_CCOEFF_NORMED)
    y, x = np.unravel_index(result.argmax(), result.shape)
    # Move the mouse to the desired position
    x = x + move_x
    y = y + move_y
    pyautogui.moveTo(x, y, duration=1)
    # Click on the center of the text
    width, height = template.shape[::-1]
    pyautogui.moveTo(x + width / 2, y + height / 2)


def click_by_mouse_on(
        position: str,
        move_x: int = 0,
        move_y: int = 0,
        confidence: float = 0.0
) -> None:
    """Find an image or text on the screen and perform a mouse click.

    Args:
        position (str): The image or text to be searched on the screen.
        move_x (int, optional): The amount to move the mouse cursor on the X-axis
            after finding the image or text. Defaults to 0.
        move_y (int, optional): The amount to move the mouse cursor on the Y-axis
            after finding the image or text. Defaults to 0.
        confidence (float, optional): The minimum confidence level for the image
            or text match. If set to 0, the function will use text recognition
            instead of image matching. Defaults to 0.

    Returns:
        None: The function does not return a value.

    Raises:
        Exception: If the image or text cannot be located on the screen.

    """
    if confidence != 0:
        # Use image matching to find the position on the screen
        pos = pyautogui.locateCenterOnScreen(position, confidence=confidence)
        if pos is None:
            print(f"Could not locate {position} image on the screen")
        else:
            x, y = pos
            pyautogui.moveTo(x + move_x, y + move_y, duration=1)
            pyautogui.click()
    else:
        # Use text recognition to find the position on the screen
        # Take a screenshot of the entire screen
        screenshot = pyautogui.screenshot()
        # Load the screenshot as a NumPy array
        screenshot_np = np.array(screenshot)
        # Convert the screenshot to grayscale
        gray = cv2.cvtColor(screenshot_np, cv2.COLOR_BGR2GRAY)
        # Load the template image
        template = cv2.imread(position, cv2.IMREAD_GRAYSCALE)
        # Find the text on the screenshot using template matching
        result = cv2.matchTemplate(gray, template, cv2.TM_CCOEFF_NORMED)
        y, x = np.unravel_index(result.argmax(), result.shape)
        # Move the mouse to the desired position
        x = x + move_x
        y = y + move_y
        pyautogui.moveTo(x, y, duration=1)
        # Click on the center of the text
        width, height = template.shape[::-1]
        pyautogui.click(x + width / 2, y + height / 2)


def write_dataframe_to_excel(df, file_path, index=False):
    """
    Write the DataFrame to an Excel file.

    Args:
        df (pd.DataFrame): DataFrame to be written to the Excel file.
        file_path (str): The path to the file where the DataFrame should be written.
        index (bool, optional): Whether to write row names (index). Defaults to False.
    """
    df.to_excel(file_path, index=index)


def bring_window_to_front(window_title):
    import win32gui
    """
    Brings a window with the specified title to the front.

    Args:
        window_title (str): The title of the window to bring to the front.
    """

    # Find the window with the specified title
    window_handle = win32gui.FindWindow(None, window_title)

    # If the window is found, bring it to the front
    if window_handle:
        win32gui.SetForegroundWindow(window_handle)


def save_data_to_excel(path, name, data, simulate_ctrl_v=False):
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
    # Wait for a moment
    time.sleep(2)

    if simulate_ctrl_v:
        # Select the first cell
        sheet.Cells(1, 1).Select()

        # Wait for a moment before simulating the keyboard action
        time.sleep(1)

        # Get the Excel window title
        excel_window_title = excel.Caption

        # Bring the Excel window to the front
        bring_window_to_front(excel_window_title)

        # Wait for a moment
        time.sleep(1)

        # Copy the data's content to the clipboard
        pyperclip.copy(data)

        # Simulate pressing 'Ctrl+V'
        press_ctrl_v()
    else:
        # Split the all_data string into separate lines
        data_lines = data.split('\n')

        # Paste each line in a separate row in the first column
        for i, line in enumerate(data_lines):
            sheet.Cells(i + 1, 1).Value = line

    # Save the workbook with the given file name in the specified path
    file_path = os.path.join(path, name + '.xlsx')
    workbook.SaveAs(file_path)

    # Close the workbook
    workbook.Close()

    # Quit Excel
    excel.Quit()


import cv2
import numpy as np
import pyautogui


def detect_scrollbar(scrollbar_template):
    # Load the template image
    template = cv2.imread(scrollbar_template, cv2.IMREAD_GRAYSCALE)

    # Check if the scrollbar template image is valid
    if template is None or template.size == 0:
        raise Exception("Error: The scrollbar template image could not be loaded. Please check the file path and "
                        "format.")

    # Take a screenshot
    screenshot = pyautogui.screenshot()

    # Load the screenshot
    screenshotNP = np.array(screenshot)

    # Convert to grayscale
    gray = cv2.cvtColor(screenshotNP, cv2.COLOR_BGR2GRAY)

    # Calculate the match between the images
    result = cv2.matchTemplate(gray, template, cv2.TM_CCOEFF_NORMED)

    # Calculate the maximum value (confidence)
    max_val = result.max()

    # Set a threshold to consider the template as found
    threshold = 0.5

    return max_val >= threshold


def scroll_down_most(end_scroll, scroll_down_by_x_pixels):
    # Scroll down until red-line found
    while True:
        # Check if the red line image is on the screen
        if pyautogui.locateOnScreen(end_scroll, confidence=0.8):
            break  # exit the loop if the image is found
        else:
            # Scroll down by 'scroll_down_by' pixels
            pyautogui.scroll(-scroll_down_by_x_pixels)


def press_ctrl_c():
    # Press and hold the Ctrl key
    keyboard.press('ctrl')

    # Press and release the C key
    keyboard.press('c')
    keyboard.release('c')

    # Release the Ctrl key
    keyboard.release('ctrl')

    # Delay for 1 second
    time.sleep(1)


def press_ctrl_v():
    # Press and hold the Ctrl key
    keyboard.press('ctrl')

    # Press and release the V key
    keyboard.press('v')
    keyboard.release('v')

    # Release the Ctrl key
    keyboard.release('ctrl')

    # Delay for 1 second
    time.sleep(1)


def select_and_copy_data_from_table(up_left_corner_position, up_right_corner_position, end_scroll_position,
                                    confidence_of_end_scroll,
                                    scroll_if_sign_found=None, scroll_down_by_x_pixels=250, move_x=0, move_y=0,
                                    move_y_end_scroll=0):
    """
    Extract data from a table within an application window.

    Parameters:
    up_left_corner_position (str): The filename or path of the image representing the top-left corner of the table.
    up_right_corner_position (str): The filename or path of the image representing the top-right corner of the table.
    end_scroll_position (str): The filename or path of the image representing the end position of the scrolling.
    scroll_down_by_x_pixels (int, optional): The number of pixels to scroll down each time. Default is 250.
    move_x (int, optional): The amount of pixels to move the mouse horizontally after reaching the top-right corner of the table. Default is 0.
    move_y (int, optional): The amount of pixels to move the mouse vertically after reaching the top-right corner of the table. Default is 0.

    Returns:
    table_data (str): The selected data from the table, as a string.
    """

    # Move the mouse to the up-left corner image and click it
    click_by_mouse_on(up_left_corner_position, move_x, move_y, confidence=0.9)

    # Wait for 1 second
    time.sleep(1)

    # Click and hold the left mouse button
    pyautogui.mouseDown(button='left')

    print(up_right_corner_position)
    # Get the position of the up-right corner image, and move the mouse to that position
    result = pyautogui.locateCenterOnScreen(up_right_corner_position, confidence=0.9)
    x, y = 0, 0
    if result is not None:
        x, y = result
        pyautogui.moveTo(x + move_x, y + move_y, duration=1)
    else:
        # Handle no scroll-bar case
        x, y = pyautogui.locateCenterOnScreen('images/pato_bank/07-temp-top-right-selection-rekv-nr.jpg',
                                              confidence=0.8)
        pyautogui.moveTo(x + move_x, y + move_y, duration=1)

    if scroll_if_sign_found:
        if detect_scrollbar(scroll_if_sign_found):
            # Scroll down to the specified end position
            scroll_down_most(end_scroll_position, scroll_down_by_x_pixels)

    # Get the current position of the mouse cursor
    x, _ = pyautogui.position()

    # Get the position of the end_scroll image to handle the bottom-most edge case
    # result = pyautogui.locateCenterOnScreen(end_scroll_position, confidence=0.8)
    result = pyautogui.locateCenterOnScreen(end_scroll_position, confidence=confidence_of_end_scroll)
    if result is not None:
        _, y = result
        # 50 to not go beyond the red line
        pyautogui.moveTo(x, y - move_y_end_scroll, duration=1)
    else:
        # Handle the case where there is only one rekv.nr in the PatoBank page
        pyautogui.moveTo(x, y + 900, duration=1)

    # Release the left mouse button
    pyautogui.mouseUp(button='left')

    # Copy the selected data to the clipboard
    press_ctrl_c()

    # Get the text from the clipboard and return it
    table_data = pyperclip.paste()
    return table_data


def create_directory(directory_name, directory_path='~/Documents'):
    import os

    documents_path = os.path.expanduser(directory_path)
    directory_path = os.path.join(documents_path, directory_name)

    if not os.path.exists(directory_path):
        os.mkdir(directory_path)
        print(f"Directory {directory_path} created.")
    else:
        print(f"Directory {directory_path} already exists.")
