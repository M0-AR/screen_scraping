import time
import cv2
import keyboard
import numpy as np
import pyautogui
import pyperclip


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
            raise Exception("Could not locate the image on the screen")
        else:
            x, y = pos
            pyautogui.moveTo(x + move_x, y + move_y, duration=1)
            pyautogui.click()
    else:
        # Use text recognition to find the position on the screen
        # Take a screenshot of the entire screen
        screenshot = pyautogui.screenshot()
        # Save the screenshot as a file to read
        screenshot.save("number.png")
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


def save_data_to_excel(path, name, data):
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

    # Split the all_data string into separate lines
    data_lines = data.split('\n')

    # Paste each line in a separate row in the first column
    for i, line in enumerate(data_lines):
        sheet.Cells(i + 1, 1).Value = line

    # Save the workbook with the given file name in the specified path
    file_path = os.path.join(path, name + '.xlsx')
    print(file_path)
    workbook.SaveAs(file_path)

    # Close the workbook
    workbook.Close()

    # Quit Excel
    excel.Quit()


def scroll_down_most(end_scroll, scroll_down_by_x_pixels):
    # Scroll down until red-line found
    while True:
        # Check if the red line image is on the screen
        if pyautogui.locateOnScreen(end_scroll, confidence=0.9):
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


def select_and_copy_data_from_table(up_left_corner_position, up_right_corner_position, end_scroll_position,
                                    scroll_down_by_x_pixels=250, move_x=0, move_y=0):
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
    click_by_mouse_on(up_left_corner_position)

    # Wait for 1 second
    time.sleep(1)

    # Click and hold the left mouse button
    pyautogui.mouseDown(button='left')

    # Get the position of the up-right corner image, and move the mouse to that position
    x, y = pyautogui.locateCenterOnScreen(up_right_corner_position, confidence=0.9)
    pyautogui.moveTo(x + move_x, y + move_y, duration=1)

    # Scroll down to the specified end position
    scroll_down_most(end_scroll_position, scroll_down_by_x_pixels)

    # Get the current position of the mouse cursor
    x, _ = pyautogui.position()

    # Get the position of the end_scroll image to handle the bottom-most edge case
    _, y = pyautogui.locateCenterOnScreen(end_scroll_position, confidence=0.9)
    pyautogui.moveTo(x, y + 100, duration=1)

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
