import cv2
import numpy as np
import pyautogui


def click_by_mouse_on(position, move_x=0, move_y=0):
    # take a screenshot of the entire screen
    screenshot = pyautogui.screenshot()

    # save the screenshot as a file for to read
    screenshot.save("screenshot.png")

    # Load the screenshot
    screenshotNP = np.array(pyautogui.screenshot())

    # Convert to grayscale
    gray = cv2.cvtColor(screenshotNP, cv2.COLOR_BGR2GRAY)

    # Load the template image
    template = cv2.imread(position, cv2.IMREAD_GRAYSCALE)

    # Find the text on the screenshot using template matching
    result = cv2.matchTemplate(gray, template, cv2.TM_CCOEFF_NORMED)
    y, x = np.unravel_index(result.argmax(), result.shape)

    # Move the mouse to the desire area
    x = x + move_x
    y = y + move_y

    # Click on the center of the text
    width, height = template.shape[::-1]
    pyautogui.moveTo(x + width / 2, y + height / 2)

    # Simulate a right-click at the current mouse position
    pyautogui.click()


def save_data_to_excel(cpr_nr, data):
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

    # Save the workbook with the name 'cprNumber'
    workbook.SaveAs(cpr_nr)

    # Close the workbook
    workbook.Close()

    # Quit Excel
    excel.Quit()
