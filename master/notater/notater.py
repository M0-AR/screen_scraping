import os
import time

from global_functions import click_by_mouse_on, input_text_in_field, move_file


def main_extract_notater_data(path_to_save_data: str) -> None:
    """
    Extracts notater data from a graphical user interface and saves it to a specified patient path.

    :param path_to_save_data: A string representing the path to the patient directory where the notater data will be saved.
    :raises TypeError: If the patient_path parameter is not a string.
    :raises ValueError: If the patient_path parameter is an invalid or empty path.
    :raises Exception: If any of the image files cannot be found or clicked, or if there is an error saving the notater data.
    """
    # Validate input parameter
    if not isinstance(path_to_save_data, str):
        raise TypeError("The patient_path parameter must be a string.")

    if not path_to_save_data or not os.path.exists(path_to_save_data):
        raise ValueError(f"The patient path '{path_to_save_data}' is invalid or empty.")

    try:
        # Get path to Documents directory
        documents_path = os.path.join(os.path.expanduser('~'), 'Downloads')

        # Create path to source file
        source_path = os.path.join(documents_path, 'notater.pdf')

        # Check if file exists in source_path
        if os.path.exists(source_path):
            # If the file exists, remove it
            os.remove(source_path)
            print(f"Notater.pdf found for the previous patient: {path_to_save_data}")

        # Click on the first image and wait for 2 seconds
        click_by_mouse_on('master/images/notater/01-notater.jpg')
        time.sleep(2)

        # Click on the third image, move 60 pixels on the x-axis, and wait for 1 second
        click_by_mouse_on('master/images/notater/02-flere.jpg', move_x=27, confidence=0.9)
        time.sleep(2)

        # Click on the second image and wait for 2 seconds
        click_by_mouse_on('master/images/notater/03-select-all.jpg')
        time.sleep(2)

        # Click on the second image and wait for 2 seconds
        click_by_mouse_on('master/images/notater/04-print.jpg')
        time.sleep(10)

        # Click on the second image and wait for 2 seconds
        click_by_mouse_on('master/images/notater/05-print-out.jpg')
        time.sleep(20)

        input_text_in_field('notater', 'master/images/notater/07-input-text.jpg', 30, 0)
        time.sleep(2)

        click_by_mouse_on('master/images/notater/07-01-chosse-folder.jpg')
        time.sleep(2)

        # Click on the second image and wait for 2 seconds
        click_by_mouse_on('master/images/notater/08-save.jpg')
        time.sleep(4)

        # Click on the second image and wait for 2 seconds
        click_by_mouse_on('master/images/notater/06-no.jpg')
        time.sleep(3)

        # Move the file to the destination directory
        move_file(source_path, path_to_save_data)
    except Exception as e:
        print(f"Failed to extract notater data: {e} {path_to_save_data}")