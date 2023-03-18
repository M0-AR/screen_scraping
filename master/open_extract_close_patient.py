import os
import time
from global_functions import click_by_mouse_on, save_data_to_excel, create_directory


def read_cpr_nr_from_excel():
    import pandas as pd

    # Read the excel file
    df = pd.read_excel('cprs.xlsx')

    # Extract the second column without the first column
    cprs = df.iloc[:, 0].values

    return cprs


def input_cpr_nr(cpr_nr, image, x, y):
    click_by_mouse_on(image, x, y)

    import pyperclip
    # Copy a variable value to the clipboard
    pyperclip.copy(cpr_nr)

    import keyboard
    # Press and hold the Ctrl key
    keyboard.press('ctrl')

    # Press and release the C key
    keyboard.press('v')
    keyboard.release('v')

    # Release the Ctrl key
    keyboard.release('ctrl')


def extract_blood_test_data(patient_path):
    """
    Extracts blood test data from a patient's record and saves it to an Excel file.
    """

    from blood_test.extract_data_from_columns import extract_blood_test_data_from_columns

    # Extract blood test data
    click_by_mouse_on('06-laboratoriesvar.jpg')

    # Delay the click for 5 second
    time.sleep(5)

    all_data = extract_blood_test_data_from_columns()

    save_data_to_excel(patient_path, 'blood_test', all_data)

    # Delay the click for 2 second
    time.sleep(2)


def extract_miba_data(patient_path):
    from miba.miba import extract_miba_all_data

    # Click on the 'MIBA' button
    click_by_mouse_on('images/miba/01-miba.jpg')

    # Wait for 1 second
    time.sleep(1)

    # Click on the 'All results' button
    click_by_mouse_on('images/miba/02-all-result.jpg')

    # Wait for 1 second
    time.sleep(1)

    # Click on the 'Select all' button
    click_by_mouse_on('images/miba/03-select-all.jpg', -50, 0)

    # Wait for 1 second
    time.sleep(1)

    # Click on the 'Display' button
    click_by_mouse_on('images/miba/04-display-all-result.jpg')

    # Wait for 1 second
    time.sleep(1)

    all_data = extract_miba_all_data()

    save_data_to_excel(patient_path, 'miba', all_data)

    # Delay the click for 2 second
    time.sleep(2)


if __name__ == "__main__":
    # Create a directory named 'hospital_data' in the current working directory
    root_path = "../hospital_data"
    create_directory(root_path)

    cpr_nrs = read_cpr_nr_from_excel()
    for cpr_nr in cpr_nrs:

        # Create a directory named '{patient's cpr-nr}' under 'hospital_data'
        patient_path = root_path + '/' + cpr_nr
        create_directory(patient_path)

        click_by_mouse_on('images/general/01-patientopslag.jpg')
        # Delay the click for 1 second
        time.sleep(2)

        input_cpr_nr(cpr_nr, 'images/general/02-navn-or-cprNr.jpg', 250, 0)
        # Delay the click for 1 second
        time.sleep(2)

        click_by_mouse_on('images/general/03-find-patient.jpg')
        # Delay the click for 1 second
        time.sleep(3)

        click_by_mouse_on('images/general/04-vaelg.jpg')
        # Delay the click for 1 second
        time.sleep(2)

        click_by_mouse_on('images/general/05-if-aabn-journal.jpg')
        # Delay the click for 1 second
        time.sleep(4)

        #################################################
        # Extract blood test data
        """
        extract_blood_test_data(patient_path)
        """
        # End of extracting blood test data
        #################################################

        #################################################
        # Extract miba data
        extract_miba_data(patient_path)
        # End of extracting miba data
        #################################################

        """
        #################################################
        click_by_mouse_on('images/pato_bank/01-patobank.jpg')
        # Delay the click for 1 second
        time.sleep(2)

        #################################################
        """

        click_by_mouse_on('images/general/08-close-by-x.jpg', 75, 0)
        # Delay the click for 1 second
        time.sleep(1)
