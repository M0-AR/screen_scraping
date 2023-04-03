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
    click_by_mouse_on('images/blood_test/02-laboratoriesvar.jpg')

    # Delay the click for 5 second
    time.sleep(5)

    all_data = extract_blood_test_data_from_columns()

    save_data_to_excel(patient_path, 'blood_test', all_data)

    # Delay the click for 5 second
    time.sleep(5)


if __name__ == "__main__":
    # Create a directory named 'HospitalData' in the current working directory
    root_directory = 'HospitalData'
    create_directory(root_directory)

    documents_path = os.path.expanduser("~\Documents")
    root_path = os.path.join(documents_path, root_directory)

    cpr_nrs = read_cpr_nr_from_excel()
    for cpr_nr in cpr_nrs:
        create_directory(cpr_nr, root_path)
        patient_directory = root_directory + '/' + cpr_nr
        patient_path = os.path.join(documents_path, patient_directory)

        click_by_mouse_on('images/general/01-patientopslag.jpg')
        # Delay the click for 2 second
        time.sleep(2)

        input_cpr_nr(cpr_nr, 'images/general/02-navn-or-cprNr.jpg', 250, 0)
        # Delay the click for 2 second
        time.sleep(2)

        click_by_mouse_on('images/general/03-find-patient.jpg')
        # Delay the click for 1 second
        time.sleep(3)

        click_by_mouse_on('images/general/04-vaelg.jpg')
        # Delay the click for 2 second
        time.sleep(2)

        click_by_mouse_on('images/general/05-if-aabn-journal.jpg')
        # Delay the click for 2 second
        time.sleep(4)

        #################################################
        # Extract blood test data
        extract_blood_test_data(patient_path)
        # End of extracting blood test data
        #################################################

        time.sleep(2)

        #################################################
        # Extract miba data
        from miba.miba import extract_miba_data

        extract_miba_data(patient_path)
        # End of extracting miba data
        #################################################

        time.sleep(2)

        #################################################
        # Extract pato_bank data
        from pato_bank.extract_rekav_data_by_number import main_extract_pato_bank_data

        main_extract_pato_bank_data(patient_path)
        # End of extracting pato_bank data
        #################################################

        time.sleep(2)

        click_by_mouse_on('images/general/08-close-by-x.jpg', 40, 0)
        # Delay the click for 1 second
        time.sleep(1)
