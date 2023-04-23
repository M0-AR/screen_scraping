import os
import time
import pandas as pd

from global_functions import click_by_mouse_on, save_data_to_excel, create_directory


def read_cpr_nr_from_excel(file_path: str) -> list:
    """
    Reads CPR numbers from an Excel file and returns them as a list of strings.

    :param file_path: A string representing the path to the Excel file containing the CPR numbers.
    :return: A list of strings representing the CPR numbers.
    :raises TypeError: If the file_path parameter is not a string.
    :raises ValueError: If the file_path parameter is an invalid or empty path, or if the Excel file cannot be read.
    """
    # Validate input parameter
    if not isinstance(file_path, str):
        raise TypeError("The file_path parameter must be a string.")

    if not file_path or not os.path.exists(file_path):
        raise ValueError(f"The file path '{file_path}' is invalid or empty.")

    try:
        # Read the Excel file
        df = pd.read_excel(file_path)

        # Extract the second column without the first column
        cprs = df.iloc[:, 0].values.tolist()

        return cprs
    except Exception as e:
        raise ValueError(f"Failed to read CPR numbers from Excel file '{file_path}': {e}")


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


def extract_blood_test_data(path_to_save_data: str) -> None:
    """
    Extracts blood test data from a patient's record and saves it to an Excel file.

    :param path_to_save_data: A string representing the path to the patient directory where the blood test data will be saved.
    :raises TypeError: If the patient_path parameter is not a string.
    :raises ValueError: If the patient_path parameter is an invalid or empty path, or if there is an error extracting or saving the blood test data.
    """
    from blood_test.extract_data_from_columns import extract_blood_test_data_from_columns

    # Validate input parameter
    if not isinstance(path_to_save_data, str):
        raise TypeError("The patient_path parameter must be a string.")

    if not path_to_save_data or not os.path.exists(path_to_save_data):
        raise ValueError(f"The patient path '{path_to_save_data}' is invalid or empty.")

    try:
        # Extract blood test data
        click_by_mouse_on('images/blood_test/02-laboratoriesvar.jpg')

        # Delay the click for 5 seconds
        time.sleep(5)

        all_data = extract_blood_test_data_from_columns()

        save_data_to_excel(path_to_save_data, 'blood_test', all_data)

        # Delay the click for 5 seconds
        time.sleep(5)
    except Exception as e:
        print(f"Failed to extract blood test data: {e} {path_to_save_data}")


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

        #################################################
        # Extract diagnose_list data
        from diagnose_liste.diagnose_list import main_extract_diagnose_list_data

        main_extract_diagnose_list_data(patient_path)
        # End of extracting diagnose_list data
        #################################################

        time.sleep(2)

        #################################################
        # Extract vitale data
        from vitale.vitale import main_extract_vitale_data

        main_extract_vitale_data(patient_path)
        # End of extracting vitale data
        #################################################

        time.sleep(2)
        click_by_mouse_on('images/general/06-vis-journal.jpg')
        time.sleep(2)

        #################################################
        # Extract notater data
        from notater.notater import main_extract_notater_data

        main_extract_notater_data(patient_path)
        # End of extracting notater data
        #################################################

        time.sleep(4)

        #################################################
        # Extract medicin data
        from medicin.medicin import main_extract_medicin_data

        main_extract_medicin_data(patient_path)
        # End of extracting medicin data
        #################################################

        time.sleep(2)

        click_by_mouse_on('images/general/08-close-by-x.jpg', 40, 0)
        # Delay the click for 1 second
        time.sleep(2)
