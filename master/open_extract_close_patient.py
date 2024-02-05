import os
import time
import pandas as pd
import math
import re
from global_functions import click_by_mouse_on, click_by_mouse_on_without_threshold, save_data_to_excel, create_directory


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

    #if not file_path or not os.path.exists(file_path):
        #raise ValueError(f"The file path '{file_path}' is invalid or empty.")

    try:
        # Read the Excel file
        df = pd.read_excel(file_path)

        # Extract the second column without the first column
        cprs = df.iloc[:, 0].values.tolist()

        return cprs
    except Exception as e:
        raise ValueError(f"Failed to read CPR numbers from Excel file '{file_path}': {e}")


def input_cpr_nr(cpr_nr, image, x, y):
    # create a regular expression pattern to match "xxxxxx-xxxx" format
    pattern = r'^\d{6}-\d{4}$'
    if not bool(re.match(pattern, cpr_nr)):
        return False

    click_by_mouse_on(image, x, y)

    import pyperclip
    # Copy a variable value to the clipboard
    pyperclip.copy(cpr_nr)

    import keyboard
    # Press and hold the Ctrl key
    keyboard.press('ctrl')
    time.sleep(.2)

    # Press and release the C key
    keyboard.press('v')
    time.sleep(.2)
    keyboard.release('v')

    time.sleep(.2)
    # Release the Ctrl key
    keyboard.release('ctrl')

    return True


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
        click_by_mouse_on('master/images/blood_test/02-laboratoriesvar.jpg')

        # Delay the click for 5 seconds
        time.sleep(5)

        all_data = extract_blood_test_data_from_columns()

        # create a regular expression pattern to match "xxxxxx-xxxx" format
        pattern = r'^\d{6}-\d{4}\n*$'
        match = re.search(pattern, all_data)

        if match:
            print(f"No blood test data exist for: {path_to_save_data}")
        else:
            save_data_to_excel(path_to_save_data, 'blood_test', all_data)
            
        # Delay the click for 5 seconds
        time.sleep(5)
    except Exception as e:
        print(f"Failed to extract blood test data: {e} {path_to_save_data}")

def isnan(value):
    try:
        return math.isnan(float(value))
    except:
        return False

def all_patient_files_exist(patient_path):
    # List of files to be checked
    files_to_check = [
        'blood_test.xlsx',
        'pato_bank.xlsx',
        'notater.pdf',
        'medicin.xlsx',
        'miba.xlsx',
        'diagnose_list.xlsx',
        'vitale.xlsx',
    ]
    
    # Check if each file exists in the patient_path
    for file_name in files_to_check:
        file_path = os.path.join(patient_path, file_name)
        if not os.path.exists(file_path):
            return False  # Return False if any file does not exist
    
    # Return True if all files exist
    return True

if __name__ == "__main__":
    # Create a directory named 'HospitalData' in the current working directory
    root_directory = 'HospitalData'
    create_directory(root_directory)

    documents_path = os.path.expanduser("~\Documents")
    root_path = os.path.join(documents_path, root_directory)

    prev_cpr = ''
    cpr_nrs = read_cpr_nr_from_excel('master/cprs.xlsx')
    for cpr_nr in cpr_nrs: 
        if prev_cpr == cpr_nr or isnan(cpr_nr) or cpr_nr is ' ' or cpr_nr is '':
            continue
        
        prev_cpr = cpr_nr

        create_directory(cpr_nr, root_path)
        patient_directory = root_directory + '/' + cpr_nr
        patient_path = os.path.join(documents_path, patient_directory)

        if not all_patient_files_exist(patient_path):
            time.sleep(2)

            click_by_mouse_on('master/images/general/01-patientopslag.jpg')
            # Delay the click for 2 second
            time.sleep(2)

            is_cpr_inputed = input_cpr_nr(cpr_nr, 'master/images/general/02-navn-or-cprNr.jpg', 250, 0)
            if not is_cpr_inputed:
                import logging
                logging.basicConfig(filename='master/log/cpr.log', level=logging.DEBUG)
                logging.debug(f"\n\n\nFailed to extract cpr's data: {cpr_nr} \n\n")
                continue
            
            # Delay the click for 2 second
            time.sleep(2)

            click_by_mouse_on('master/images/general/03-find-patient.jpg')
            # Delay the click for 1 second
            time.sleep(3)

            click_by_mouse_on('master/images/general/04-vaelg.jpg')
            # Delay the click for 2 second
            time.sleep(15)

            click_by_mouse_on('master/images/general/05-if-aabn-journal.jpg')
            # Delay the click for 2 second
            time.sleep(10)

            #################################################
            # Extract blood test data
            blood_test_file_path = os.path.join(patient_path, 'blood_test.xlsx')
            # Check if the file exists
            if not os.path.isfile(blood_test_file_path):
                # If the file does not exist, proceed with extracting blood test data
                extract_blood_test_data(patient_path)
                time.sleep(2)
            else:
                time.sleep(1)
            # End of extracting blood test data
            #################################################

            

            #################################################
            # Extract pato_bank data
            # Define the path to the pato_bank.xlsx file within the patient_path folder
            pato_bank_file_path = os.path.join(patient_path, 'pato_bank.xlsx')

            # Check if the file exists
            if not os.path.exists(pato_bank_file_path):
                from pato_bank.extract_rekav_data_by_number import main_extract_pato_bank_data
                main_extract_pato_bank_data(patient_path)
                time.sleep(2)
                click_by_mouse_on('master/images/general/06-vis-journal.jpg')
                time.sleep(6)
            else:
                time.sleep(1)
            # End of extracting pato_bank data
            #################################################

            

            #################################################
            # Extract notater data
            # Check if notater.pdf exists in patient_path folder
            notater_file_path = os.path.join(patient_path, "notater.pdf")

            # Check for notater.pdf existence
            if not os.path.exists(notater_file_path):
                from notater.notater import main_extract_notater_data
                main_extract_notater_data(patient_path)
                time.sleep(4)
            else:
                time.sleep(1)
            # End of extracting notater data
            #################################################


            #################################################
            # Extract medicin data               
            # Check if medicin.xlsx exists in patient_path folder
            medicin_file_path = os.path.join(patient_path, "medicin.xlsx")

            # Check for medicin.xlsx existence
            if not os.path.exists(medicin_file_path):
                from medicin.medicin import main_extract_medicin_data
                main_extract_medicin_data(patient_path)
                time.sleep(2)
            else:
                time.sleep(1)
            # End of extracting medicin data
            #################################################


            #################################################
            # Extract miba data             
            # Define the path for miba.xlsx in the patient_path folder
            miba_file_path = os.path.join(patient_path, "miba.xlsx")

            # Check for the existence of miba.xlsx, and if it does not exist, prepare to call the extraction function
            if not os.path.exists(miba_file_path):
                from miba.miba import extract_miba_data
                extract_miba_data(patient_path)
                time.sleep(2)
            else:
                time.sleep(1)
            # End of extracting miba data
            #################################################


            #################################################
            # Extract diagnose_list data
            # Define the path for diagnose_list.xlsx in the patient_path folder
            diagnose_list_file_path = os.path.join(patient_path, "diagnose_list.xlsx")

            # Check for the existence of diagnose_list.xlsx, and if it does not exist, prepare to call the extraction function
            if not os.path.exists(diagnose_list_file_path):
                from diagnose_liste.diagnose_list import main_extract_diagnose_list_data
                main_extract_diagnose_list_data(patient_path)
                time.sleep(4)
            else:
                time.sleep(1)
            # End of extracting diagnose_list data
            #################################################


            #################################################
            # Extract vitale data      
            # Define the path for vitale.xlsx in the patient_path folder
            vitale_file_path = os.path.join(patient_path, "vitale.xlsx")

            # Check for the existence of vitale.xlsx, and if it does not exist, prepare to call the extraction function
            if not os.path.exists(vitale_file_path):
                from vitale.vitale import main_extract_vitale_data

                main_extract_vitale_data(patient_path)
            # End of extracting vitale data
            #################################################

            time.sleep(5)
            click_by_mouse_on_without_threshold('master/images/general/08-close-by-x.jpg', 20, 0)
            # Delay the click for 18 second
            time.sleep(18)
