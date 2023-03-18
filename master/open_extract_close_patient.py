import time
from global_functions import click_by_mouse_on


def read_cpr_nr_from_excel():
    import pandas as pd

    # Read the excel file
    df = pd.read_excel('cprs.xlsx')

    # Extract the second column without the first column
    cprs = df.iloc[:, 0].values

    return cprs


def input_cpr_nr(cpr_nr):
    click_by_mouse_on('02-navn-or-cprNr.jpg', 250, 0)

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


if __name__ == "__main__":
    cpr_nrs = read_cpr_nr_from_excel()
    for cpr_nr in cpr_nrs:
        click_by_mouse_on('01-patientopslag.jpg')
        # Delay the click for 1 second
        time.sleep(2)

        input_cpr_nr(cpr_nr)
        # Delay the click for 1 second
        time.sleep(2)

        click_by_mouse_on('03-find-patient.jpg')
        # Delay the click for 1 second
        time.sleep(3)

        click_by_mouse_on('04-vaelg.jpg')
        # Delay the click for 1 second
        time.sleep(2)

        click_by_mouse_on('05-if-aabn-journal.jpg')
        # Delay the click for 1 second
        time.sleep(4)

        #################################################
        # Extract blood test data
        #################################################
        click_by_mouse_on('06-laboratoriesvar.jpg')
        # Delay the click for 1 second
        time.sleep(5)

        from blood_test.extract_data_from_columns \
            import extract_blood_test_data_from_columns
        extract_blood_test_data_from_columns(cpr_nr)
        # Delay the click for 1 second
        time.sleep(2)
        # End of extracting blood test data
        #################################################

        #################################################
        click_by_mouse_on('images/pato_bank/01-patobank.jpg')
        # Delay the click for 1 second
        time.sleep(2)

        #################################################

        click_by_mouse_on('images/pato_bank/01-patobank.jpg')
        click_by_mouse_on('08-close-by-x.jpg', 75, 0)
        # Delay the click for 1 second
        time.sleep(1)
