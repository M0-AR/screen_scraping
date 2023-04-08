import time
import pandas as pd
import pyautogui
import pyperclip
from master.global_functions import click_by_mouse_on, press_ctrl_c, save_data_to_excel


def struct_miba_data(patient_path):
    # Load the excel file into a Pandas dataframe
    df = pd.read_excel(patient_path + '/miba.xlsx', header=None)

    # Define the column names
    columns = ['Prøvens art', 'Taget d.', 'Kvantitet', 'Analyser', 'Resistens', 'Mikroskopi']

    # Create a new dataframe to store the extracted data
    new_df = pd.DataFrame(columns=columns)

    # Loop over the rows of the original dataframe, extracting data from rows with the correct format
    for i, row in df.iterrows():
        # Check if the first cell of the row matches the pattern for the column names
        if 'Prøvens nr.' in str(row.iloc[0]):
            # Extract the data from the row and add it to the new dataframe
            data = df.iloc[i + 1, 0]
            # Split the string using tab as a separator and assign each element to its corresponding column
            new_df.at[len(new_df), 'Prøvens art'] = data.split('\t')[1]
            new_df.at[len(new_df) - 1, 'Taget d.'] = data.split('\t')[2]
        elif 'Kvantitet' in str(row.iloc[0]):
            # Set the current column to 'Kvantitet'
            current_column = 'Kvantitet'

            data = ''
            # Loop over the subsequent rows until a new keyword is encountered or an empty row is reached
            for j in range(i + 1, len(df)):
                if pd.isna(df.iloc[j, 0]) or 'Analyser' in str(df.iloc[j, 0]) \
                        or 'Resistens' in str(df.iloc[j, 0]) or 'Mikroskopi' in str(df.iloc[j, 0]) or \
                        str(df.iloc[j, 0]).startswith(' _x000D_'):
                    break
                data += str(df.iloc[j, 0]).replace('_x000D_', '') + '\n'

            # Add the extracted data to the new dataframe
            new_df[current_column][len(new_df) - 1] = data
        elif 'Analyser' in str(row.iloc[0]):
            # Set the current column to 'Analyser'
            current_column = 'Analyser'

            data = ''
            # Loop over the subsequent rows until a new keyword is encountered or an empty row is reached
            for j in range(i + 1, len(df)):
                if pd.isna(df.iloc[j, 0]) or 'Kvantitet' in str(df.iloc[j, 0]) \
                        or 'Resistens' in str(df.iloc[j, 0]) or 'Mikroskopi' in str(df.iloc[j, 0]) or \
                        str(df.iloc[j, 0]).startswith(' _x000D_'):
                    break
                data += str(df.iloc[j, 0]).replace('_x000D_', '') + '\n'

            # Add the extracted data to the new dataframe
            new_df[current_column][len(new_df) - 1] = data
        elif 'Resistens' in str(row.iloc[0]):
            # Set the current column to 'Analyser'
            current_column = 'Resistens'

            data = ''
            # Loop over the subsequent rows until a new keyword is encountered or an empty row is reached
            for j in range(i + 1, len(df)):
                if pd.isna(df.iloc[j, 0]) or 'Kvantitet' in str(df.iloc[j, 0]) \
                        or 'Analyser' in str(df.iloc[j, 0]) or 'Mikroskopi' in str(df.iloc[j, 0]) or \
                        str(df.iloc[j, 0]).startswith(' _x000D_'):
                    break
                data += str(df.iloc[j, 0]).replace('_x000D_', '') + '\n'

            # Add the extracted data to the new dataframe
            new_df[current_column][len(new_df) - 1] = data

        elif 'Mikroskopi' in str(row.iloc[0]):
            # Set the current column to 'Analyser'
            current_column = 'Mikroskopi'

            data = ''
            # Loop over the subsequent rows until a new keyword is encountered or an empty row is reached
            for j in range(i + 1, len(df)):
                if pd.isna(df.iloc[j, 0]) or 'Kvantitet' in str(df.iloc[j, 0]) \
                        or 'Analyser' in str(df.iloc[j, 0]) or 'Resistens' in str(df.iloc[j, 0]) or \
                        str(df.iloc[j, 0]).startswith(' _x000D_'):
                    break
                data += str(df.iloc[j, 0]).replace('_x000D_', '') + '\n'

            # Add the extracted data to the new dataframe
            new_df[current_column][len(new_df) - 1] = data

    # Save the new dataframe to a new excel file
    new_df.to_excel(patient_path + '/miba.xlsx', index=False)


def extract_miba_data(patient_path):
    # Wait for 1 second
    time.sleep(1)

    # Click on the 'MIBA' button
    click_by_mouse_on('images/miba/01-miba.jpg')

    # Wait for 1 second
    time.sleep(3)

    # Click on the 'All results' button
    click_by_mouse_on('images/miba/02-all-result.jpg', confidence=0.7)

    # Wait for 1 second
    time.sleep(1)

    # Click on
    click_by_mouse_on('images/miba/03-select-all.jpg', -50, 0)

    # Wait for 1 second
    time.sleep(1)

    # Click on
    click_by_mouse_on('images/miba/04-display-all-result.jpg')

    # Wait for 1 second
    time.sleep(1)

    # Click on
    click_by_mouse_on('images/miba/05-up-left-corner.jpg')

    # Wait for 1 second
    time.sleep(1)

    # Simulate a right-click at the current mouse position
    pyautogui.rightClick()
    # Delay the click for 2 second
    time.sleep(2)

    # Click on
    click_by_mouse_on('images/miba/06-select-all.jpg')
    # Wait for 1 second
    time.sleep(1)

    # Copy the selected data to the clipboard
    press_ctrl_c()

    # Get the text from the clipboard and return it
    all_data = pyperclip.paste()

    save_data_to_excel(patient_path, 'miba', all_data)
    # Delay the click for 3 second
    time.sleep(3)

    struct_miba_data(patient_path)

    # Click on the 'x' button to close miba
    click_by_mouse_on('images/miba/08-close-miba.jpg')

    # Delay the click for 3 second
    time.sleep(3)
