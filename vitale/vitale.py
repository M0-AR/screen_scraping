import os
import time
import re
import pyperclip
import keyboard
import pandas as pd

from master.global_functions import (
    click_by_mouse_on,
    select_and_copy_data_from_table,
    write_dataframe_to_excel,
)


def click_search_button():
    """Click the search button."""
    time.sleep(1)
    click_by_mouse_on('images/vitale/01-search.jpg', confidence=0.95)
    time.sleep(1)


def enter_search_term(search_term):
    """
    Enter the search term into the search field and submit.

    :param search_term: The term to search for.
    """
    pyperclip.copy(search_term)
    keyboard.press('ctrl')
    keyboard.press('v')
    keyboard.release('v')
    keyboard.release('ctrl')
    time.sleep(1)


def choose_search_result():
    """Choose the search result."""
    click_by_mouse_on('images/vitale/02-choose-vitale.jpg')
    time.sleep(2)


def collect_table_data():
    """
    Collect table data from the page.

    :return: The raw table data as a string.
    """
    return select_and_copy_data_from_table(
        'images/vitale/03-top-left-corner.jpg',
        'images/vitale/04-top-right-corner.jpg',
        'images/vitale/05-bottom-right-corner.jpg',
        confidence_of_end_scroll=0.9
    )


def preprocess_raw_data(data):
    """
    Preprocess the raw data.

    :param data: The raw table data as a string.
    :return: A tuple containing the preprocessed data and data rows.
    """
    rows = data.split('\n')
    data = [x.split("\t") for x in rows]
    last_time_index = -1

    for index, element in enumerate(data):
        if isinstance(element[0], str):
            time = re.search(r"\d{2}:\d{2}", element[0])
            if len(element) == 1 and time:
                last_time_index = index
                break

    data_rows = data[last_time_index + 1:]
    data = data[:last_time_index + 1]
    return data, data_rows


def combine_date_and_time(processed_data):
    """
    Combine date and time for each row.

    :param processed_data: A tuple containing the preprocessed data and data rows.
    :return: A tuple containing the combined data and data rows.
    """
    data, data_rows = processed_data
    combined_data = []

    for index, element in enumerate(data):
        date_match = re.search(r"\d{2}-\d{2}-\d{4}", element[-1])
        if date_match:
            time_match = re.search(r"\d{2}:\d{2}", data[index + 1][0]) if len(element) > 0 else None
            if time_match:
                datetime_str = f"{date_match.group()} {time_match.group()}"
                combined_data.append([datetime_str])

    return combined_data, data_rows


def create_dataframe_with_processed_data(combined_data):
    """
    Create a DataFrame with the processed data.

    :param combined_data: A tuple containing the combined data and data rows.
    :return: A pandas DataFrame containing the processed data.
    """
    combined_data, data_rows = combined_data
    column_names = ["værdier", "Seneste værdi"] + [cd[0] for cd in combined_data]
    df = pd.DataFrame(columns=column_names)

    for row in data_rows:
        if len(row) > 1:
            data_dict = dict(zip(column_names, [row[0], row[1]] + row[2:]))
            df = df.append(data_dict, ignore_index=True)
        else:
            # Handle empty rows
            df = df.append({}, ignore_index=True)

    return df


def save_dataframe_to_excel(df, path_to_save, file_name):
    """
    Save the DataFrame to an Excel file.

    :param df: A pandas DataFrame to save.
    :param path_to_save: The path to the directory where the resulting Excel file should be saved.
    :param file_name: The name of the Excel file to save the data.
    """
    write_dataframe_to_excel(df, os.path.join(path_to_save, file_name))


def search_and_export_data_to_excel(search_term: str, path_to_save: str, file_name: str) -> None:
    """
    Searches for table data related to a given search term in the Vitale system and exports it to an Excel file.

    :param search_term: A string representing the search term to look for in the Vitale system.
    :param path_to_save: A string representing the path to the directory where the resulting Excel file should be saved.
    :param file_name: A string representing the name of the Excel file to save the data.
    :raises TypeError: If any of the input parameters are not strings.
    :raises ValueError: If the path_to_save parameter is an invalid or empty path, or if there is an error searching for or exporting the Vitale data.
    """
    # Validate input parameters
    if not all(isinstance(param, str) for param in (search_term, path_to_save, file_name)):
        raise TypeError("All input parameters must be strings.")

    if not path_to_save or not os.path.exists(path_to_save):
        raise ValueError(f"The path to save data '{path_to_save}' is invalid or empty.")

    # Perform a search for the given search term in the Vitale system
    click_search_button()
    enter_search_term(search_term)
    choose_search_result()
    data = collect_table_data()

    # Preprocess the raw data
    processed_data = preprocess_raw_data(data)
    combined_data = combine_date_and_time(processed_data)

    # Create a pandas DataFrame with the processed data and save to an Excel file
    df = create_dataframe_with_processed_data(combined_data)
    save_dataframe_to_excel(df, path_to_save, file_name)


def main_extract_vitale_data(path_to_save_data: str) -> None:
    """
    Main function to extract data from the Vitale system and save it to an Excel file.

    :param path_to_save_data: A string representing the path to the directory where the resulting Excel file should be saved.
    :raises TypeError: If the input parameter is not a string.
    :raises ValueError: If the path_to_save parameter is an invalid or empty path, or if there is an error searching for or exporting the Vitale data.
    """
    # Validate input parameter
    if not isinstance(path_to_save_data, str):
        raise TypeError("Input parameter 'path_to_save' must be a string.")

    if not path_to_save_data or not os.path.exists(path_to_save_data):
        raise ValueError(f"The path to save data '{path_to_save_data}' is invalid or empty.")

    try:
        # Example usage:
        search_and_export_data_to_excel('vitale', path_to_save_data, 'vitale.xlsx')
    except Exception as e:
        print(f"Failed to extract vitale data: {e} {path_to_save_data}")
