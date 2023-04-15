# import os
# import time
#
# from master.global_functions import click_by_mouse_on, select_and_copy_data_from_table, save_data_to_excel, \
#     write_dataframe_to_excel
#
# # Wait for 1 second
# time.sleep(1)
#
# click_by_mouse_on('images/vitale/01-search.jpg', confidence=0.95)
#
# # Wait for 1 second
# time.sleep(1)
#
# import pyperclip
#
# # Copy a variable value to the clipboard
# pyperclip.copy('vitale')
#
# import keyboard
#
# # Press and hold the Ctrl key
# keyboard.press('ctrl')
#
# # Press and release the C key
# keyboard.press('v')
# keyboard.release('v')
#
# # Release the Ctrl key
# keyboard.release('ctrl')
# # Wait for 1 second
# time.sleep(1)
#
# click_by_mouse_on('images/vitale/02-choose-vitale.jpg')
#
# # Wait for 1 second
# time.sleep(2)
#
# data = select_and_copy_data_from_table('images/vitale/03-top-left-corner.jpg',
#                                        'images/vitale/04-top-right-corner.jpg',
#                                        'images/vitale/05-bottom-right-corner.jpg',
#                                        confidence_of_end_scroll=0.9)
# # save_data_to_excel('.', 'vitale',data)
# print(data)
# import pandas as pd
#
# rows = data.split('\n')
#
# import re
#
# # split each element by tab character
# data = [x.split("\t") for x in rows]
#
# last_time_index = -1
#
# for index, element in enumerate(data):
#     if isinstance(element[0], str):
#         time = re.search(r"\d{2}:\d{2}", element[0])
#         if len(element) == 1 and time:
#             last_time_index = index
#             break
#
# data_rows = data[last_time_index+1:]
# print(data_rows)
#
# data = data[:last_time_index+1]
# print(data)
#
#
# combined_data = []
# for index, element in enumerate(data):
#     date_match = re.search(r"\d{2}-\d{2}-\d{4}", element[-1])
#     if date_match:
#         time_match = re.search(r"\d{2}:\d{2}", data[index+1][0]) if len(element) > 0 else None
#         if time_match:
#             datetime_str = f"{date_match.group()} {time_match.group()}"
#             combined_data.append([datetime_str])
#
# print(combined_data)
#
# column_names = ["værdier", "Seneste værdi"] + [cd[0] for cd in combined_data]
# df = pd.DataFrame(columns=column_names)
#
# print(df)
#
# # Iterate through the rows of data and add them to the DataFrame
# for row in data_rows:
#     if len(row) > 1:
#         data_dict = dict(zip(column_names, [row[0], row[1]] + row[2:]))
#         df = df.append(data_dict, ignore_index=True)
#     else:
#         # Handle empty rows
#         df = df.append({}, ignore_index=True)
#
# print(df)
# write_dataframe_to_excel(df, os.path.join('C:\src\screen_scraping', 'vitale.xlsx'))

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


def search_and_export_data_to_excel(search_term, path_to_save, file_name):
    """
    Search for a term and export the resulting table data to an Excel file.

    :param search_term: The term to search for.
    :param path_to_save: The path to the directory where the resulting Excel file should be saved.
    :param file_name: The name of the Excel file to save the data.
    """
    click_search_button()
    enter_search_term(search_term)
    choose_search_result()
    data = collect_table_data()
    processed_data = preprocess_raw_data(data)
    combined_data = combine_date_and_time(processed_data)
    df = create_dataframe_with_processed_data(combined_data)
    save_dataframe_to_excel(df, path_to_save, file_name)


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


def main_extract_vitale_data(path_to_save):
    # Example usage:
    search_and_export_data_to_excel('vitale', path_to_save, 'vitale.xlsx')
