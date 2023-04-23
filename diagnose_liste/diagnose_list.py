import os
import time
import pandas as pd
from master.global_functions import move_mouse_on, select_and_copy_data_from_table, save_data_to_excel, \
    write_dataframe_to_excel, click_by_mouse_on


def extract_table_data():
    """
    This function extracts the table data from the screen.

    :return: The table data.
    :rtype: str
    """
    return select_and_copy_data_from_table('images/diagnose_list/02-top-left-corner.jpg',
                                           'images/diagnose_list/03-top-right-corner.jpg',
                                           'images/diagnose_list/04-bottom-right-corner.jpg',
                                           confidence_of_end_scroll=0.95,
                                           move_x=25)


def split_table_data_into_rows(table_data):
    """
    This function splits the table data into rows.

    :param table_data: The table data.
    :type table_data: str
    :return: The rows.
    :rtype: list
    """
    return [row.strip() for row in table_data.split("\n")]


def remove_noteret_rows(rows):
    """
    This function removes rows that start with "Noteret".

    :param rows: The rows.
    :type rows: list
    :return: The rows without "Noteret".
    :rtype: list
    """
    return [row for row in rows if not row.startswith("Noteret")]


def split_rows_into_pairs(rows):
    """
    This function splits the rows into pairs of text and date.

    :param rows: The rows.
    :type rows: list
    :return: The pairs of text and date.
    :rtype: list
    """
    return [(row.split("\t")[0], row.split("\t")[1]) for row in rows if "\t" in row]


def create_dataframe_from_pairs(pairs):
    """
    This function creates a DataFrame from pairs of text and date.

    :param pairs: The pairs of text and date.
    :type pairs: list
    :return: The DataFrame.
    :rtype: pandas.DataFrame
    """
    return pd.DataFrame(pairs, columns=["note", "date"])


def main_extract_diagnose_list_data(path_to_save_data: str) -> None:
    """
    Extracts data from the Ambulant Diagnose List and saves it to an Excel file.

    :param path_to_save_data: A string representing the path to the directory where the resulting Excel file should be saved.
    :raises TypeError: If the path_to_save parameter is not a string.
    :raises ValueError: If the path_to_save parameter is an invalid or empty path, or if there is an error extracting or saving the Ambulant Diagnose List data.
    """
    # Validate input parameter
    if not isinstance(path_to_save_data, str):
        raise TypeError("The path_to_save parameter must be a string.")

    if not path_to_save_data or not os.path.exists(path_to_save_data):
        raise ValueError(f"The path to save data '{path_to_save_data}' is invalid or empty.")

    try:
        # Scroll down to load more data, then extract the table data from the Ambulant Diagnose List page
        click_by_mouse_on('images/diagnose_list/00-scroll-down.jpg', confidence=0.8)
        time.sleep(3)

        move_mouse_on('images/diagnose_list/01-diagnoseoverblik.jpg')
        time.sleep(3)

        table_data = extract_table_data()

        # Clean up and format the table data
        rows = split_table_data_into_rows(table_data)
        rows = remove_noteret_rows(rows)
        pairs = split_rows_into_pairs(rows)
        df = create_dataframe_from_pairs(pairs)

        # Save the extracted data to an Excel file
        write_dataframe_to_excel(df, os.path.join(path_to_save_data, 'ambulant_diagnose_list.xlsx'))
    except Exception as e:
        print(f"Failed to extract Ambulant Diagnose List data: {e} {path_to_save_data}")
