import os
import time
import pandas as pd
from master.global_functions import move_mouse_on, select_and_copy_data_from_table, save_data_to_excel, \
    write_dataframe_to_excel

def extract_table_data():
    """
    This function extracts the table data from the screen.

    :return: The table data.
    :rtype: str
    """
    return select_and_copy_data_from_table('images/diagnose_list/02-top-left-corner.jpg',
                                            'images/diagnose_list/03-top-right-corner.jpg',
                                            'images/diagnose_list/04-bottom-right-corner.jpg',
                                            confidence=0.95,
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


def main_extract_diagnose_list_data(path_to_save):
    """
    This function extracts data from the Ambulant Diagnose List and saves it to an Excel file.

    :param path_to_save: The path to the directory where the resulting Excel file should be saved.
    :type path_to_save: str
    """
    move_mouse_on('images/diagnose_list/01-diagnoseoverblik.jpg')
    time.sleep(2)

    table_data = extract_table_data()
    rows = split_table_data_into_rows(table_data)
    rows = remove_noteret_rows(rows)
    pairs = split_rows_into_pairs(rows)
    df = create_dataframe_from_pairs(pairs)

    write_dataframe_to_excel(df, os.path.join(path_to_save, 'ambulant_diagnose_list.xlsx'))


