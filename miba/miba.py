import os
import time
from master.global_functions import click_by_mouse_on, save_data_to_excel, select_and_copy_data_from_table


def extract_miba_all_data():
    """
    Extracts all MIBA data from an application window and saves it to an Excel file.
    """



    # Extract data from the table
    miba_data = select_and_copy_data_from_table('images/miba/05-up-left-corner.jpg',
                                                'images/miba/06-up-right-corner.jpg',
                                                'images/miba/07-scroll-down-until-sign-found.jpg', 800, 0, 100)
    return miba_data
