"""
Author: Rayla Kurosaki

File: phase2_print_to_file.py

Description: This file contains functions to analyze and print the data to a
             text file.
"""

import copy

import numpy as np

import __utils__ as utils
import __utils__.miscellaneous_functions as misc


def get_data_to_print(header, data):
    """
    Gets all the necessary data to print.

    :param header: First row of the table.
    :param data: Desired data to print.
    :return: The necessary data to print.
    """
    # Compute the number of runs.
    num_runs = sum(data.values())
    # Initialize the data to print.
    data_to_print = copy.deepcopy(np.array(header))
    # Iterate through each item in the dictionary.
    for k, v in data.items():
        # Compute the value of this type of drop.
        val = misc.compute_drop_value(k[1:-1].split("/"))
        # Create a list of data to print from this drop.
        lst = np.array([
            [k, val, v, f"{utils.format_num_2(100 * v / num_runs)}%"]
        ])
        # Add the list to the end of the main list.
        data_to_print = np.vstack((data_to_print, lst))
    return data_to_print


def print_analysis(datas):
    """
    Prints out the analysis for the Talent material drops and for the Weapon
    material drops.

    :param datas: A list of data to print.
    """
    # Initialize the header for each section.
    sxn_headers = ["Talent Domain Drop Analysis",
                   "Weapon Domain Drop Analysis"]
    # Initialize the first row of the tables.
    table_header = [["Drop", "Value", "Occurrence", "Probability"]]
    # Write onto a file.
    with open("../data/talent_weapon_drop_analysis.txt", "w") as f:
        # For each data set.
        for i, (data_i, sub_header) in enumerate(zip(datas, sxn_headers)):
            # Get the data to print.
            data = get_data_to_print(table_header, data_i)
            # Get the maximum length of each column.
            max_len = utils.get_max_len(data)
            # Get the spacing of the section header.
            spacing = utils.get_spacing(max_len, sub_header)
            # Write the header onto the file.
            f.write(f"{'=' * spacing} {sub_header} {'=' * spacing}\n")
            # Print the boundary of the table.
            utils.print_boundary(f, max_len)
            # For each row in the data to print.
            for j, row in enumerate(data):
                # If the row is the fist non-header row.
                if j == 1:
                    # Print the line that separates the header with the
                    # analyzed data.
                    utils.print_separator(f, max_len)
                    pass
                # Print the row.
                utils.print_row(f, row, max_len)
                pass
            # Print the boundary of the table.
            utils.print_boundary(f, max_len)
            # If this is not the last set of data to print.
            if i + 1 < len(datas):
                # Print some new lines to separate the tables.
                f.write("\n\n")
            pass
        pass
    pass


def phase2_main(talent, weapon):
    """
    This file calls functions to analyze the Talent and Weapon material data
    sets and prints it out to a text file.

    :param talent: The Talent data set to analyze.
    :param weapon: The Weapon data set to analyze.
    """
    print_analysis([talent, weapon])
    pass
